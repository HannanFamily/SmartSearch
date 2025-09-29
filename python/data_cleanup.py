#!/usr/bin/env python3
import csv
import os
import re
import io
import chardet
import yaml
from rapidfuzz import fuzz
from unidecode import unidecode
from datetime import datetime
from typing import List, Dict, Tuple

# Determine data and output directories; allow environment overrides for isolation use cases
_DEFAULT_DATA_DIR = os.path.join(os.path.dirname(__file__), '..', 'Data_Cleanup')
DATA_DIR = os.environ.get('DATA_CLEANUP_DIR') or _DEFAULT_DATA_DIR
DATA_DIR = os.path.abspath(DATA_DIR)

_DEFAULT_OUTPUT_DIR = os.path.join(DATA_DIR, 'output')
OUTPUT_DIR = os.environ.get('DATA_CLEANUP_OUTPUT_DIR') or _DEFAULT_OUTPUT_DIR
OUTPUT_DIR = os.path.abspath(OUTPUT_DIR)

SAMPLE_FILE = os.path.join(DATA_DIR, 'Sample Original Data.csv')
MAIN_FILE = os.path.join(DATA_DIR, 'Equipment Data.csv')
CONFIG_FILE = os.path.join(DATA_DIR, 'data_cleanup_config.yaml')

# Basic normalization helpers
_space_re = re.compile(r"\s+")
_multi_caps_re = re.compile(r"\b([A-Z]{2,})([a-z]+)\b")


def title_case_equipment(text: str) -> str:
    # Keep known acronyms fully upper, otherwise title-case tokens
    if not text:
        return text
    # Normalize whitespace first
    text = _space_re.sub(' ', text.strip())
    # Simple title casing per word, but keep all-caps tokens
    words = []
    for w in text.split(' '):
        if w.isupper() and len(w) <= 4:
            words.append(w)
        else:
            words.append(w.capitalize())
    return ' '.join(words)


def fix_spaced_letters(text: str) -> Tuple[str, List[str]]:
    notes = []
    # Fix common spaced letter artifacts like "Pres sure", "Indicat ing"
    patterns = [
        (r"\bPres\s*sure\b", "Pressure"),
        (r"\bPres\s*s\b", "Press"),
        (r"\bIndicat\s*ing\b", "Indicating"),
        (r"\bCabin\s*et\b", "Cabinet"),
        (r"\bCir\s*cuit\b", "Circuit"),
        (r"\bGen\s*erator\b", "Generator"),
    ]
    new = text
    for pat, repl in patterns:
        if re.search(pat, new, flags=re.IGNORECASE):
            new = re.sub(pat, repl, new, flags=re.IGNORECASE)
            notes.append(f"Fixed spaced letters: {pat} -> {repl}")
    return new, notes


def avoid_bad_expansions(text: str) -> Tuple[str, List[str]]:
    notes = []
    # Avoid expanding stray N/S/E/W to North/South/etc unless clearly directional
    # Here we only collapse double spaces and do not expand single letters
    return text, notes


def apply_regex_patterns(text: str, patterns: List[Dict[str, str]], label: str) -> Tuple[str, List[str]]:
    """Generic regex pattern applier using a list of {pattern, replacement}."""
    notes: List[str] = []
    new = text
    for rule in patterns or []:
        pat = rule.get('pattern')
        repl = rule.get('replacement')
        if not pat or repl is None:
            continue
        if re.search(pat, new, flags=re.IGNORECASE):
            new2 = re.sub(pat, repl, new, flags=re.IGNORECASE)
            if new2 != new:
                notes.append(f"{label}: {pat} -> {repl}")
                new = new2
    return new, notes


def merge_split_tokens(text: str, cfg: dict) -> Tuple[str, List[str]]:
    """Merge tokens that were erroneously split inside domain terms.
    Uses regex patterns from YAML to correct e.g., 'Va Riable' -> 'Variable'."""
    notes: List[str] = []
    patterns = cfg.get('split_merge_patterns', [])
    new = text
    for rule in patterns:
        pat = rule.get('pattern')
        repl = rule.get('replacement')
        if not pat or repl is None:
            continue
        if re.search(pat, new, flags=re.IGNORECASE):
            new = re.sub(pat, repl, new, flags=re.IGNORECASE)
            notes.append(f"Merged split token: {repl}")
    return new, notes


def correct_misspellings(text: str, cfg: dict) -> Tuple[str, List[str]]:
    """Fix common misspellings (post-merge) using explicit rules and light fuzzy match within words."""
    notes: List[str] = []
    rules = cfg.get('misspellings', [])
    new = text
    for r in rules:
        wrong = r.get('wrong')
        right = r.get('right')
        if not wrong or right is None:
            continue
        pattern = re.compile(rf"\b{re.escape(wrong)}\b", flags=re.IGNORECASE)
        if pattern.search(new):
            new = pattern.sub(right, new)
            notes.append(f"Fixed misspelling: {wrong} -> {right}")
    return new, notes


def load_config(path: str) -> dict:
    if not os.path.exists(path):
        return {}
    with open(path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f) or {}


def enforce_acronyms_and_protected(text: str, cfg: dict) -> Tuple[str, List[str]]:
    notes: List[str] = []
    if not text:
        return text, notes
    tokens = text.split(' ')
    acronyms = set([a.lower() for a in cfg.get('acronyms_upper', [])])
    protected = cfg.get('protected_words', [])

    # Force acronyms uppercase by token
    for i, t in enumerate(tokens):
        if t.lower() in acronyms:
            if tokens[i] != t.upper():
                tokens[i] = t.upper()
                notes.append(f"Forced acronym uppercase: {tokens[i]}")
    text2 = ' '.join(tokens)

    # Enforce protected phrases casing (case-insensitive replace)
    for phrase in protected:
        pattern = re.compile(re.escape(phrase), flags=re.IGNORECASE)
        if pattern.search(text2):
            text2 = pattern.sub(phrase, text2)
            notes.append(f"Enforced protected phrase casing: {phrase}")
    return text2, notes


def maybe_expand_directionals(text: str, cfg: dict) -> Tuple[str, List[str]]:
    notes: List[str] = []
    words = text.split()
    ctx = set(cfg.get('directional_context_keywords', []))
    exp = cfg.get('directional_expansions', {})
    changed = False
    for i, w in enumerate(words):
        up = w.upper()
        if len(w) == 1 and up in exp:
            neighbors = [words[j].lower() for j in range(max(0, i-2), min(len(words), i+3)) if j != i]
            if any(n in ctx for n in neighbors):
                words[i] = exp[up]
                changed = True
    if changed:
        notes.append('Expanded directional based on contextual keywords')
    return ' '.join(words), notes


def _read_csv_text(path: str) -> str:
    # Detect encoding with chardet, then decode
    with open(path, 'rb') as fb:
        raw = fb.read()
    guess = chardet.detect(raw) or {}
    enc = guess.get('encoding') or 'utf-8'
    try:
        text = raw.decode(enc)
    except UnicodeDecodeError:
        # Fallback to common Excel encodings
        for enc2 in ('utf-8-sig', 'cp1252', 'latin-1'):
            try:
                text = raw.decode(enc2)
                break
            except UnicodeDecodeError:
                continue
        else:
            text = raw.decode('latin-1', errors='ignore')
    # Strip BOM if present
    if text and text[0] == '\ufeff':
        text = text[1:]
    return text


def learn_from_sample(sample_path: str) -> Dict[str, str]:
    mapping = {}
    text = _read_csv_text(sample_path)
    reader = csv.DictReader(io.StringIO(text))
    for row in reader:
        orig = (row.get('Original Description') or '').strip()
        upd = (row.get('Updated Description') or '').strip()
        if orig and upd and orig.lower() != upd.lower():
            mapping[orig.lower()] = upd
    return mapping


def apply_learned_mapping(text: str, mapping: Dict[str, str]) -> Tuple[str, List[str]]:
    notes = []
    key = text.strip().lower()
    if key in mapping:
        notes.append('Applied learned mapping from sample set')
        return mapping[key], notes
    return text, notes


def clean_description(desc: str, obj_type: str, mapping: Dict[str, str], cfg: dict) -> Tuple[str, int, List[str]]:
    notes: List[str] = []
    original = desc or ''
    cur = unidecode(original)

    # Step 0: learned direct mapping
    new, ns = apply_learned_mapping(cur, mapping)
    if new != cur:
        cur = new
        notes += ns
        # High confidence if direct mapping
        conf = 95
    else:
        conf = 0

    # Step 1: merge split tokens from YAML patterns
    merged, ns = merge_split_tokens(cur, cfg)
    if merged != cur:
        cur = merged
        notes += ns
        conf = max(conf, 88)

    # Step 1b: object-context merges (e.g., Heat Er -> Heater if type Heater)
    if obj_type:
        key = (obj_type or '').strip().upper()
        mapped = cfg.get('object_type_map', {}).get(key)
        label = mapped or obj_type
        ctx_rules = (cfg.get('object_context_merge_patterns', {}) or {}).get(label, [])
        if ctx_rules:
            cur2, ns = apply_regex_patterns(cur, ctx_rules, f'ContextMerge[{label}]')
            if cur2 != cur:
                cur = cur2
                notes += ns
                conf = max(conf, 90)

    # Step 1c: fix spaced letters (generic regexes)
    fixed, ns = fix_spaced_letters(cur)
    if fixed != cur:
        cur = fixed
        notes += ns
        conf = max(conf, 85)

    # Step 1d: correct misspellings
    corrected, ns = correct_misspellings(cur, cfg)
    if corrected != cur:
        cur = corrected
        notes += ns
        conf = max(conf, 90)

    # Step 2: normalize whitespace and title case
    norm = title_case_equipment(cur)
    if norm != cur:
        notes.append('Normalized whitespace/title case')
        cur = norm
        conf = max(conf, 90 if conf >= 85 else 80)

    # Step 3: avoid dangerous expansions (placeholder)
    cur, ns = avoid_bad_expansions(cur)
    notes += ns

    # Step 4: enforce acronyms/protected phrases
    cur2, ns = enforce_acronyms_and_protected(cur, cfg)
    if cur2 != cur:
        cur = cur2
        notes += ns
        conf = max(conf, 95)

    # Step 5: context-aware directional expansions
    cur2, ns = maybe_expand_directionals(cur, cfg)
    if cur2 != cur:
        cur = cur2
        notes += ns
        conf = max(conf, 80)

    # Confidence heuristic
    if cur == original:
        # No change
        return '', 100, ['No change needed']

    # For trivial casing/space-only changes, push confidence up
    if original.replace(' ', '').lower() == cur.replace(' ', '').lower():
        conf = max(conf, 95)

    conf = min(max(conf, 75), 100)
    return cur, conf, notes


def process_file(input_path: str, mapping: Dict[str, str], cfg: dict) -> List[Dict[str, str]]:
    rows_out: List[Dict[str, str]] = []
    text = _read_csv_text(input_path)
    reader = csv.DictReader(io.StringIO(text))
    for row in reader:
        # Only use the required columns
        original = row.get('Equipment Description') or ''
        obj_type = row.get('Object Type') or ''
        func_cat = row.get('Functional System Category') or ''
        func_sys = row.get('Functional System') or ''

        suggestion, conf, notes = clean_description(original, obj_type, mapping, cfg)

        # Object Type mapping (confident only)
        obj_type_suggest = ''
        key = (obj_type or '').strip().upper()
        obj_map = cfg.get('object_type_map', {})
        if key in obj_map and obj_map[key] != obj_type:
            obj_type_suggest = obj_map[key]
            notes.append(f"Object Type mapped {key} -> {obj_type_suggest}")
            conf = max(conf, 90)

        rows_out.append({
            'Equipment Number': row.get('SAP ID') or row.get('Equipment Number') or '',
            'Original Description': original,
            'Suggested Edit': suggestion,
            'Confidence': str(conf),
            'Reasoning': '; '.join(notes),
            'Object Type (Original)': obj_type,
            'Object Type (Suggested)': obj_type_suggest,
            'Functional System Category': func_cat,
            'Functional System': func_sys,
        })
    return rows_out


def _tokenize(text: str) -> List[str]:
    text = unidecode(text or '')
    return re.findall(r"[A-Za-z0-9]+", text)


def analyze_vocab(input_path: str, cfg: dict) -> Tuple[Dict[str, int], Dict[Tuple[str, str], int], List[Dict[str, str]]]:
    """Build token and bigram frequencies using ORIGINAL descriptions and collect anomalies."""
    token_freq: Dict[str, int] = {}
    bigram_freq: Dict[Tuple[str, str], int] = {}
    anomalies: List[Dict[str, str]] = []
    text = _read_csv_text(input_path)
    reader = csv.DictReader(io.StringIO(text))
    rows = list(reader)

    settings = (cfg.get('anomaly_settings') or {})
    short_max = int(settings.get('suspect_short_token_max_len', 2))
    watch = set((settings.get('watch_tokens') or []))
    allowed_short = set([t.lower() for t in (cfg.get('acronyms_upper') or [])] + ['of', 'to', 'in', 'on', 'by', 'at', 'id'])

    # Frequency counts
    for row in rows:
        orig = row.get('Equipment Description') or ''
        toks = [t.lower() for t in _tokenize(orig)]
        for t in toks:
            token_freq[t] = token_freq.get(t, 0) + 1
        for i in range(len(toks) - 1):
            bg = (toks[i], toks[i+1])
            bigram_freq[bg] = bigram_freq.get(bg, 0) + 1

    # Anomaly detection per row
    for idx, row in enumerate(rows, start=1):
        orig = row.get('Equipment Description') or ''
        toks = _tokenize(orig)
        suspects: List[str] = []
        for t in toks:
            tl = t.lower()
            if tl.isdigit():
                continue
            if len(t) <= short_max and tl not in allowed_short:
                # Short odd tokens like Er, Nd
                if re.match(r"^[A-Za-z]{1,2}$", t):
                    suspects.append(t)
            if token_freq.get(tl, 0) == 1 and len(tl) > 2:
                suspects.append(t)
            if tl in watch:
                suspects.append(t)
        if suspects:
            anomalies.append({
                'Row': str(idx),
                'Equipment Number': row.get('SAP ID') or row.get('Equipment Number') or '',
                'Original Description': orig,
                'Suspect Tokens': ', '.join(sorted(set(suspects), key=str.lower))
            })

    return token_freq, bigram_freq, anomalies


def write_csv(rows: List[Dict[str, str]], out_path: str):
    if not rows:
        return
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    fieldnames = list(rows[0].keys())
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in rows:
            writer.writerow(r)


def write_text(text: str, out_path: str):
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(text)


def write_freq_csv(freq: Dict, out_path: str):
    if not freq:
        return
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    # Make rows
    if isinstance(next(iter(freq.keys())), tuple):
        rows = [{'Token1': k[0], 'Token2': k[1], 'Count': v} for k, v in sorted(freq.items(), key=lambda kv: kv[1], reverse=True)]
    else:
        rows = [{'Token': k, 'Count': v} for k, v in sorted(freq.items(), key=lambda kv: kv[1], reverse=True)]
    with open(out_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def main():
    cfg = load_config(CONFIG_FILE)
    mapping = learn_from_sample(SAMPLE_FILE)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    # Build vocab and anomalies from ORIGINAL descriptions
    tok_freq, bigram_freq, anomalies = analyze_vocab(MAIN_FILE, cfg)
    tokens_file = os.path.join(OUTPUT_DIR, f'Equipment_Data_Tokens_{timestamp}.csv')
    bigrams_file = os.path.join(OUTPUT_DIR, f'Equipment_Data_Bigrams_{timestamp}.csv')
    anomalies_file = os.path.join(OUTPUT_DIR, f'Equipment_Data_Anomalies_{timestamp}.csv')
    write_freq_csv(tok_freq, tokens_file)
    write_freq_csv(bigram_freq, bigrams_file)
    write_csv(anomalies, anomalies_file)

    # Process main equipment data
    out_rows = process_file(MAIN_FILE, mapping, cfg)
    out_file = os.path.join(OUTPUT_DIR, f'Equipment_Data_Cleanup_Suggestions_{timestamp}.csv')
    write_csv(out_rows, out_file)

    # Changes-only CSV
    changes_only = [r for r in out_rows if (r.get('Suggested Edit') or '').strip() or (r.get('Object Type (Suggested)') or '').strip()]
    changes_file = os.path.join(OUTPUT_DIR, f'Equipment_Data_Cleanup_ChangesOnly_{timestamp}.csv')
    write_csv(changes_only, changes_file)

    # Produce CLEANED dataset (apply suggestions to Equipment Description only)
    original_text = _read_csv_text(MAIN_FILE)
    reader = csv.DictReader(io.StringIO(original_text))
    orig_rows = list(reader)
    id_keys = ['SAP ID', 'Equipment Number']
    def get_id(r: Dict[str, str]):
        for k in id_keys:
            if r.get(k):
                return r.get(k)
        return None
    sugg_by_id: Dict[str, str] = {}
    for r in out_rows:
        rid = r.get('Equipment Number')
        s = (r.get('Suggested Edit') or '').strip()
        if rid and s:
            sugg_by_id[rid] = s
    cleaned_rows: List[Dict[str, str]] = []
    for r in orig_rows:
        rid = get_id(r)
        s = sugg_by_id.get(rid)
        new_r = dict(r)
        if s:
            new_r['Equipment Description'] = s
        cleaned_rows.append(new_r)
    cleaned_file = os.path.join(OUTPUT_DIR, f'Equipment_Data_CLEANED_{timestamp}.csv')
    with open(cleaned_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=reader.fieldnames)
        writer.writeheader()
        writer.writerows(cleaned_rows)

    # Summary report
    total = len(out_rows)
    changed = len(changes_only)
    # Confidence histogram (10-point buckets)
    buckets: Dict[str, int] = {}
    for r in out_rows:
        try:
            c = int(float(r.get('Confidence') or 0))
        except Exception:
            c = 0
        bstart = (c // 10) * 10
        bend = bstart + 9
        key = f"{bstart:02d}-{bend:02d}"
        buckets[key] = buckets.get(key, 0) + 1

    # Reasoning tag counts (simple by full note string)
    reason_counts: Dict[str, int] = {}
    for r in out_rows:
        reasons = [x.strip() for x in (r.get('Reasoning') or '').split(';') if x.strip()]
        for tag in reasons:
            reason_counts[tag] = reason_counts.get(tag, 0) + 1

    lines: List[str] = []
    lines.append(f"Total rows: {total}")
    pct = (changed / total * 100.0) if total else 0.0
    lines.append(f"Changed rows: {changed} ({pct:.2f}%)")
    lines.append("")
    lines.append("Confidence histogram (by 10s):")
    for bucket in sorted(buckets.keys()):
        lines.append(f"  {bucket}: {buckets[bucket]}")
    lines.append("")
    lines.append("Top reasoning tags:")
    for tag, count in sorted(reason_counts.items(), key=lambda kv: kv[1], reverse=True)[:15]:
        lines.append(f"  {tag}: {count}")

    summary_file = os.path.join(OUTPUT_DIR, f'Equipment_Data_Cleanup_Summary_{timestamp}.txt')
    write_text('\n'.join(lines), summary_file)

    print(f"Wrote tokens: {tokens_file}")
    print(f"Wrote bigrams: {bigrams_file}")
    print(f"Wrote anomalies: {anomalies_file}")
    print(f"Wrote suggestions: {out_file}")
    print(f"Wrote changes-only: {changes_file}")
    print(f"Wrote CLEANED: {cleaned_file}")
    print(f"Wrote summary: {summary_file}")

if __name__ == '__main__':
    main()
