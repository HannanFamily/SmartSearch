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

DATA_DIR = os.path.join(os.path.dirname(__file__), '..', 'Data_Cleanup')
OUTPUT_DIR = os.path.join(DATA_DIR, 'output')

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

    # Step 1: fix spaced letters
    fixed, ns = fix_spaced_letters(cur)
    if fixed != cur:
        cur = fixed
        notes += ns
        conf = max(conf, 85)

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


def main():
    cfg = load_config(CONFIG_FILE)
    mapping = learn_from_sample(SAMPLE_FILE)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    # Process main equipment data
    out_rows = process_file(MAIN_FILE, mapping, cfg)
    out_file = os.path.join(OUTPUT_DIR, f'Equipment_Data_Cleanup_Suggestions_{timestamp}.csv')
    write_csv(out_rows, out_file)

    # Changes-only CSV
    changes_only = [r for r in out_rows if (r.get('Suggested Edit') or '').strip() or (r.get('Object Type (Suggested)') or '').strip()]
    changes_file = os.path.join(OUTPUT_DIR, f'Equipment_Data_Cleanup_ChangesOnly_{timestamp}.csv')
    write_csv(changes_only, changes_file)

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

    print(f"Wrote suggestions: {out_file}")
    print(f"Wrote changes-only: {changes_file}")
    print(f"Wrote summary: {summary_file}")

if __name__ == '__main__':
    main()
