#!/usr/bin/env python3
import csv
import os
import re
import io
import chardet
from datetime import datetime
from typing import List, Dict, Tuple

DATA_DIR = os.path.join(os.path.dirname(__file__), '..', 'Data_Cleanup')
OUTPUT_DIR = os.path.join(DATA_DIR, 'output')

SAMPLE_FILE = os.path.join(DATA_DIR, 'Sample Original Data.csv')
MAIN_FILE = os.path.join(DATA_DIR, 'Equipment Data.csv')

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
        (r"\bPres\s*s\b", "Press"),
        (r"\bPres\s*sure\b", "Pressure"),
        (r"\bIndicat\s*ing\b", "Indicating"),
        (r"\bCabin\s*et\b", "Cabinet"),
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


def clean_description(desc: str, obj_type: str, mapping: Dict[str, str]) -> Tuple[str, int, List[str]]:
    notes: List[str] = []
    original = desc or ''
    cur = original

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

    # Confidence heuristic
    if cur == original:
        # No change
        return '', 100, ['No change needed']

    # For trivial casing/space-only changes, push confidence up
    if original.replace(' ', '').lower() == cur.replace(' ', '').lower():
        conf = max(conf, 95)

    conf = min(max(conf, 75), 100)
    return cur, conf, notes


def process_file(input_path: str, mapping: Dict[str, str]) -> List[Dict[str, str]]:
    rows_out: List[Dict[str, str]] = []
    text = _read_csv_text(input_path)
    reader = csv.DictReader(io.StringIO(text))
    for row in reader:
        original = row.get('Equipment Description') or ''
        obj_type = row.get('Object Type') or ''
        suggestion, conf, notes = clean_description(original, obj_type, mapping)
        rows_out.append({
            'Equipment Number': row.get('SAP ID') or row.get('Equipment Number') or '',
            'Original Description': original,
            'Suggested Edit': suggestion,
            'Confidence': str(conf),
            'Reasoning': '; '.join(notes)
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


def main():
    mapping = learn_from_sample(SAMPLE_FILE)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')

    # Process main equipment data
    out_rows = process_file(MAIN_FILE, mapping)
    out_file = os.path.join(OUTPUT_DIR, f'Equipment_Data_Cleanup_Suggestions_{timestamp}.csv')
    write_csv(out_rows, out_file)

    print(f"Wrote suggestions: {out_file}")

if __name__ == '__main__':
    main()
