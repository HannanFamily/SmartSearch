#!/usr/bin/env python3
import argparse
import csv
import os
import sys
from datetime import datetime
from typing import Dict, Tuple, Optional

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
except ImportError:
    print("ERROR: pywin32 not installed. Install with: pip install pywin32")
    sys.exit(1)


def find_latest_cleaned(output_dir: str) -> Optional[str]:
    if not os.path.isdir(output_dir):
        return None
    prefix = "Equipment_Data_CLEANED_"
    candidates = [
        os.path.join(output_dir, f)
        for f in os.listdir(output_dir)
        if f.startswith(prefix) and f.lower().endswith('.csv')
    ]
    if not candidates:
        return None
    candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidates[0]


def _norm_id(val: object) -> str:
    s = '' if val is None else str(val).strip()
    # If looks like a float with .0, strip it
    if s.endswith('.0') and s.replace('.', '', 1).isdigit():
        try:
            s = str(int(float(s)))
        except Exception:
            pass
    # Prefer digit-only normalization when present
    digits = ''.join(ch for ch in s if ch.isdigit())
    return digits if digits else s


def load_mapping_from_cleaned(cleaned_csv: str, id_col_override: Optional[str] = None, desc_col_override: Optional[str] = None) -> Tuple[Dict[str, str], int]:
    mapping: Dict[str, str] = {}
    total = 0
    with open(cleaned_csv, 'r', encoding='utf-8', newline='') as f:
        reader = csv.DictReader(f)
        orig_fields = reader.fieldnames or []
        fields = [c.lower() for c in orig_fields]
        # columns (prefer explicit overrides)
        id_candidates = [
            (id_col_override or '').lower(),
            'sap equipment id', 'sap id', 'equipment number', 'equipment no', 'equipment #'
        ]
        id_field = next((f for f in id_candidates if f and f in fields), None)
        if not id_field:
            # fallback: try contains checks
            id_field = next((f for f in fields if 'sap' in f and 'id' in f), None)
        desc_candidates = [
            (desc_col_override or '').lower(),
            'equipment description', 'description'
        ]
        desc_field = next((f for f in desc_candidates if f and f in fields), None)
        if not desc_field:
            desc_field = next((f for f in fields if 'description' in f), None)
        for row in reader:
            # Access row values case-insensitively
            lower_row = { (k or '').lower(): (v or '') for k, v in row.items() }
            rid = (lower_row.get(id_field or '') or '').strip()
            if rid:
                mapping[_norm_id(rid)] = (lower_row.get(desc_field or '') or '') if desc_field else ''
                total += 1
    return mapping, total


def apply_to_excel(workbook_path: str, mapping: Dict[str, str], visible: bool = False, debug: bool = False) -> Tuple[int, int]:
    """Apply Equipment Description updates by matching Equipment ID in the DataTable listobject."""
    pythoncom.CoInitialize()
    app = None
    try:
        # Create a fresh Excel instance to avoid cross-process property issues
        app = win32com.client.DispatchEx("Excel.Application")
        if debug:
            print("Created new Excel instance")
        try:
            app.AutomationSecurity = 1
        except Exception:
            pass
        # Open workbook before toggling visibility to avoid some COM quirks
        wb = app.Workbooks.Open(os.path.abspath(workbook_path))
        try:
            app.Visible = bool(visible)
        except Exception:
            pass
        try:
            app.DisplayAlerts = False
        except Exception:
            pass
        if debug:
            print(f"Opened workbook: {wb.FullName}")

        # Find ListObject named DataTable
        data_lo = None
        for sheet in wb.Worksheets:
            try:
                for lo in sheet.ListObjects:
                    if str(lo.Name).strip().lower() == 'datatable':
                        data_lo = lo
                        break
            except Exception:
                continue
            if data_lo is not None:
                break
        if data_lo is None:
            raise RuntimeError("ListObject 'DataTable' not found")

        if data_lo.DataBodyRange is None:
            if debug:
                print("DataTable has no rows; nothing to update")
            return 0, 0

        # Header indices
        headers = [str(c.Value).strip() for c in data_lo.HeaderRowRange.Columns]
        def hidx(name: str) -> int:
            for i, h in enumerate(headers, start=1):
                if h.lower() == name.lower():
                    return i
            return 0

        # Flexible header resolution
        def resolve_idx(options):
            for opt in options:
                idx = hidx(opt)
                if idx:
                    return idx
            # contains fallback
            for i, h in enumerate(headers, start=1):
                hl = h.lower()
                if any(tok in hl for tok in options[0].lower().split()):
                    return i
            return 0

        id_col = resolve_idx(['SAP Equipment ID', 'SAP ID', 'Equipment Number', 'Equipment No', 'Equipment #'])
        desc_col = resolve_idx(['Equipment Description', 'Description'])
        if id_col == 0 or desc_col == 0:
            if debug:
                print("Headers detected:")
                for i, h in enumerate(headers, start=1):
                    print(f"  {i}: {h}")
            raise RuntimeError("Required columns not found: need an ID column (SAP Equipment ID/SAP ID/Equipment Number) and Equipment Description")

        rng = data_lo.DataBodyRange
        values = rng.Value  # 2D tuple (rows x cols in table)
        if not isinstance(values, tuple):
            values = ((values,),)

        total_rows = len(values)
        updated = 0
        # Convert to list of lists for mutability
        new_values = [list(row) for row in values]

        for r in range(total_rows):
            rid = _norm_id(new_values[r][id_col - 1])
            if not rid or rid == 'None':
                continue
            new_desc = mapping.get(rid)
            if new_desc is None:
                continue
            cur_desc = str(new_values[r][desc_col - 1]) if new_values[r][desc_col - 1] is not None else ''
            if cur_desc.strip() != new_desc.strip():
                new_values[r][desc_col - 1] = new_desc
                updated += 1

        if updated > 0:
            if debug:
                print(f"Writing {updated} updated descriptions back to worksheet...")
            # Write back in one shot
            rng.Value = tuple(tuple(row) for row in new_values)
            wb.Save()
        else:
            if debug:
                print("No changes detected; nothing saved")

        return total_rows, updated
    finally:
        try:
            if app is not None and not visible:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def main():
    parser = argparse.ArgumentParser(description="Apply CLEANED Equipment Descriptions to Excel DataTable")
    parser.add_argument('--workbook', '-w', required=True, help='Path to Excel workbook (xlsm)')
    parser.add_argument('--cleaned-csv', '-c', help='Path to CLEANED CSV (from data cleanup)')
    parser.add_argument('--output-dir', '-o', help='Directory to search for CLEANED CSV if not provided')
    parser.add_argument('--visible', action='store_true', help='Show Excel window while applying')
    parser.add_argument('--debug', action='store_true', help='Verbose output')
    args = parser.parse_args()

    cleaned = args.cleaned_csv
    if not cleaned:
        # Try provided output dir, then default tools/Data_Cleanup_Package/output, then Data_Cleanup/output
        search_dirs = []
        if args.output_dir:
            search_dirs.append(args.output_dir)
        # repo-relative options
        repo_root = os.path.dirname(os.path.abspath(args.workbook))
        search_dirs.append(os.path.join(repo_root, 'tools', 'Data_Cleanup_Package', 'output'))
        search_dirs.append(os.path.join(repo_root, 'Data_Cleanup', 'output'))
        for d in search_dirs:
            cleaned = find_latest_cleaned(d)
            if cleaned:
                break
    if not cleaned or not os.path.isfile(cleaned):
        print("ERROR: CLEANED CSV not found. Provide --cleaned-csv or ensure outputs exist.")
        sys.exit(2)

    mapping, total = load_mapping_from_cleaned(cleaned)
    if args.debug:
        print(f"Loaded {total} rows from CLEANED CSV: {cleaned}")

    rows, updated = apply_to_excel(args.workbook, mapping, visible=args.visible, debug=args.debug)
    print(f"Applied to workbook: {args.workbook}")
    print(f"DataTable rows: {rows}")
    print(f"Descriptions updated: {updated}")


if __name__ == '__main__':
    main()
