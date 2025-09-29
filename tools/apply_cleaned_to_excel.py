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


def load_mapping_from_cleaned(cleaned_csv: str) -> Tuple[Dict[str, str], int]:
    mapping: Dict[str, str] = {}
    total = 0
    with open(cleaned_csv, 'r', encoding='utf-8', newline='') as f:
        reader = csv.DictReader(f)
        fields = [c.lower() for c in reader.fieldnames or []]
        # columns
        id_field = 'sap equipment id'
        if id_field not in fields and 'equipment number' in fields:
            id_field = 'equipment number'
        desc_field = 'equipment description'
        for row in reader:
            rid = (row.get(id_field) or '').strip()
            if rid:
                mapping[rid] = row.get(desc_field) or ''
                total += 1
    return mapping, total


def apply_to_excel(workbook_path: str, mapping: Dict[str, str], visible: bool = False, debug: bool = False) -> Tuple[int, int]:
    """Apply Equipment Description updates by matching Equipment ID in the DataTable listobject."""
    pythoncom.CoInitialize()
    app = None
    try:
        try:
            app = win32com.client.GetActiveObject("Excel.Application")
            if debug:
                print("Connected to existing Excel instance")
        except Exception:
            app = win32com.client.Dispatch("Excel.Application")
            if debug:
                print("Created new Excel instance")
        try:
            app.AutomationSecurity = 1
        except Exception:
            pass
        app.Visible = visible
        app.DisplayAlerts = False

        wb = app.Workbooks.Open(os.path.abspath(workbook_path))
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

        id_col = hidx('SAP Equipment ID') or hidx('Equipment Number')
        desc_col = hidx('Equipment Description')
        if id_col == 0 or desc_col == 0:
            raise RuntimeError("Required columns not found: SAP Equipment ID/Equipment Number and Equipment Description")

        rng = data_lo.DataBodyRange
        values = rng.Value  # 2D tuple (rows x cols in table)
        if not isinstance(values, tuple):
            values = ((values,),)

        total_rows = len(values)
        updated = 0
        # Convert to list of lists for mutability
        new_values = [list(row) for row in values]

        for r in range(total_rows):
            rid = str(new_values[r][id_col - 1]).strip()
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
