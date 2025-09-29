#!/usr/bin/env python3
import os
import sys
import shutil
import glob
from datetime import datetime
from typing import Optional, Tuple

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
except ImportError:
    print("ERROR: pywin32 not installed. Install with: pip install pywin32")
    sys.exit(1)


def _find_latest(pattern: str) -> Optional[str]:
    files = glob.glob(pattern)
    if not files:
        return None
    files.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return files[0]


def _find_latest_in_dir(dir_path: str, prefix: str) -> Optional[str]:
    if not os.path.isdir(dir_path):
        return None
    candidates = [os.path.join(dir_path, f) for f in os.listdir(dir_path) if f.startswith(prefix)]
    if not candidates:
        return None
    candidates.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    return candidates[0]


def locate_datatable(wb) -> Tuple[Optional[str], Optional[str]]:
    """Return (sheet_name, table_name) for DataTable if found."""
    for sheet in wb.Worksheets:
        try:
            for lo in sheet.ListObjects:
                name = str(lo.Name).strip()
                if name.lower() == 'datatable':
                    return str(sheet.Name), name
        except Exception:
            continue
    return None, None


def save_macro_free_copy(src_workbook: str, dest_xlsx: str) -> Tuple[str, Optional[str]]:
    """Open workbook and save a macro-free .xlsx copy; return (xlsx_path, datatable_sheet)."""
    pythoncom.CoInitialize()
    app = None
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Open(os.path.abspath(src_workbook))
        sheet_name, _ = locate_datatable(wb)
        # 51 = xlOpenXMLWorkbook (xlsx)
        wb.SaveAs(os.path.abspath(dest_xlsx), FileFormat=51)
        wb.Close(SaveChanges=False)
        return dest_xlsx, sheet_name
    finally:
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def main():
    repo_root = os.path.dirname(os.path.abspath(__file__))
    repo_root = os.path.dirname(repo_root)  # up from tools/
    workbook = os.path.join(repo_root, 'Search Dashboard v1.3.xlsm')
    if len(sys.argv) > 1:
        workbook = sys.argv[1]
    if not os.path.isfile(workbook):
        print(f"ERROR: Workbook not found: {workbook}")
        sys.exit(2)

    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    out_root = os.path.join(repo_root, 'Refined_Exports', f'Refined_Export_{ts}')
    os.makedirs(out_root, exist_ok=True)

    # 1) Save macro-free .xlsx copy
    refined_xlsx = os.path.join(out_root, 'Search Dashboard v1.3 - Refined.xlsx')
    xlsx_path, dt_sheet = save_macro_free_copy(workbook, refined_xlsx)

    # 2) Copy change log and CLEANED outputs from isolated package, if present
    pkg_out = os.path.join(repo_root, 'tools', 'Data_Cleanup_Package', 'output')
    changes = _find_latest_in_dir(pkg_out, 'Equipment_Data_Cleanup_ChangesOnly_')
    cleaned = _find_latest_in_dir(pkg_out, 'Equipment_Data_CLEANED_')
    summary = _find_latest_in_dir(pkg_out, 'Equipment_Data_Cleanup_Summary_')
    if changes:
        shutil.copy2(changes, os.path.join(out_root, 'Change_Log.csv'))
    if cleaned:
        shutil.copy2(cleaned, os.path.join(out_root, 'CLEANED_Equipment_Data.csv'))
    if summary:
        shutil.copy2(summary, os.path.join(out_root, 'Cleanup_Summary.txt'))

    # 3) Copy rule base (YAML)
    cfg_yaml = os.path.join(repo_root, 'tools', 'Data_Cleanup_Package', 'config', 'data_cleanup_config.yaml')
    if os.path.isfile(cfg_yaml):
        shutil.copy2(cfg_yaml, os.path.join(out_root, 'Rule_Base.yaml'))

    # 4) Write a README explaining contents and where changes were applied
    readme_path = os.path.join(out_root, 'README.txt')
    with open(readme_path, 'w', encoding='utf-8') as f:
        f.write("Refined Export\n")
        f.write("==============\n\n")
        f.write("This folder contains a macro-free copy of the Search Dashboard workbook,\n")
        f.write("ready for environments where macros/VBA cannot run.\n\n")
        f.write("Included files:\n")
        f.write("- Search Dashboard v1.3 - Refined.xlsx: Macro-free copy with updated Equipment Descriptions.\n")
        f.write("- Change_Log.csv: The rows where a suggested description change exists.\n")
        f.write("- CLEANED_Equipment_Data.csv: Full dataset with suggested descriptions applied.\n")
        f.write("- Cleanup_Summary.txt: High-level counts and reasoning histogram.\n")
        f.write("- Rule_Base.yaml: The rule base used for cleanup.\n\n")
        f.write("Where were changes applied?\n")
        f.write("- In the Excel table named 'DataTable' on sheet: {}.\n".format(dt_sheet or "<not found>"))
        f.write("- Column: 'Equipment Description', matched by the Equipment ID.\n\n")
        f.write("Notes:\n")
        f.write("- This .xlsx contains no macros or code modules, eliminating compile/security prompts.\n")
        f.write("- Use Change_Log.csv to review only the impacted rows.\n")

    print("Refined export written to:", out_root)
    print("Workbook:", xlsx_path)
    if dt_sheet:
        print("DataTable located on sheet:", dt_sheet)


if __name__ == '__main__':
    main()
