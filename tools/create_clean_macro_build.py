#!/usr/bin/env python3
import os
import sys
from datetime import datetime
from typing import List, Tuple

try:
    import win32com.client  # type: ignore
    import pythoncom  # type: ignore
except ImportError:
    print("ERROR: pywin32 not installed. Install with: pip install pywin32")
    sys.exit(1)


DEV_PREFIXES = (
    'Dev_', 'temp_mod_',
)

DEV_EXACT = {
    'ActiveModuleImporter',
    'MasterVBAImporter',
    'PythonVBAConverter',
    'DevEnvironmentAnalyzer',
    'QuickDevAnalysis',
    'Dev_Exports',
    'Dev_ModuleCatalog',
    'Dev_ControlCenter',
    'Dev_SmokeTests',
    'SyncManager',
    'FileSystemManager',
    'SootblowerFormFactory',  # deprecated
}

TEST_FORMS = {
    'UserForm1', 'MyDynamicForm'
}

# Runtime modules to keep even if they match a prefix accidentally
RUNTIME_WHITELIST = {
    'mod_PrimaryConsolidatedModule',
    'mod_ModeDrivenSearch',
    'mod_UniversalTools',
    'DataTableUpdater',
    'ModeConfigBootstrap',
    'mod_SSB_RuntimeBinder',
    'C_SSB_BtnHandler',
    'SootblowerFormCreator',
    'frmSootblowerLocator',  # form
}


def should_remove(name: str, kind: int) -> bool:
    base = name.strip()
    # Never remove whitelisted runtime pieces
    if base in RUNTIME_WHITELIST:
        return False
    # Remove obvious dev/test forms
    if base in TEST_FORMS:
        return True
    # Remove exact dev names
    if base in DEV_EXACT:
        return True
    # Remove by dev prefixes
    if any(base.startswith(p) for p in DEV_PREFIXES):
        return True
    return False


def clean_macro_copy(src_workbook: str, out_dir: str) -> Tuple[str, List[str], List[str]]:
    """Create a cleaned .xlsm copy with dev/test modules removed.
    Returns (path, removed, kept).
    """
    pythoncom.CoInitialize()
    app = None
    removed: List[str] = []
    kept: List[str] = []
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Open(os.path.abspath(src_workbook))

        # Save an initial copy under out_dir
        os.makedirs(out_dir, exist_ok=True)
        dest = os.path.join(out_dir, 'Search Dashboard v1.3 - CleanMacro.xlsm')
        wb.SaveCopyAs(os.path.abspath(dest))
        wb.Close(SaveChanges=False)

        # Re-open the copy for VBComponent edits
        wbc = app.Workbooks.Open(os.path.abspath(dest))
        try:
            vbproj = wbc.VBProject  # Requires Trust access to VBOM
        except Exception:
            # If trust is not enabled, we still deliver the copy
            wbc.Close(SaveChanges=True)
            return dest, removed, kept

        # Iterate and remove dev components
        components = list(vbproj.VBComponents)
        # Collect names first to avoid invalidation during iteration
        names_kinds = [(str(c.Name), int(c.Type)) for c in components]
        for name, kind in names_kinds:
            if should_remove(name, kind):
                try:
                    vbproj.VBComponents.Remove(vbproj.VBComponents(name))
                    removed.append(name)
                except Exception:
                    # If removal fails, keep it
                    kept.append(name)
            else:
                kept.append(name)

        wbc.Save()
        wbc.Close(SaveChanges=True)
        return dest, removed, kept
    finally:
        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def main():
    repo_root = os.path.dirname(os.path.abspath(__file__))
    repo_root = os.path.dirname(repo_root)  # from tools/
    workbook = os.path.join(repo_root, 'Search Dashboard v1.3.xlsm')
    if len(sys.argv) > 1:
        workbook = sys.argv[1]
    if not os.path.isfile(workbook):
        print(f"ERROR: Workbook not found: {workbook}")
        sys.exit(2)

    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    out_dir = os.path.join(repo_root, 'Clean_Macro_Builds', f'CleanMacro_{ts}')
    path, removed, kept = clean_macro_copy(workbook, out_dir)

    # Write a simple build report
    report = os.path.join(out_dir, 'Build_Report.txt')
    with open(report, 'w', encoding='utf-8') as f:
        f.write('Clean Macro Build Report\n')
        f.write('=========================\n\n')
        f.write(f'Workbook: {path}\n\n')
        f.write('Removed components:\n')
        for n in sorted(set(removed)):
            f.write(f'  - {n}\n')
        f.write('\nKept components:\n')
        for n in sorted(set(kept)):
            f.write(f'  - {n}\n')
        if not removed:
            f.write('\nNote: Trust access to the VBA project may be disabled, so no components were removed in this build.\n')

    print('Clean macro build created at: ', out_dir)
    print('Workbook: ', path)
    if removed:
        print('Removed components count: ', len(removed))
    else:
        print('No components removed (likely VBOM trust disabled).')


if __name__ == '__main__':
    main()
