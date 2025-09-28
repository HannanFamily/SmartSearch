# Changelog - Project Cleanup and Terminal Control

## 2025-09-28

- Added robust Python path discovery to `tools/excel_vba.ps1`:
  - Auto-detects Python across common install locations, respects `PYTHON_PATH`
  - Supports `-Hidden` and `-Visible` switches explicitly
  - Removes hard-coded username dependency
- Created `.gitignore` tailored for Excel/Python project:
  - Ignores temp files, logs, outputs, and exported module snapshots
- Verified terminal control via `vba_dashboard.bat` / `vba_interactive.bat`
- Documented capabilities in `TERMINAL_CONTROL_COMPLETE.md` and `tools/README_VBA_Terminal.md`

Next: consider adding a `tools\tasks.ps1` for common scripted operations (export snapshot, run diagnostics, sync modules), and wire buttons in Excel to call them if desired.
