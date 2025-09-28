# One-click project export

Use these macros to produce a timestamped, reproducible snapshot under `logs/Project_Export_YYYYMMDD_HHMMSS/`:

- `RUN_Test_And_Export` (recommended)
  1. Syncs ActiveModules into the workbook
  2. Runs `RUN_SmokeTest_Workbook`
  3. Exports:
     - DataTable.csv, ConfigTable.csv, ModeConfigTable.csv (to Data_Exports/)
     - VBA modules (to ActiveModules/)
     - Environment info (env.txt)

- `RUN_Export_ProjectSnapshot`
  - Runs only the export portion

Optional convenience
- Run `RUN_Dev_AddDashboardButtons` to add two buttons (Test + Export, Export Snapshot) on the `Dashboard` sheet.

Notes
- Exports land in the workbook folder under `logs/`.
- If Trust access to VBOM is disabled, env.txt will still be created, and table CSVs will export; module export requires VBOM.
