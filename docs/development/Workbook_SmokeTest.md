# Workbook smoke test

Use the macro RUN_SmokeTest_Workbook to validate core dependencies:
- DataTable exists and has data
- Required headers are present
- ModeConfig table exists and is non-empty
- OutputAllVisible runs without error

Steps
1) Open the workbook in Excel
2) Alt+F8 > Run: RUN_SmokeTest_Workbook
3) Expect a "Smoke test passed" message. If it fails, the dialog will show what's missing

Notes
- The test calls OutputAllVisible which depends on ConfigTable keys and the active mode; ensure config is valid
- For convenience, you can wire this macro to a Dashboard button
