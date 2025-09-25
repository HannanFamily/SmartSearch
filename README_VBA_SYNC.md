# VBA Synchronization Workflow
# ============================
# Quick reference for rebuilding Excel VBA from updated code files

## Overview
This project now includes automated tools to synchronize VBA code from `.bas` and `.cls` files back into Excel workbooks. This is essential for your development workflow where you edit VBA code in external editors and need to import changes back into Excel.

## Available Scripts

### 1. Python Script (Cross-platform)
```bash
python3 sync_vba_to_excel.py [workbook_name]
```
- **Best for**: General use, cross-platform compatibility
- **Dependencies**: Python 3 (built-in libraries only)

### 2. PowerShell Script (Windows)
```powershell
.\sync_vba_to_excel.ps1 [workbook_name]
```
- **Best for**: Windows environments, native Excel integration
- **Dependencies**: PowerShell (built into Windows)

### 3. Bash Script (Linux/macOS)
```bash
./sync_vba_to_excel.sh [workbook_name]
```
- **Best for**: Linux/macOS environments
- **Dependencies**: Bash shell (standard on Linux/macOS)

## What These Scripts Do

1. **Find your Excel workbook** (or let you choose if multiple exist)
2. **Create a timestamped backup** in `Old_Code/` directory
3. **Generate an import script** (`import_vba_modules.bas`) that can be run in Excel
4. **Create detailed instructions** (`VBA_SYNC_INSTRUCTIONS.txt`)
5. **List all VBA files** found and their status

## How to Use

### Step 1: Run the sync script
Choose your preferred script and run it:
```bash
# Python version
python3 sync_vba_to_excel.py

# PowerShell version (Windows)
.\sync_vba_to_excel.ps1

# Bash version (Linux)
./sync_vba_to_excel.sh
```

### Step 2: Open Excel and import
1. Open your target Excel workbook
2. Press `Alt+F11` to open VBA Editor
3. **Automatic method**: Import and run `import_vba_modules.bas`
4. **Manual method**: Import each `.bas/.cls` file individually

### Step 3: Verify and save
1. Test your updated functionality
2. Save the workbook
3. Remove the temporary import module

## Files Created

Each run creates:
- **Backup**: `Old_Code/WorkbookName_backup_YYYYMMDD_HHMMSS.xlsm`
- **Import script**: `import_vba_modules.bas` (temporary, can be deleted after use)
- **Instructions**: `VBA_SYNC_INSTRUCTIONS.txt` (detailed step-by-step guide)

## Current VBA Modules

The scripts automatically detect and handle these files:
- `Dashboard.cls` - Dashboard worksheet event handlers
- `ThisWorkbook.cls` - Workbook-level event handlers  
- `mod_PrimaryConsolidatedModule.bas` - Main search engine and utilities
- `mod_ModeDrivenSearch.bas` - Mode-driven search functionality
- `temp_mod_ConfigTableTools.bas` - Configuration table maintenance
- `export.bas` - Metadata export utilities

## Troubleshooting

### "Programmatic access to VBA not allowed"
1. In Excel: File → Options → Trust Center → Trust Center Settings
2. Macro Settings → Check "Trust access to the VBA project object model"

### Import script doesn't work
- Run the script from the project directory
- Ensure all `.bas/.cls` files exist in the same folder as the Excel workbook
- Check that macros are enabled in Excel

### Manual import preferred
- Use File → Import in VBA Editor
- Import each module individually
- Remove old modules first to avoid conflicts

## Integration with Development Workflow

1. **Edit VBA code** in your preferred editor (VS Code, Notepad++, etc.)
2. **Commit changes** to Git when satisfied
3. **Run sync script** to prepare for Excel import
4. **Import into Excel** using generated scripts
5. **Test functionality** in Excel
6. **Save workbook** to preserve changes

## Best Practices

- Always run the sync script after editing VBA files
- Keep backups (automatically created in `Old_Code/`)
- Test thoroughly after importing
- Use version control (Git) for VBA source files
- Document significant changes in commit messages

## Notes

- VBA files (`.bas/.cls`) are the source of truth
- Excel workbook serves as the runtime environment
- Backups are created automatically for safety
- Scripts work with multiple workbooks in the same directory
- Import process preserves existing worksheet and workbook structure