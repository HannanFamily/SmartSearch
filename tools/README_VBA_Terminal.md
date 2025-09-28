# Excel VBA Terminal Interface

This toolkit provides complete command-line control over Excel VBA projects, including the Search Dashboard. You can run VBA macros, manipulate UserForms, read/write data, and interact with the workbook through terminal commands.

## Quick Start

1. **Interactive Mode** (Recommended for exploration):
   ```
   .\tools\excel_vba.ps1 -Interactive
   ```

2. **Run a VBA Macro**:
   ```
   .\tools\excel_vba.ps1 -RunMacro "QuickSearchDiagnostics.RunQuickSearchDiagnostics"
   ```

3. **Show Workbook Information**:
   ```
   .\tools\excel_vba.ps1 -ShowInfo
   ```

## Installation Complete ✅

The following components are now installed and configured:
- **Python 3.12** with pywin32 for COM automation
- **Excel VBA Controller** (`tools/excel_vba_controller.py`)
- **PowerShell Wrapper** (`tools/excel_vba.ps1`)
- **Batch File** (`tools/excel_vba.bat`)
- **Additional packages**: xlwings, pandas, openpyxl

## Available Tools

### 1. Python Script (Core)
```bash
python tools/excel_vba_controller.py [options]
```

### 2. PowerShell Wrapper (Windows-friendly)
```powershell
.\tools\excel_vba.ps1 [options]
```

### 3. Batch File (Simple)
```cmd
tools\excel_vba.bat [options]
```

## Command Examples

### Basic VBA Operations
```powershell
# Run diagnostics
.\tools\excel_vba.ps1 -RunMacro "QuickSearchDiagnostics.RunQuickSearchDiagnostics"

# List all VBA modules
.\tools\excel_vba.ps1 -ListModules

# Show workbook info
.\tools\excel_vba.ps1 -ShowInfo

# Export modules to ActiveModules folder
.\tools\excel_vba.ps1 -RunMacro "Dev_Exports.ExportModulesToActiveFolder"
```

### Data Manipulation
```powershell
# Get cell value
.\tools\excel_vba.ps1 -GetRange "A1"

# Set cell value
.\tools\excel_vba.ps1 -SetRange "A1","Hello World"

# Get named range value
.\tools\excel_vba.ps1 -GetName "InputCell_DescripSearch"

# Set named range value
.\tools\excel_vba.ps1 -SetName "InputCell_DescripSearch","pump"
```

### Search Dashboard Specific
```powershell
# Perform a search
.\tools\excel_vba.ps1 -SetName "InputCell_DescripSearch","motor"
.\tools\excel_vba.ps1 -RunMacro "mod_PrimaryConsolidatedModule.Safe_PerformSearch"

# Show Sootblower form
.\tools\excel_vba.ps1 -RunMacro "SootblowerFormCreator.CreateAndShowSootblowerForm"

# Run configuration diagnostics
.\tools\excel_vba.ps1 -RunMacro "mod_PrimaryConsolidatedModule.RunConfigDiagnostics"
```

### Interactive Mode Commands
When you run with `-Interactive`, you get a command prompt with these commands:

```
VBA> run QuickSearchDiagnostics.RunQuickSearchDiagnostics
VBA> modules
VBA> get A1
VBA> set A1 Hello
VBA> getname InputCell_DescripSearch
VBA> setname InputCell_DescripSearch pump
VBA> exec MsgBox "Hello from VBA"
VBA> showform frmSootblowerLocator
VBA> info
VBA> quit
```

## Key Features

### 1. Full VBA Access
- Run any VBA procedure or function
- Execute single VBA statements
- Access all workbook objects and properties

### 2. UserForm Control
- Show/hide UserForms programmatically
- Access form controls and properties
- Handle form events through VBA

### 3. Data Integration
- Read/write cells and ranges
- Access named ranges
- Manipulate tables and data structures

### 4. Module Management
- List all VBA modules
- View module code
- Import/export modules

### 5. Real-time Interaction
- Interactive command mode
- Background execution support
- Live workbook manipulation

## Search Dashboard Integration

The controller is specifically designed to work with the Search Dashboard project:

### Quick Operations
```powershell
# Full diagnostic run
.\tools\excel_vba.ps1 -RunMacro "QuickSearchDiagnostics.RunQuickSearchDiagnostics"

# Sync modules from ActiveModules folder
.\tools\excel_vba.ps1 -RunMacro "SyncManager.SyncModules_FromActiveFolder"

# Test search functionality
.\tools\excel_vba.ps1 -RunMacro "Dev_SmokeTests.RunBasicSearchTest"

# Create sootblower form
.\tools\excel_vba.ps1 -RunMacro "SootblowerFormCreator.CreateAndShowSootblowerForm"
```

### Mode-based Operations
```powershell
# Set search mode
.\tools\excel_vba.ps1 -SetName "ModeSelector","Sootblower Location"

# Perform mode-based search
.\tools\excel_vba.ps1 -RunMacro "mod_ModeDrivenSearch.OutputModeResults"
```

## Troubleshooting

### Excel Trust Settings
If you get VBA security errors, enable:
1. **Developer Tab**: File → Options → Customize Ribbon → Developer
2. **Macro Settings**: Developer → Macro Settings → Enable all macros
3. **Trust VBA**: File → Options → Trust Center → Trust access to VBA project object model

### COM Automation Issues
If Excel doesn't respond:
```powershell
# Try with hidden Excel
.\tools\excel_vba.ps1 -Hidden -ShowInfo

# Or restart with fresh instance
.\tools\excel_vba.ps1 -RunMacro "Application.Quit"
```

### Python Path Issues
If Python isn't found, update the path in `excel_vba.ps1`:
```powershell
$pythonPath = "C:\Users\joshu\AppData\Local\Programs\Python\Python312\python.exe"
```

## Advanced Usage

### Custom Scripts
You can create custom automation scripts by importing the controller:

```python
from tools.excel_vba_controller import ExcelVBAController

controller = ExcelVBAController("Search Dashboard v1.3.xlsm")
if controller.connect():
    # Run diagnostics
    controller.run_macro("QuickSearchDiagnostics.RunQuickSearchDiagnostics")
    
    # Set search input
    controller.set_named_range_value("InputCell_DescripSearch", "pump")
    
    # Perform search
    controller.run_macro("mod_PrimaryConsolidatedModule.Safe_PerformSearch")
    
    controller.disconnect()
```

### Batch Operations
Create batch files for common operations:

```cmd
@echo off
echo Running Search Dashboard Diagnostics...
tools\excel_vba.bat --run-macro "QuickSearchDiagnostics.RunQuickSearchDiagnostics"
echo.
echo Exporting modules...
tools\excel_vba.bat --run-macro "Dev_Exports.ExportModulesToActiveFolder"
echo Done!
```

## Security Notes

- The controller runs with full Excel automation privileges
- VBA execution happens in the Excel security context
- All macros respect Excel's security settings
- UserForm access requires Trust VBA project object model

---

**You now have complete terminal control over your Excel VBA project!** 🎉