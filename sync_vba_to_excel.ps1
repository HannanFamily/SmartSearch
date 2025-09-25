# VBA Synchronization PowerShell Script
# ====================================
# This script helps synchronize VBA code from .bas/.cls files into Excel workbooks
# Usage: .\sync_vba_to_excel.ps1 [workbook_name]

param(
    [string]$WorkbookName = ""
)

# Configuration
$ProjectDir = Get-Location
$BackupDir = Join-Path $ProjectDir "Old_Code"

Write-Host "=== VBA to Excel Synchronization ===" -ForegroundColor Green

# Find Excel workbook
$xlsmFiles = Get-ChildItem -Path $ProjectDir -Filter "*.xlsm" | Where-Object { $_.Name -notlike "*backup*" }

if ($xlsmFiles.Count -eq 0) {
    Write-Host "ERROR: No .xlsm files found in current directory" -ForegroundColor Red
    exit 1
}

if ($WorkbookName) {
    $targetWorkbook = $xlsmFiles | Where-Object { $_.Name -eq $WorkbookName }
    if (-not $targetWorkbook) {
        Write-Host "ERROR: Workbook '$WorkbookName' not found" -ForegroundColor Red
        exit 1
    }
} elseif ($xlsmFiles.Count -eq 1) {
    $targetWorkbook = $xlsmFiles[0]
} else {
    Write-Host "Multiple workbooks found:" -ForegroundColor Yellow
    for ($i = 0; $i -lt $xlsmFiles.Count; $i++) {
        Write-Host "  $($i + 1). $($xlsmFiles[$i].Name)"
    }
    $choice = Read-Host "Select workbook (1-$($xlsmFiles.Count))"
    try {
        $index = [int]$choice - 1
        if ($index -ge 0 -and $index -lt $xlsmFiles.Count) {
            $targetWorkbook = $xlsmFiles[$index]
        } else {
            Write-Host "Invalid selection" -ForegroundColor Red
            exit 1
        }
    } catch {
        Write-Host "Invalid input" -ForegroundColor Red
        exit 1
    }
}

Write-Host "Target workbook: $($targetWorkbook.Name)" -ForegroundColor Cyan

# Find VBA files
$vbaFiles = Get-ChildItem -Path $ProjectDir -Include "*.bas", "*.cls" -Name

if ($vbaFiles.Count -eq 0) {
    Write-Host "No VBA files found in project directory" -ForegroundColor Yellow
    exit 0
}

Write-Host "Found $($vbaFiles.Count) VBA files:" -ForegroundColor Cyan
foreach ($file in $vbaFiles) {
    $type = if ($file.EndsWith('.bas')) { 'Module' } else { 'Class' }
    Write-Host "  - $file ($type)"
}

# Create backup directory if it doesn't exist
if (-not (Test-Path $BackupDir)) {
    New-Item -ItemType Directory -Path $BackupDir | Out-Null
}

# Create backup
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupName = $targetWorkbook.BaseName + "_backup_$timestamp" + $targetWorkbook.Extension
$backupPath = Join-Path $BackupDir $backupName

Copy-Item -Path $targetWorkbook.FullName -Destination $backupPath
Write-Host "Backup created: $backupPath" -ForegroundColor Green

# Generate VBA import script
$importScript = @"
Sub ImportVBAModules()
    ' Auto-generated VBA import script
    ' Run this macro in Excel to import updated modules
    
    Dim fso As Object
    Dim projectPath As String
    Dim vbcomp As Object
    Dim moduleName As String
    Dim filePath As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    projectPath = ThisWorkbook.Path
    
    ' Disable alerts and screen updating for smoother operation
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Debug.Print "Starting VBA module import at " & Now()
    Debug.Print "Project path: " & projectPath
    Debug.Print String(50, "=")
    
"@

foreach ($file in $vbaFiles) {
    $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file)
    $importScript += @"
    
    ' Process $file
    moduleName = "$moduleName"
    filePath = projectPath & "\$file"
    
    If fso.FileExists(filePath) Then
        ' Remove existing module if it exists
        On Error Resume Next
        ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(moduleName)
        On Error GoTo 0
        
        ' Import the module
        Set vbcomp = ThisWorkbook.VBProject.VBComponents.Import(filePath)
        Debug.Print "✓ Imported: $file"
    Else
        Debug.Print "✗ File not found: $file"
    End If
"@
}

$importScript += @"

    
    ' Re-enable alerts and screen updating
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    Debug.Print String(50, "=")
    Debug.Print "VBA module import completed at " & Now()
    MsgBox "VBA module import complete!" & vbCrLf & vbCrLf & _
           "Check the Immediate Window (Ctrl+G) for detailed results.", vbInformation, "Import Complete"
End Sub

Sub ShowImportInstructions()
    ' Helper macro to display import instructions
    MsgBox "To import VBA modules:" & vbCrLf & vbCrLf & _
           "1. Run the 'ImportVBAModules' macro" & vbCrLf & _
           "2. Check Immediate Window for results" & vbCrLf & _
           "3. Save your workbook" & vbCrLf & _
           "4. Test your updated functionality", vbInformation, "VBA Import Instructions"
End Sub
"@

$importScriptPath = Join-Path $ProjectDir "import_vba_modules.bas"
$importScript | Out-File -FilePath $importScriptPath -Encoding UTF8

# Generate detailed instructions
$instructions = @"
VBA MODULE SYNCHRONIZATION INSTRUCTIONS
=======================================
Workbook: $($targetWorkbook.Name)
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Backup: $backupName

AUTOMATIC METHOD (Recommended):
1. Open $($targetWorkbook.Name) in Excel
2. Enable macros if prompted
3. Press Alt+F11 to open VBA Editor
4. Import the file 'import_vba_modules.bas':
   - File > Import File... > Select 'import_vba_modules.bas'
5. Run the 'ImportVBAModules' macro:
   - In VBA Editor: Run > Run Sub/UserForm or press F5
6. Check Immediate Window (Ctrl+G) for import results
7. Save the workbook (Ctrl+S)
8. Delete the temporary 'import_vba_modules' module when done

MANUAL METHOD (Alternative):
1. Open $($targetWorkbook.Name) in Excel
2. Press Alt+F11 to open VBA Editor
3. For each module below:
   a. In Project Explorer, right-click on existing module (if exists)
   b. Choose "Remove [ModuleName]" (export first if you want to keep old version)
   c. Right-click in Project Explorer
   d. Choose "Import File..."
   e. Select the corresponding .bas or .cls file

MODULES TO UPDATE:
"@

foreach ($file in $vbaFiles) {
    $moduleName = [System.IO.Path]::GetFileNameWithoutExtension($file)
    $fileInfo = Get-Item (Join-Path $ProjectDir $file)
    $type = if ($file.EndsWith('.bas')) { 'Standard Module' } else { 'Class Module' }
    
    $instructions += @"

- $file ($type)
  Module Name: $moduleName
  File Size: $($fileInfo.Length) bytes
  Modified: $($fileInfo.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'))
  Full Path: $($fileInfo.FullName)
"@
}

$instructions += @"


TROUBLESHOOTING:
- If you get "Programmatic access to VBA not allowed" error:
  1. File > Options > Trust Center > Trust Center Settings
  2. Macro Settings > Check "Trust access to the VBA project object model"
- If modules don't import, check file permissions and paths
- Always test functionality after importing

WORKFLOW NOTES:
- VBA files in this directory are the source of truth
- Use Git to track changes to .bas/.cls files
- Run this sync script after making changes to VBA files
- Keep regular backups in the Old_Code directory

NEXT STEPS:
1. Open Excel workbook
2. Follow the AUTOMATIC METHOD above
3. Test your updated functionality
4. Commit changes to Git if everything works correctly
"@

$instructionsPath = Join-Path $ProjectDir "VBA_SYNC_INSTRUCTIONS.txt"
$instructions | Out-File -FilePath $instructionsPath -Encoding UTF8

# Summary
Write-Host "`n=== SYNCHRONIZATION READY ===" -ForegroundColor Green
Write-Host "Created files:" -ForegroundColor Cyan
Write-Host "  - $backupName (backup in Old_Code/)"
Write-Host "  - import_vba_modules.bas (import script)"
Write-Host "  - VBA_SYNC_INSTRUCTIONS.txt (detailed instructions)"

Write-Host "`n=== NEXT STEPS ===" -ForegroundColor Yellow
Write-Host "1. Open $($targetWorkbook.Name) in Excel"
Write-Host "2. Follow instructions in 'VBA_SYNC_INSTRUCTIONS.txt'"
Write-Host "3. Use 'import_vba_modules.bas' for automatic import"
Write-Host "`nSynchronization preparation complete!" -ForegroundColor Green