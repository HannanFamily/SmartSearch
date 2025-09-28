# Excel VBA Controller - PowerShell Wrapper
# =======================================
param(
    [string]$Workbook = "",
    [string]$RunMacro = "",
    [string[]]$MacroArgs = @(),
    [string]$GetRange = "",
    [string[]]$SetRange = @(),
    [string]$GetName = "",
    [string[]]$SetName = @(),
    [switch]$Interactive,
    [switch]$ListModules,
    [switch]$ShowInfo,
    [switch]$Hidden,
    [switch]$Help
)

# Set up environment
$pythonPath = "C:\Users\joshu\AppData\Local\Programs\Python\Python312\python.exe"
$scriptPath = Join-Path $PSScriptRoot "excel_vba_controller.py"

# Default to current workbook
if ($Workbook -eq "") {
    $Workbook = "Search Dashboard v1.3.xlsm"
}

# Help
if ($Help) {
    Write-Host "Excel VBA Controller - PowerShell Wrapper" -ForegroundColor Green
    Write-Host "==========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Parameters:"
    Write-Host "  -Workbook <path>           Path to Excel workbook (default: Search Dashboard v1.3.xlsm)"
    Write-Host "  -RunMacro <name>           Run VBA macro"
    Write-Host "  -MacroArgs <args...>       Arguments for the macro"
    Write-Host "  -GetRange <range>          Get value from range (e.g., 'A1')"
    Write-Host "  -SetRange <range> <value>  Set range value"
    Write-Host "  -GetName <name>            Get named range value"
    Write-Host "  -SetName <name> <value>    Set named range value"
    Write-Host "  -Interactive               Start interactive mode"
    Write-Host "  -ListModules               List VBA modules"
    Write-Host "  -ShowInfo                  Show workbook information"
    Write-Host "  -Hidden                    Keep Excel hidden"
    Write-Host "  -Help                      Show this help"
    Write-Host ""
    Write-Host "Examples:"
    Write-Host "  .\excel_vba.ps1 -RunMacro 'QuickSearchDiagnostics.RunQuickSearchDiagnostics'"
    Write-Host "  .\excel_vba.ps1 -Interactive"
    Write-Host "  .\excel_vba.ps1 -GetRange 'A1'"
    Write-Host "  .\excel_vba.ps1 -SetRange 'A1','Hello World'"
    Write-Host "  .\excel_vba.ps1 -ListModules"
    Write-Host "  .\excel_vba.ps1 -RunMacro 'SootblowerFormCreator.CreateAndShowSootblowerForm'"
    exit 0
}

# Build command line arguments
$args = @()
$args += "--workbook"
$args += "`"$Workbook`""

if ($Hidden) {
    $args += "--hidden"
}

if ($RunMacro) {
    $args += "--run-macro"
    $args += "`"$RunMacro`""
    if ($MacroArgs.Count -gt 0) {
        $args += "--macro-args"
        $args += $MacroArgs
    }
}

if ($GetRange) {
    $args += "--get-range"
    $args += "`"$GetRange`""
}

if ($SetRange.Count -eq 2) {
    $args += "--set-range"
    $args += "`"$($SetRange[0])`""
    $args += "`"$($SetRange[1])`""
}

if ($GetName) {
    $args += "--get-name"
    $args += "`"$GetName`""
}

if ($SetName.Count -eq 2) {
    $args += "--set-name"
    $args += "`"$($SetName[0])`""
    $args += "`"$($SetName[1])`""
}

if ($Interactive) {
    $args += "--interactive"
}

if ($ListModules) {
    $args += "--list-modules"
}

if ($ShowInfo) {
    $args += "--show-info"
}

# Execute the Python script
Write-Host "Launching Excel VBA Controller..." -ForegroundColor Yellow
Write-Host "Command: $pythonPath `"$scriptPath`" $($args -join ' ')" -ForegroundColor Gray
& $pythonPath $scriptPath @args