# One-Click Repair & Diagnostics for Search Dashboard
# ===================================================
param(
    [string]$Workbook = "Search Dashboard v1.3.xlsm",
    [switch]$Hidden,
    [switch]$Visible
)

$ErrorActionPreference = "Continue"

function New-LogDir {
    $root = Join-Path (Split-Path -Parent $PSScriptRoot) "logs"
    $ocRoot = Join-Path $root "OneClickRuns"
    if (-not (Test-Path $ocRoot)) { New-Item -ItemType Directory -Force -Path $ocRoot | Out-Null }
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $dir = Join-Path $ocRoot "Run_$ts"
    New-Item -ItemType Directory -Force -Path $dir | Out-Null
    return $dir
}

function Invoke-Controller {
    param(
        [string[]]$Args,
        [string]$LogFile
    )
    $exe = Join-Path $PSScriptRoot "excel_vba.ps1"
    Write-Host "→ excel_vba.ps1 $($Args -join ' ')" -ForegroundColor Gray
    $out = & $exe @Args 2>&1
    if ($out) { $out | Out-File -FilePath $LogFile -Append -Encoding utf8 }
    return ,$out
}

function Was-Success {
    param([string[]]$Output)
    if (-not $Output) { return $false }
    # Heuristics indicating success
    return ($Output -match "Macro completed successfully").Length -gt 0 -or \
           ($Output -match "Macro result:").Length -gt 0 -or \
           ($Output -match "VBA Modules:").Length -gt 0
}

function Try-Macros {
    param(
        [string]$Title,
        [string[]]$MacroNames,
        [string]$LogDir
    )
    $stepLog = Join-Path $LogDir ("${Title}.log")
    Write-Host ("=== {0} ===" -f $Title) -ForegroundColor Cyan
    $visibilityArgs = @()
    if ($Hidden) { $visibilityArgs += "-Hidden" } elseif ($Visible) { $visibilityArgs += "-Visible" } else { $visibilityArgs += "-Visible" }
    foreach ($m in $MacroNames) {
        Write-Host ("Trying macro: {0}" -f $m) -ForegroundColor Yellow
        $args = @("-Workbook", $Workbook, "-RunMacro", $m) + $visibilityArgs
        $out = Invoke-Controller -Args $args -LogFile $stepLog
        if (Was-Success -Output $out) {
            Write-Host ("SUCCESS: {0}" -f $m) -ForegroundColor Green
            return $true
        } else {
            Write-Host ("Failed: {0}" -f $m) -ForegroundColor DarkYellow
        }
    }
    Write-Host ("No macro variant succeeded for {0}" -f $Title) -ForegroundColor Red
    return $false
}

function OneClick-Run {
    $logDir = New-LogDir
    Write-Host ("Logs: {0}" -f $logDir) -ForegroundColor Gray

    # 0) Show workbook info and modules
    $infoLog = Join-Path $logDir "00_Info.log"
    Invoke-Controller -Args @("-Workbook", $Workbook, "-ShowInfo", "-Visible") -LogFile $infoLog | Out-Null
    Invoke-Controller -Args @("-Workbook", $Workbook, "-ListModules", "-Visible") -LogFile $infoLog | Out-Null

    # 1) Sync modules (non-destructive)
    Try-Macros -Title "01_SyncModules" -MacroNames @(
        "SyncModules_FromActiveFolder",
        "ActiveModuleImporter.SyncModules_FromActiveFolder",
        "Dev_ControlCenter.RUN_Dev_SyncModules"
    ) -LogDir $logDir | Out-Null

    # 2) Quick smoke test
    Try-Macros -Title "02_SmokeTest" -MacroNames @(
        "Dev_SmokeTests.RUN_SmokeTest_Workbook",
        "RUN_SmokeTest_Workbook"
    ) -LogDir $logDir | Out-Null

    # 3) Diagnostics (config + search)
    Try-Macros -Title "03_RunConfigDiagnostics" -MacroNames @(
        "RunConfigDiagnostics",
        "mod_PrimaryConsolidatedModule.RunConfigDiagnostics"
    ) -LogDir $logDir | Out-Null

    Try-Macros -Title "04_RunQuickSearchDiagnostics" -MacroNames @(
        "QuickSearchDiagnostics.RunQuickSearchDiagnostics",
        "RunQuickSearchDiagnostics"
    ) -LogDir $logDir | Out-Null

    # 4) Ensure Sootblower mode + config keys
    Try-Macros -Title "05_Ensure_SSB_ModeConfig" -MacroNames @(
        "temp_mod_ConfigTableTools.Ensure_ModeConfigEntry_SootblowerLocation",
        "ConfigTableTools.Ensure_ModeConfigEntry_SootblowerLocation"
    ) -LogDir $logDir | Out-Null

    Try-Macros -Title "06_Ensure_SSB_ConfigKeys" -MacroNames @(
        "temp_mod_ConfigTableTools.Ensure_ConfigKeys_Sootblower"
    ) -LogDir $logDir | Out-Null

    # 5) Initialize Sootblower Locator (auto-creates form if needed)
    Try-Macros -Title "07_Init_SootblowerLocator" -MacroNames @(
        "mod_SootblowerLocator.Init_SootblowerLocator",
        "Init_SootblowerLocator"
    ) -LogDir $logDir | Out-Null

    # 6) Attempt to show the form explicitly (if available)
    Try-Macros -Title "08_Show_SSB_Form" -MacroNames @(
        "temp_mod_SSB_FormStandardizer.Show_SSB_Form"
    ) -LogDir $logDir | Out-Null

    # 7) Export snapshot (modules + CSVs)
    Try-Macros -Title "09_Export_ProjectSnapshot" -MacroNames @(
        "Dev_Exports.RUN_Export_ProjectSnapshot",
        "RUN_Export_ProjectSnapshot"
    ) -LogDir $logDir | Out-Null

    Write-Host "One-click run complete. Review logs and Excel window." -ForegroundColor Green
    Write-Host "Log folder: $logDir" -ForegroundColor Green
}

OneClick-Run
