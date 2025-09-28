param(
    [string]$Workbook = "Search Dashboard v1.3.xlsm",
    [switch]$Hidden,
    [switch]$Visible
)

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$oneClick = Join-Path $scriptRoot "tools\vba_oneclick_repair.ps1"
if (-not (Test-Path $oneClick)) {
    Write-Error "One-click script not found: $oneClick"
    exit 1
}
Write-Host "Launching: $oneClick" -ForegroundColor Yellow
& powershell -NoProfile -ExecutionPolicy Bypass -File $oneClick -Workbook $Workbook @($Hidden ? "-Hidden" : $null) @($Visible ? "-Visible" : $null)
