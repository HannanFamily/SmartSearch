<#
Run the isolated data cleanup using the packaged inputs under tools/Data_Cleanup_Package.
This sets environment variables so python/data_cleanup.py reads from the isolated package.
#>
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$packageRoot = Split-Path -Parent $here
$dataDir = Join-Path $packageRoot 'data'
$outputDir = Join-Path $packageRoot 'output'

# Prefer the system Python 3.12 path; fallback to PATH
$pythonExe = Join-Path $env:LOCALAPPDATA 'Programs/Python/Python312/python.exe'
if (-not (Test-Path $pythonExe)) { $pythonExe = 'python' }

# Locate repo root from package path (..\.. from scripts)
$repoRoot = Split-Path -Parent (Split-Path -Parent $packageRoot)
$scriptPath = Join-Path (Join-Path $repoRoot 'python') 'data_cleanup.py'
if (-not (Test-Path $scriptPath)) {
  Write-Host "Cannot find data_cleanup.py at $scriptPath" -ForegroundColor Red
  exit 1
}

# Set environment overrides for this process
$env:DATA_CLEANUP_DIR = $dataDir
$env:DATA_CLEANUP_OUTPUT_DIR = $outputDir

Write-Host "Running data cleanup..." -ForegroundColor Cyan
Write-Host "  DATA_CLEANUP_DIR=$env:DATA_CLEANUP_DIR" -ForegroundColor DarkGray
Write-Host "  DATA_CLEANUP_OUTPUT_DIR=$env:DATA_CLEANUP_OUTPUT_DIR" -ForegroundColor DarkGray

& $pythonExe $scriptPath
$exitCode = $LASTEXITCODE
if ($exitCode -ne 0) {
  Write-Host "Cleanup script exited with code $exitCode" -ForegroundColor Red
  exit $exitCode
}

Write-Host "Cleanup complete. Outputs written under: $outputDir" -ForegroundColor Green
