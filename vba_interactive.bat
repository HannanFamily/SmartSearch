@echo off
REM Quick VBA Interactive Mode Launcher
REM ==================================
cd /d "\\unraid\systemfiles\allshares\nvmeshare\Dashboard_Project"
powershell -ExecutionPolicy Bypass -Command ".\tools\excel_vba.ps1 -Interactive"