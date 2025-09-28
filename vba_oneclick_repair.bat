@echo off
setlocal
set SCRIPT_DIR=%~dp0
pushd "%SCRIPT_DIR%"
set TOOL_SCRIPT=tools\vba_oneclick_repair.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File "%CD%\%TOOL_SCRIPT%" %*
set ERR=%ERRORLEVEL%
popd
exit /b %ERR%
