@echo off
REM Excel VBA Controller - Batch Wrapper
REM ====================================

set PYTHON_PATH=C:\Users\joshu\AppData\Local\Programs\Python\Python312\python.exe
set SCRIPT_PATH=%~dp0excel_vba_controller.py

REM Check if Python exists
if not exist "%PYTHON_PATH%" (
    echo ERROR: Python not found at %PYTHON_PATH%
    exit /b 1
)

REM Check if script exists
if not exist "%SCRIPT_PATH%" (
    echo ERROR: Script not found at %SCRIPT_PATH%
    exit /b 1
)

REM If no arguments, show help
if "%1"=="" (
    echo Excel VBA Controller - Batch Wrapper
    echo ====================================
    echo.
    echo Usage:
    echo   excel_vba.bat [arguments for excel_vba_controller.py]
    echo.
    echo Common examples:
    echo   excel_vba.bat --interactive
    echo   excel_vba.bat --run-macro "QuickSearchDiagnostics.RunQuickSearchDiagnostics"
    echo   excel_vba.bat --list-modules
    echo   excel_vba.bat --show-info
    echo.
    echo For full help:
    echo   excel_vba.bat --help
    exit /b 0
)

REM Execute with all arguments passed through
"%PYTHON_PATH%" "%SCRIPT_PATH%" %*