@echo off
REM PST/OST Email Exporter - Windows Batch Wrapper
REM This batch file provides a convenient way to run the PST exporter on Windows

REM Check if Python is available
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python 3.7+ and make sure it's added to your PATH
    pause
    exit /b 1
)

REM Run the PST exporter with all passed arguments
python "%~dp0pst-exporter.py" %*

REM Pause only if the script was double-clicked (not run from command line)
if "%1"=="" (
    pause
)
