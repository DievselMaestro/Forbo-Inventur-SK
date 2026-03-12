@echo off
:: ============================================================
:: INVENTORY Forbo SK - Malacky Launcher
:: ============================================================
title INVENTORY Forbo SK - Malacky

:: Change to the directory containing this batch file
cd /d "%~dp0"

:: Check if Python is available
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python not found on PATH.
    echo Please install Python 3.11 or later and make sure it is on your PATH.
    pause
    exit /b 1
)

:: Check if the app script exists
if not exist "inventur_app_sk.py" (
    echo ERROR: inventur_app_sk.py not found in %~dp0
    pause
    exit /b 1
)

:: Launch the application
echo Starting INVENTORY Forbo SK - Malacky...
python inventur_app_sk.py

:: If Python exited with an error, pause so the user can read the message
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo The application exited with an error (code %ERRORLEVEL%).
    pause
)
