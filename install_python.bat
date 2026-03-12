@echo off
echo ===============================================
echo FORBO INVENTUR-PROGRAMM - PYTHON INSTALLATION
echo ===============================================
echo.

:: Prüfe ob Python bereits installiert ist
python --version >nul 2>&1
if %errorlevel% == 0 (
    echo [OK] Python ist bereits installiert:
    python --version
    echo.
    goto :install_modules
)

echo [INFO] Python ist nicht installiert. Starte Download...
echo.

:: Erstelle temporären Download-Ordner
if not exist "%TEMP%\forbo_install" mkdir "%TEMP%\forbo_install"
cd /d "%TEMP%\forbo_install"

:: Download Python 3.11.9 (Windows x64)
echo [INFO] Lade Python 3.11.9 herunter...
echo Bitte warten, dies kann einige Minuten dauern...
powershell -Command "Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe' -OutFile 'python-installer.exe'"

if not exist "python-installer.exe" (
    echo [FEHLER] Python-Download fehlgeschlagen!
    echo.
    echo Bitte laden Sie Python manuell herunter:
    echo https://www.python.org/downloads/
    echo.
    pause
    exit /b 1
)

echo [INFO] Installiere Python...
echo.
:: Stille Installation mit pip und PATH
python-installer.exe /quiet InstallAllUsers=1 PrependPath=1 Include_pip=1

:: Warte auf Installation
timeout /t 10 /nobreak >nul

:: Prüfe Installation
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [FEHLER] Python-Installation fehlgeschlagen!
    echo Bitte installieren Sie Python manuell und starten Sie dieses Script erneut.
    pause
    exit /b 1
)

echo [OK] Python erfolgreich installiert:
python --version
echo.

:: Aufräumen
cd /d "%~dp0"
rmdir /s /q "%TEMP%\forbo_install" 2>nul

:install_modules
echo ===============================================
echo INSTALLIERE PYTHON-MODULE
echo ===============================================
echo.

:: Upgrade pip
echo [INFO] Aktualisiere pip...
python -m pip install --upgrade pip

:: Installiere erforderliche Module
echo [INFO] Installiere pandas...
python -m pip install pandas

echo [INFO] Installiere openpyxl...
python -m pip install openpyxl

echo [INFO] Installiere pillow (für Icons)...
python -m pip install pillow

echo.
echo ===============================================
echo INSTALLATION ABGESCHLOSSEN
echo ===============================================
echo.

:: Teste Module
echo [TEST] Teste Python-Module...
python -c "import pandas; import openpyxl; import tkinter; print('[OK] Alle Module erfolgreich geladen')" 2>nul
if %errorlevel% neq 0 (
    echo [WARNUNG] Einige Module konnten nicht geladen werden.
    echo Bitte prüfen Sie die Installation.
) else (
    echo [OK] Alle Module funktionieren korrekt!
)

echo.
echo Sie können jetzt das Inventur-Programm starten mit:
echo start_inventur.bat
echo.
pause
