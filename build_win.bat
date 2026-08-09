@echo off
REM BKTC_Ledger Windows build (onedir). ASCII-only to avoid cmd encoding issues.
REM Needs Python 3.11 or 3.12 on Win10/11. See README_win.md.
cd /d "%~dp0"

echo === Select Python (3.11 preferred, 3.12 fallback) ===
set "PYCMD="
py -3.11 --version >nul 2>&1 && set "PYCMD=py -3.11"
if not defined PYCMD ( py -3.12 --version >nul 2>&1 && set "PYCMD=py -3.12" )
if not defined PYCMD (
  echo [X] Python 3.11/3.12 not found. Install from python.org (tick "Add to PATH" + py launcher).
  pause & exit /b 1
)
%PYCMD% --version

echo === Create venv (first run) ===
if not exist .venv ( %PYCMD% -m venv .venv )
call .venv\Scripts\activate.bat
python -m pip install -q -U pip
echo === Install deps ===
python -m pip install -q pywebview openpyxl pyinstaller

echo === Run reliability tests ===
python -m unittest discover -s tests -v
if errorlevel 1 (
  echo [X] Tests failed. Build stopped.
  pause & exit /b 1
)

echo === PyInstaller build (onedir, unsigned) ===
pyinstaller --noconfirm --windowed --clean --name BKTC_Ledger --add-data "ui;ui" --collect-all webview --hidden-import "webview.platforms.edgechromium" app.py

echo.
echo === Done. Output: dist\BKTC_Ledger\ (copy the WHOLE folder to users) ===
echo See README_win.md for SmartScreen / antivirus / data-file notes.
pause
