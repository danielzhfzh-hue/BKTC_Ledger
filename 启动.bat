@echo off
cd /d "%~dp0"
set "PYCMD="
py -3.11 --version >nul 2>&1 && set "PYCMD=py -3.11"
if not defined PYCMD ( py -3.12 --version >nul 2>&1 && set "PYCMD=py -3.12" )
if not defined PYCMD ( python --version >nul 2>&1 && set "PYCMD=python" )
if not defined PYCMD (
  echo Python 3.11 or 3.12 was not found.
  echo Install Python from python.org and enable "Add Python to PATH".
  pause
  exit /b 1
)
if not exist .venv (
  %PYCMD% -m venv .venv
  if errorlevel 1 goto :failed
)
call .venv\Scripts\activate.bat
python -c "import webview, openpyxl" >nul 2>&1
if errorlevel 1 (
  python -m pip install -r requirements.txt
  if errorlevel 1 goto :failed
)
python create_portable_starter.py .
if errorlevel 1 goto :failed
python app.py
if errorlevel 1 goto :failed
exit /b 0

:failed
echo BKTC Ledger failed to start. See the error above.
pause
