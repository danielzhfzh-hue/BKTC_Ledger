@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv (
  python -m venv .venv
  .venv\Scripts\pip install -r requirements.txt
)
call .venv\Scripts\activate.bat
python app.py
pause
