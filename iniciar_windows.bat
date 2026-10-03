@echo off
setlocal
cd /d "%~dp0"

if not exist .venv\Scripts\python.exe (
  python -m venv .venv
)

call .venv\Scripts\activate.bat
python -m pip install -r requirements.txt
python app.py

if errorlevel 1 pause
