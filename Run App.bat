@echo off
cd /d "%~dp0app_files"

:: Activate the virtual environment
call venv\Scripts\activate.bat

:: Install dependencies
pip install -r requirements.txt

:: Check Python version (optional)
python --version

:: Run Flask app and open browser

start /B venv\Scripts\python.exe app.py
timeout /t 8 /nobreak > NUL

pause

