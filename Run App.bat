@echo off
cd /d "%~dp0app_files"

:: Activate the virtual environment
call env\Scripts\activate.bat

:: Install dependencies
pip install -r requirements.txt

:: Check Python version (optional)
python --version

:: Run Flask app and open browser

start /B python app.py
timeout /t 5 /nobreak > NUL
start "" http://localhost:5000

pause