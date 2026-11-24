@echo off
cd /d "%~dp0"
echo Iniciando Caminata Gemba...
start "" "http://127.0.0.1:5150"
".venv\Scripts\python.exe" app.py
pause
