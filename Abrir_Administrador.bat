@echo off
setlocal
cd /d "%~dp0"
python app\server.py
if errorlevel 1 pause
endlocal
