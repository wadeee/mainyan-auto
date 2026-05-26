@echo off

set "SCRIPT=%~dp0mainyan_daily_task.py"
set "PYTHON=%LOCALAPPDATA%Local\Programs\Python\Python311\python.exe"

"%PYTHON%" "%SCRIPT%"

echo Done
pause
