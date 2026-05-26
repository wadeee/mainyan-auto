@echo off

set "SCRIPT=C:\mainyan-auto\mainyan_daily_task.py"
set "PYTHON=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"

echo PYTHON=%PYTHON% >> "%~dp0mainyan_daily_task.log"
echo SCRIPT=%SCRIPT% >> "%~dp0mainyan_daily_task.log"

"%PYTHON%" "%SCRIPT%"

echo Done
pause
