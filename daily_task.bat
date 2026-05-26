@echo off

set PYTHON=%USERPROFILE%\AppData\Local\Programs\Python\Python311\python.exe
set SCRIPT=%~dp0product_request_export.py

schtasks /create /f /tn "ProductExport_Days1" /tr "\"%PYTHON%\" \"%SCRIPT%\" --headless --days 1" /sc daily /st 07:50 /ru SYSTEM

schtasks /create /f /tn "ProductExport_Days2" /tr ""\"%PYTHON%\" \"%SCRIPT%\" --headless --days 2" /sc daily /st 12:50 /ru SYSTEM

echo Done
pause
