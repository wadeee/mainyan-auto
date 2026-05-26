@echo off

set SCRIPT=D:\project\product_request_export.py

schtasks /create /f ^
/tn "ProductExport_Days1" ^
/tr python \"%SCRIPT%\" --headless --days 1" ^
/sc daily /st 08:00

schtasks /create /f ^
/tn "ProductExport_Days2" ^
/tr python \"%SCRIPT%\" --headless --days 2" ^
/sc daily /st 13:00

echo Done
pause
