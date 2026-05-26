@echo off

set SCRIPT=C:\Users\Wadec\Documents\projects\mainyan-autoproduct_request_export.py

schtasks /create /f ^
  /tn "ProductExport_Days1" ^
  /tr python \"%SCRIPT%\" --headless --days 1" ^
  /sc daily /st 07:50 ^
  /ru SYSTEM

schtasks /create /f ^
  /tn "ProductExport_Days2" ^
  /tr python \"%SCRIPT%\" --headless --days 2" ^
  /sc daily /st 12:50 ^
  /ru SYSTEM

echo Done
pause
