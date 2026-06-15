@echo off

start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
    --remote-debugging-port=9226 ^
    --user-data-dir="C:\ChromeDebug_MTJYB" ^
    --new-window ^
    "https://ecom.meituan.com/meishi" ^
    "https://ym.o2o.cmbchina.com/mc/merchant/handms/login.html" ^
    "https://life.douyin.com/p/login"
