@echo off

start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
    --remote-debugging-port=9225 ^
    --user-data-dir="C:\ChromeDebug_XT" ^
    --new-window ^
    "https://e.waimai.meituan.com/" ^
