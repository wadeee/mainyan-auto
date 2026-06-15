@echo off

start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
    --remote-debugging-port=9224 ^
    --user-data-dir="C:\ChromeDebug_LJ" ^
    --new-window ^
    "https://e.waimai.meituan.com/" ^
