@echo off

start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
    --remote-debugging-port=9223 ^
    --user-data-dir="C:\ChromeDebug_BT" ^
    --new-window ^
    "https://e.waimai.meituan.com/" ^
    "https://melody.shop.ele.me/"
