@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

REM 将项目内所有 .bat 文件的换行符转换为 CRLF

set "SCRIPT_DIR=%~dp0"
set count=0

for /r "%SCRIPT_DIR%" %%f in (*.bat) do (
    powershell -NoProfile -Command "$p='%%f'; $c=[IO.File]::ReadAllText($p); $c=$c -replace \"`r`n\",\"`n\"; $c=$c -replace \"`n\",\"`r`n\"; [IO.File]::WriteAllText($p,$c)"
    echo   CRLF: %%f
    set /a count+=1
)

echo Done. !count! .bat files converted to CRLF.
endlocal
