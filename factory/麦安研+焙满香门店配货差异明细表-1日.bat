@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: 从文件名提取天数（如 -1日 → -1, +1日 → +1）
set "BAT_NAME=%~n0"
for /f %%a in ('powershell -NoProfile -Command "if ('%BAT_NAME%' -match '([+-]\d+)日$') { $matches[1] }"') do set DAYS=%%a

for /f %%a in ('powershell -NoProfile -Command "(Get-Date).AddDays(%DAYS%).ToString('yyyy-MM-dd')"') do set TARGET_DATE=%%a
echo 正在导出 %TARGET_DATE% 麦安研+焙满香门店配货差异明细表...

for /f "tokens=1,* delims==" %%a in (config.env) do if "%%a"=="PYTHON_PATH" set "PYTHON_PATH=%%b"

"%PYTHON_PATH%" mainyan_store_distribution_diff_detail.py --headless --days %DAYS% %*
if errorlevel 1 (
    echo.
    echo 任务失败，请查看上方错误信息。
    pause
) else (
    echo.
    echo 完成！文件已保存到当前目录。
    timeout /t 3
)
