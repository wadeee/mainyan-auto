@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: 从文件名提取周数（如 -4周 → 4）
set "BAT_NAME=%~n0"
for /f %%a in ('powershell -NoProfile -Command "if ('%BAT_NAME%' -match '-(\d+)周$') { $matches[1] }"') do set WEEKS=%%a
set /a DAYS=WEEKS*7

for /f %%a in ('powershell -NoProfile -Command "(Get-Date).AddDays(-%DAYS%).ToString('yyyy-MM-dd')"') do set TARGET_DATE=%%a
echo 正在导出 %TARGET_DATE% 工厂配送麦安研门店周度销售报表...

for /f "tokens=1,* delims==" %%a in (config.env) do if "%%a"=="PYTHON_PATH" set "PYTHON_PATH=%%b"

"%PYTHON_PATH%" factory_delivery_mainyan_weekly.py --headless --weeks -%WEEKS% %*
if errorlevel 1 (
    echo.
    echo 任务失败，请查看上方错误信息。
    pause
) else (
    echo.
    echo 完成！文件已保存到当前目录。
    timeout /t 3
)
