@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: 从文件名提取年月（如 _2026-6 → 2026.06.01）
set "BAT_NAME=%~n0"
for /f %%a in ('powershell -NoProfile -Command "if ('%BAT_NAME%' -match '(\d{4})-(\d{1,2})$') { '{0}.{1:D2}.01' -f $matches[1], [int]$matches[2] }"') do set TARGET_DATE=%%a
echo 正在导出 %TARGET_DATE:~0,7% 工厂配送麦安研+焙满香门店商品月度销售对比表...

for /f "tokens=1,* delims==" %%a in (config.env) do if "%%a"=="PYTHON_PATH" set "PYTHON_PATH=%%b"

"%PYTHON_PATH%" factory_delivery_mainyan_store_prod_sales_monthly.py --headless --date %TARGET_DATE% %*
if errorlevel 1 (
    echo.
    echo 任务失败，请查看上方错误信息。
    pause
) else (
    echo.
    echo 完成！文件已保存到当前目录。
    timeout /t 3
)
