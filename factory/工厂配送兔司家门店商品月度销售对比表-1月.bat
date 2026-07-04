@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: 从文件名提取月数（如 -1月 → -1, +1月 → +1）
set "BAT_NAME=%~n0"
for /f %%a in ('powershell -NoProfile -Command "if ('%BAT_NAME%' -match '([+-]\d+)月$') { $matches[1] }"') do set MONTHS=%%a

for /f %%a in ('powershell -NoProfile -Command "(Get-Date).AddMonths(%MONTHS%).ToString('yyyy-MM-dd')"') do set TARGET_DATE=%%a
echo 正在导出 %TARGET_DATE% 工厂配送兔司家门店商品月度销售对比表...

for /f "tokens=1,* delims==" %%a in (config.env) do if "%%a"=="PYTHON_PATH" set "PYTHON_PATH=%%b"

"%PYTHON_PATH%" factory_delivery_tsj_store_prod_sales_compare_monthly.py --headless --months %MONTHS% %*
if errorlevel 1 (
    echo.
    echo 任务失败，请查看上方错误信息。
    pause
) else (
    echo.
    echo 完成！文件已保存到当前目录。
    timeout /t 3
)
