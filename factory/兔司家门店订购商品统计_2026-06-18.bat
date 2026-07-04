@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: 从文件名提取日期（如 _2026-06-18 → 2026-06-18）
set "BAT_NAME=%~n0"
set "TARGET_DATE=%BAT_NAME:~-10%"
echo 正在导出 %TARGET_DATE% 兔司家门店订购商品统计...

for /f "tokens=1,* delims==" %%a in (config.env) do if "%%a"=="PYTHON_PATH" set "PYTHON_PATH=%%b"

"%PYTHON_PATH%" tsj_prod_order_statistics.py --headless --date %TARGET_DATE:-=.% %*
if errorlevel 1 (
    echo.
    echo 任务失败，请查看上方错误信息。
    pause
) else (
    echo.
    echo 完成！文件已保存到当前目录。
    timeout /t 3
)
