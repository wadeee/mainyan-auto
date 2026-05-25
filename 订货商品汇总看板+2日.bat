@echo off
chcp 65001 >nul
echo 正在导出后天（%DATE%）订货商品汇总看板...
cd /d "%~dp0"

call C:\ProgramData\anaconda3\condabin\conda.bat activate mainyan-auto

python product_request_export.py --headless --days 2 %*
if errorlevel 1 (
    echo.
    echo 任务失败，请查看上方错误信息。
    pause
) else (
    echo.
    echo 完成！文件已保存到当前目录。
    timeout /t 3
)
