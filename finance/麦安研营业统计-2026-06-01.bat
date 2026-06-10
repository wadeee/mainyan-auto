@echo off
chcp 65001 >nul
cd /d "%~dp0"

set TARGET_DATE=2026-06-01
echo 正在导出 %TARGET_DATE% 兔司家门店订购商品统计...

call C:\ProgramData\anaconda3\condabin\conda.bat activate mainyan-auto

python mainyan_turnover_statistics.py --headless --date 2026.06.01 %*
if errorlevel 1 (
    echo.
    echo 任务失败，请查看上方错误信息。
    pause
) else (
    echo.
    echo 完成！文件已保存到当前目录。
    timeout /t 3
)
