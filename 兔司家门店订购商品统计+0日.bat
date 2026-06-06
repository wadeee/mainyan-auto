@echo off
chcp 65001 >nul
cd /d "%~dp0"

for /f %%a in ('powershell -NoProfile -Command "(Get-Date).AddDays(0).ToString('yyyy-MM-dd')"') do set TARGET_DATE=%%a
echo 正在导出 %TARGET_DATE% 兔司家门店订购商品统计...

call C:\ProgramData\anaconda3\condabin\conda.bat activate mainyan-auto

python tsj_prod_order_statistics.py --headless --days 0 %*
if errorlevel 1 (
    echo.
    echo 任务失败，请查看上方错误信息。
    pause
) else (
    echo.
    echo 完成！文件已保存到当前目录。
    timeout /t 3
)
