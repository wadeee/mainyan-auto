@echo off
chcp 65001 >nul
cd /d "%~dp0"

for /f %%a in ('powershell -NoProfile -Command "(Get-Date).AddDays(-21).ToString('yyyy-MM-dd')"') do set TARGET_DATE=%%a
echo 正在导出 %TARGET_DATE% 工厂配送麦安研门店周度销售报表...

call C:\ProgramData\anaconda3\condabin\conda.bat activate mainyan-auto

python factory_delivery_mainyan_weekly.py --headless --weeks -3 %*
if errorlevel 1 (
    echo.
    echo 任务失败，请查看上方错误信息。
    pause
) else (
    echo.
    echo 完成！文件已保存到当前目录。
    timeout /t 3
)
