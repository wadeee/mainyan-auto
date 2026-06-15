@echo off
chcp 65001 >nul
cd /d "%~dp0"

:: 获取文件名中的日期
set "BAT_NAME=%~n0"
set "TARGET_DATE=%BAT_NAME:~-10%"
set "PY_DATE=%TARGET_DATE:-=.%"

echo 正在导出 %TARGET_DATE% 兔司家门店订购商品统计...

call C:\ProgramData\anaconda3\condabin\conda.bat activate mainyan-auto

python mainyan_turnover_statistics.py --headless --date %PY_DATE% %*

if errorlevel 1 (
    echo.
    echo 任务失败，请查看上方错误信息。
    pause
) else (
    echo.
    echo 完成！文件已保存到当前目录。
    timeout /t 3
)
