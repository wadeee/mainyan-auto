@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo 正在清除自动化任务占用的 Chrome 进程...

for /f "tokens=1,* delims==" %%a in (config.env) do if "%%a"=="PYTHON_PATH" set "PYTHON_PATH=%%b"

"%PYTHON_PATH%" cleanup_stale_browsers.py %*
if errorlevel 1 (
    echo.
    echo 任务失败，请查看上方错误信息。
    pause
) else (
    echo.
    echo 完成！Chrome 占用已清除。
    timeout /t 3
)
