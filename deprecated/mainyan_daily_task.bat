@echo off
setlocal EnableDelayedExpansion

set "CONFIG=%~dp0config.env"

if not exist "%CONFIG%" (
    echo [ERROR] config.env not found: %CONFIG%
    echo Please run auto-install.bat first.
    exit /b 1
)

for /f "usebackq tokens=1,* delims==" %%a in ("%CONFIG%") do (
    if not "%%a"=="" if not "%%b"=="" set "%%a=%%b"
)

if not defined PYTHON_PATH (
    echo [ERROR] PYTHON_PATH not set in config.env
    exit /b 1
)

if not exist "!PYTHON_PATH!" (
    echo [ERROR] Python not found: !PYTHON_PATH!
    exit /b 1
)

"!PYTHON_PATH!" "%~dp0mainyan_daily_task.py"
