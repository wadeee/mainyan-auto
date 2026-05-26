@echo off
setlocal enabledelayedexpansion

REM ===============================
REM 基础路径配置
REM ===============================
set "SCRIPT_DIR=%~dp0"
set "SCRIPT=%SCRIPT_DIR%product_request_export.py"
set "LOG_DIR=%SCRIPT_DIR%logs"

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

REM 时间戳（用于日志）
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set dt=%%I
set "DATESTAMP=%dt:~0,8%"
set "TIMESTAMP=%dt:~8,6%"

REM ===============================
REM 自动检测 Python
REM ===============================
set "PYTHON="

for %%P in (
    python
    py
    "%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    "%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    "%ProgramFiles%\Python311\python.exe"
    "%ProgramFiles%\Python310\python.exe"
) do (
    %%P --version >nul 2>&1
    if !errorlevel! == 0 (
        set "PYTHON=%%P"
        goto :FOUND_PYTHON
    )
)

:FOUND_PYTHON
if "%PYTHON%"=="" (
    echo [ERROR] Python not found >> "%LOG_DIR%\setup_error.log"
    echo Python not found
    pause
    exit /b 1
)

echo [INFO] Using Python: %PYTHON% >> "%LOG_DIR%\setup.log"

REM ===============================
REM 任务1
REM ===============================
set "TASK1=ProductExport_Days1"
set "CMD1=%PYTHON% \"%SCRIPT%\" --headless --days 1"

schtasks /create /f ^
 /tn "%TASK1%" ^
 /tr "%CMD1%" ^
 /sc daily ^
 /st 7:50 ^
 /ru SYSTEM ^
 >> "%LOG_DIR%\setup.log" 2>&1

if %errorlevel% neq 0 (
    echo [%DATE% %TIME%] Task1 create FAILED >> "%LOG_DIR%\setup_error.log"
) else (
    echo [%DATE% %TIME%] Task1 create OK >> "%LOG_DIR%\setup.log"
)

REM ===============================
REM 任务2
REM ===============================
set "TASK2=ProductExport_Days2"
set "CMD2=%PYTHON% \"%SCRIPT%\" --headless --days 2"

schtasks /create /f ^
 /tn "%TASK2%" ^
 /tr "%CMD2%" ^
 /sc daily ^
 /st 12:50 ^
 /ru SYSTEM ^
 >> "%LOG_DIR%\setup.log" 2>&1

if %errorlevel% neq 0 (
    echo [%DATE% %TIME%] Task2 create FAILED >> "%LOG_DIR%\setup_error.log"
) else (
    echo [%DATE% %TIME%] Task2 create OK >> "%LOG_DIR%\setup.log"
)

echo Done
pause
endlocal
