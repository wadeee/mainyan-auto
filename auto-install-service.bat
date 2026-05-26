@echo off
chcp 65001 >nul
setlocal

echo ================================================
echo   NSSM 服务安装脚本
echo ================================================

:: 检查管理员权限，如果需要则自我提升
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo 需要管理员权限，正在请求提升...
    powershell -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)
echo 已获取管理员权限。
echo.

:: ------------------------------------------------
:: 1. 停止并移除旧的 mainyan_daily_task 服务
:: ------------------------------------------------
echo [1/2] 清理旧的 mainyan_daily_task 服务...
nssm stop mainyan_daily_task >nul 2>&1
nssm remove mainyan_daily_task confirm >nul 2>&1
echo 等待服务完全卸载...
timeout /t 3 /nobreak >nul 2>&1
echo 旧服务已清理。
echo.

:: ------------------------------------------------
:: 2. 安装并启动 mainyan_daily_task 服务
:: ------------------------------------------------
echo [2/2] 安装并启动 mainyan_daily_task 服务...
nssm install mainyan_daily_task "C:\mainyan-auto\mainyan_daily_task.bat"
if %errorlevel% neq 0 (
    echo 服务安装失败！请确认 nssm 已正确安装。
    pause
    exit /b 1
)

nssm start mainyan_daily_task
if %errorlevel% neq 0 (
    echo 服务启动失败！
    pause
    exit /b 1
)

:: ------------------------------------------------
:: 完成
:: ------------------------------------------------
echo.
echo ================================================
echo   服务安装完成！已注册并启动。
echo ================================================
pause
endlocal