@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo ================================================
echo   打包 factory 目录到 ..\mainyan-auto.zip
echo   业务功能的 bat（xxx+1日.bat、xxx_2026-06-18.bat 等）会被排除
echo ================================================
echo.

:: ------------------------------------------------
:: 1. 依赖检查
:: ------------------------------------------------
where git >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 git，请先安装 Git for Windows 并加入 PATH。
    goto :fail
)

where zip >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 zip.exe，请把 zip 所在目录加入 PATH（例如 C:\zip）。
    goto :fail
)

:: ------------------------------------------------
:: 2. 定位 Git 自带的 bash.exe
::    （不要用 C:\Windows\System32\bash.exe，那是 WSL 的 bash）
:: ------------------------------------------------
set "GIT_BASH="
for /f "delims=" %%p in ('git --exec-path 2^>nul') do for %%b in ("%%~p\..\..\..\bin\bash.exe") do if exist "%%~fb" set "GIT_BASH=%%~fb"
if not defined GIT_BASH if exist "%ProgramFiles%\Git\bin\bash.exe" set "GIT_BASH=%ProgramFiles%\Git\bin\bash.exe"
if not defined GIT_BASH if exist "%ProgramFiles(x86)%\Git\bin\bash.exe" set "GIT_BASH=%ProgramFiles(x86)%\Git\bin\bash.exe"
if not defined GIT_BASH if exist "%LOCALAPPDATA%\Programs\Git\bin\bash.exe" set "GIT_BASH=%LOCALAPPDATA%\Programs\Git\bin\bash.exe"

if not defined GIT_BASH (
    echo [错误] 没找到 Git Bash（bash.exe），请确认已安装 Git for Windows。
    goto :fail
)

echo [1/2] Git Bash: %GIT_BASH%
echo [2/2] 执行 zip.sh 打包...
echo.

:: zip 打印的 "adding:" 行用系统 ANSI 编码，在 UTF-8 控制台下是乱码，
:: 所以先落到日志里，过滤掉这些行，只显示脚本自己的中文摘要
set "ZIP_LOG=%TEMP%\mainyan_zip.log"
"%GIT_BASH%" -c "bash ./zip.sh" > "%ZIP_LOG%" 2>&1
set "RC=%ERRORLEVEL%"
if exist "%ZIP_LOG%" (
    findstr /v /c:"adding:" /c:"updating:" "%ZIP_LOG%"
    del "%ZIP_LOG%" >nul 2>&1
)
if not "%RC%"=="0" goto :fail

echo.
for %%z in ("%~dp0..\mainyan-auto.zip") do (
    echo 完成！已生成: %%~fz
    echo 文件大小: %%~zz 字节
)
echo.
pause
exit /b 0

:fail
echo.
echo 打包失败，请查看上方错误信息。
echo.
pause
exit /b 1
