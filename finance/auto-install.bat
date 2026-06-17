@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

echo ================================================
echo   环境一键准备脚本
echo ================================================

:: ------------------------------------------------
:: 1. 检查是否已有 Python 3.11
:: ------------------------------------------------
echo.
echo [1/5] 检查 Python 环境...

set "PYTHON_CMD="

:: 先尝试 py launcher（Windows 官方安装器会注册）
py -3.11 --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py -3.11"
    goto :python_found
)

:: 再尝试 python 命令
python --version 2>nul | findstr /C:"3.11" >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    goto :python_found
)

:: 没有找到 Python 3.11，需要下载安装
echo 未检测到 Python 3.11，开始下载安装...
echo.

set "PYTHON_INSTALLER=%TEMP%\python-3.11.9-amd64.exe"
set "PYTHON_URL=https://www.python.org/ftp/python/3.11.9/python-3.11.9-amd64.exe"

:: 用 curl 下载（Windows 10+ 自带 curl）
echo 正在下载 Python 3.11.9 安装包...
curl -L -o "%PYTHON_INSTALLER%" "%PYTHON_URL%"
if errorlevel 1 (
    echo 下载失败！请检查网络连接，或手动下载：
    echo %PYTHON_URL%
    pause
    exit /b 1
)

echo 正在安装 Python 3.11.9（静默安装，包含 pip，添加到 PATH）...
"%PYTHON_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_launcher=1
if errorlevel 1 (
    echo 自动安装失败，正在打开安装界面，请手动安装...
    echo 注意：请勾选 "Add Python to PATH"！
    "%PYTHON_INSTALLER%"
    pause
    echo 安装完成后请重新运行本脚本。
    exit /b 1
)

echo Python 3.11.9 安装完成！

:: 刷新当前 session 的 PATH
set "LOCALAPPDATA_PYTHON=%LOCALAPPDATA%\Programs\Python\Python311"
set "PATH=%LOCALAPPDATA_PYTHON%;%LOCALAPPDATA_PYTHON%\Scripts;%PATH%"

:: 再次检测
py -3.11 --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=py -3.11"
    goto :python_found
)

python --version 2>nul | findstr /C:"3.11" >nul 2>&1
if not errorlevel 1 (
    set "PYTHON_CMD=python"
    goto :python_found
)

echo 安装后仍无法检测到 Python 3.11。
echo 请关闭此窗口，重新打开命令行后再运行本脚本。
pause
exit /b 1

:python_found
for /f "tokens=*" %%v in ('!PYTHON_CMD! --version 2^>^&1') do set "PY_VER=%%v"
echo 已找到: !PY_VER!  (命令: !PYTHON_CMD!)

:: ------------------------------------------------
:: 2. 生成 config.env（记录当前机器的绝对路径）
:: ------------------------------------------------
echo.
echo [2/5] 生成 config.env...

for /f "tokens=*" %%p in ('!PYTHON_CMD! -c "import sys; print(sys.executable)"') do set "PYTHON_FULL_PATH=%%p"
set "PW_BROWSERS=%LOCALAPPDATA%\ms-playwright"

(
    echo PYTHON_PATH=!PYTHON_FULL_PATH!
    echo PLAYWRIGHT_BROWSERS_PATH=!PW_BROWSERS!
) > "%~dp0config.env"

echo   PYTHON_PATH=!PYTHON_FULL_PATH!
echo   PLAYWRIGHT_BROWSERS_PATH=!PW_BROWSERS!
echo   已写入: %~dp0config.env

:: ------------------------------------------------
:: 3. 升级 pip
:: ------------------------------------------------
echo.
echo [3/5] 升级 pip...
!PYTHON_CMD! -m pip install --upgrade pip
if errorlevel 1 (
    echo pip 升级失败，继续安装依赖...
)

:: ------------------------------------------------
:: 4. 安装 requirements.txt 依赖
:: ------------------------------------------------
echo.
echo [4/5] 安装项目依赖 (requirements.txt)...
!PYTHON_CMD! -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
    echo 依赖安装失败！请检查网络或手动运行:
    echo   !PYTHON_CMD! -m pip install -r requirements.txt
    pause
    exit /b 1
)

:: ------------------------------------------------
:: 5. 安装 Playwright Chromium
:: ------------------------------------------------
echo.
echo [5/5] 安装 Playwright Chromium 浏览器内核...
!PYTHON_CMD! -m playwright install chromium
if errorlevel 1 (
    echo Chromium 安装失败！请手动运行:
    echo   !PYTHON_CMD! -m playwright install chromium
    pause
    exit /b 1
)

:: ------------------------------------------------
:: 完成
:: ------------------------------------------------
echo.
echo ================================================
echo   环境准备完成！
echo.
echo   Python:  !PY_VER!
echo   依赖:    已安装
echo   浏览器:  Chromium 已就绪
echo ================================================
pause
