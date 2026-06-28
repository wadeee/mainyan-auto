@echo off
chcp 65001 >nul
cd /d "%~dp0"

for /f "tokens=1,* delims==" %%a in (config.env) do if "%%a"=="PYTHON_PATH" set "PYTHON_PATH=%%b"

"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.01 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.02 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.03 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.04 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.05 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.06 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.07 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.08 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.09 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.10 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.11 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.12 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.13 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.14 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.15 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.16 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.17 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.18 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.19 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.20 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.21 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.22 %*
"%PYTHON_PATH%" mainyan_turnover_daily.py --headless --date 2026.06.23 %*
