@echo off
chcp 65001 >nul
cd /d "%~dp0"

call C:\ProgramData\anaconda3\condabin\conda.bat activate mainyan-auto

python mainyan_turnover_statistics.py --headless --date 2026.06.01 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.02 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.03 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.04 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.05 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.06 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.07 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.08 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.09 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.10 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.11 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.12 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.13 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.14 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.15 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.16 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.17 %*
python mainyan_turnover_statistics.py --headless --date 2026.06.18 %*
