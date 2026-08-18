@echo off
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 pgfn_interface.py
) else (
    python pgfn_interface.py
)
if errorlevel 1 pause
