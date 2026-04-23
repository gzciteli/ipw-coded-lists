@echo off
setlocal
cd /d "%~dp0"

title IPW Coded Lists

where py >nul 2>nul
if %errorlevel%==0 (
    py main.py
    goto :eof
)

where python >nul 2>nul
if %errorlevel%==0 (
    python main.py
    goto :eof
)

echo Python was not found on this computer.
echo Make sure Python is installed and available as "py" or "python".
pause
