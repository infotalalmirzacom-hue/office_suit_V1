@echo off
title Office Suite
cd /d "%~dp0"
where py >nul 2>nul
if %errorlevel%==0 (
    py office_server.py
) else (
    python office_server.py
)
echo.
echo Office Suite band ho gaya hai.
pause
