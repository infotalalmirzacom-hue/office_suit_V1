@echo off
title Office Suite EXE Builder
cd /d "%~dp0"

echo Python check ho raha hai...
where py >nul 2>nul
if %errorlevel%==0 (
    set PY=py
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set PY=python
    ) else (
        echo Python installed nahi hai.
        echo Pehle https://www.python.org/downloads/ se Python install karein.
        pause
        exit /b
    )
)

echo PyInstaller install ho raha hai...
%PY% -m pip install --upgrade pyinstaller

echo OfficeSuite.exe ban rahi hai...
%PY% -m PyInstaller --onefile --name OfficeSuite office_server.py

echo.
if exist "dist\OfficeSuite.exe" (
    echo SUCCESS: dist\OfficeSuite.exe ban gayi hai.
    echo OfficeSuite.exe ko OfficeSuite.html ke sath ek hi folder mein rakhein.
    explorer dist
) else (
    echo ERROR: EXE nahi bani. Upar ki error message dekhein.
)
pause
