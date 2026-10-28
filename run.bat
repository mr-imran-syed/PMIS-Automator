@echo off
REM ============================================================
REM  PMIS Automation Tool - launcher
REM  Double-click this to start the app.
REM ============================================================
setlocal

set "APP=%~dp0app\app.py"

REM The code lives in the 'app' folder next to this file.
if not exist "%APP%" (
    echo Could not find "%APP%".
    echo.
    echo The 'app' folder must stay next to run.bat.
    echo If you unzipped only part of the download, unzip it again.
    echo.
    pause
    exit /b 1
)

REM pythonw runs it without leaving a console window behind.
where pythonw >nul 2>nul
if not errorlevel 1 (
    start "" pythonw "%APP%"
    exit /b 0
)

where python >nul 2>nul
if errorlevel 1 (
    echo Python is not installed.
    echo.
    echo This tool needs Python 3.10 or newer.
    echo Opening the download page - install Python, tick
    echo "Add python.exe to PATH", then run this again.
    echo.
    start "" "https://www.python.org/downloads/"
    pause
    exit /b 1
)

REM No pythonw on this machine - use python and keep this window,
REM so that any start-up error stays readable instead of flashing past.
python "%APP%"
if errorlevel 1 pause
exit /b 0
