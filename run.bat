@echo off
REM ============================================================
REM  PMIS Automation Tool - launcher
REM  Double-click this to start the app.
REM ============================================================
setlocal

REM Find Python.
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

REM Launch the app from this folder (no console window stays around).
start "" pythonw "%~dp0app.py"
exit /b 0
