@echo off
REM ============================================================
REM  Package the portable app into a shareable ZIP.
REM  Run this on YOUR machine, then send the ZIP to colleagues.
REM  (No PyInstaller / no giant exe - dependencies install in-app.)
REM
REM  NOTE: ships "PMIS Data.template.xlsx" (headers only), never the
REM  real "PMIS Data.xlsx" - that holds participant personal data
REM  (names, phone numbers, Aadhaar, bank accounts). The app creates a
REM  working copy from the template on first launch.
REM ============================================================
setlocal
set OUT=PMIS-Automation-Tool.zip

echo Packaging application files into %OUT% ...
powershell -NoProfile -Command ^
  "$files = 'app.py','pmis_geo.py','geo_check.py','browsers.py','config.py','deps.py','data_file.py','run.bat','VERSION','PMIS Data.template.xlsx','README.md','README-for-colleagues.txt';" ^
  "$missing = $files | Where-Object { -not (Test-Path $_) };" ^
  "if ($missing) { Write-Host 'MISSING:' ($missing -join ', ') -ForegroundColor Red; exit 1 };" ^
  "if (Test-Path '%OUT%') { Remove-Item '%OUT%' -Force };" ^
  "Compress-Archive -Path $files -DestinationPath '%OUT%' -Force;" ^
  "Write-Host 'Created %OUT%'"
if errorlevel 1 goto :error

echo.
echo ============================================================
echo  Done. Share %OUT% with colleagues.
echo  They unzip it and double-click run.bat.
echo ============================================================
goto :eof

:error
echo.
echo PACKAGING FAILED - see the messages above.
exit /b 1
