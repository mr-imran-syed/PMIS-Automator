@echo off
REM ============================================================
REM  Package the portable app into a shareable ZIP.
REM  Run this from the app folder, then send the ZIP to colleagues.
REM  (No PyInstaller / no giant exe - dependencies install in-app.)
REM
REM  What a colleague sees after unzipping:
REM      run.bat             <- double-click this
REM      Read me first.txt
REM      app\                <- the code; they never need to open it
REM
REM  NOTE: ships "PMIS Data.template.xlsx" (headers only), never the
REM  real "PMIS Data.xlsx" - that holds participant personal data
REM  (names, phone numbers, Aadhaar, bank accounts). The app creates a
REM  working copy from the template on first launch.
REM ============================================================
setlocal
set OUT=%~dp0..\PMIS-Automation-Tool.zip
set STAGE=%TEMP%\pmis-pack

echo Packaging application files ...
powershell -NoProfile -Command ^
  "$root = Resolve-Path '%~dp0..';" ^
  "$code = 'app.py','pmis_geo.py','geo_check.py','participant.py','shg.py','browsers.py','config.py','credstore.py','data_file.py','deps.py','settings.py','logo.py','VERSION','PMIS Data.template.xlsx','README.md';" ^
  "$top  = 'run.bat','Read me first.txt';" ^
  "$missing = @($code | ForEach-Object { Join-Path '%~dp0.' $_ }) + @($top | ForEach-Object { Join-Path $root $_ }) | Where-Object { -not (Test-Path $_) };" ^
  "if ($missing) { Write-Host 'MISSING:' ($missing -join ', ') -ForegroundColor Red; exit 1 };" ^
  "if (Test-Path '%STAGE%') { Remove-Item '%STAGE%' -Recurse -Force };" ^
  "New-Item -ItemType Directory '%STAGE%\app' -Force | Out-Null;" ^
  "$code | ForEach-Object { Copy-Item (Join-Path '%~dp0.' $_) (Join-Path '%STAGE%\app' $_) };" ^
  "$top  | ForEach-Object { Copy-Item (Join-Path $root $_) (Join-Path '%STAGE%' $_) };" ^
  "if (Test-Path '%OUT%') { Remove-Item '%OUT%' -Force };" ^
  "Compress-Archive -Path '%STAGE%\*' -DestinationPath '%OUT%' -Force;" ^
  "Remove-Item '%STAGE%' -Recurse -Force;" ^
  "Write-Host ('Created ' + (Resolve-Path '%OUT%'))"
if errorlevel 1 goto :error

echo.
echo ============================================================
echo  Done. Share PMIS-Automation-Tool.zip with colleagues.
echo  They unzip it and double-click run.bat.
echo ============================================================
goto :eof

:error
echo.
echo PACKAGING FAILED - see the messages above.
exit /b 1
