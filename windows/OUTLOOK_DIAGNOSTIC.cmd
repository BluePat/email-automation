@echo off
setlocal
where pwsh.exe >nul 2>&1
if errorlevel 1 goto :missing_pwsh
pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Test-SiolaOutlookDiagnostic.ps1"
set "RESULT=%ERRORLEVEL%"
echo.
if not "%RESULT%"=="0" echo Diagnostika skoncila chybou. Poslete vytvoreny outlook-diagnostic soubor.
pause
exit /b %RESULT%
:missing_pwsh
echo Chybi PowerShell 7. Spustte nejdrive INSTALL.cmd.
pause
exit /b 1
