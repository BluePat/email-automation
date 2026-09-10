@echo off
setlocal
set "PWSH=pwsh.exe"
where pwsh.exe >nul 2>&1
if not errorlevel 1 goto :run_installer
echo PowerShell 7 is required. Windows may ask for administrator approval.
winget install --id Microsoft.PowerShell --source winget --accept-package-agreements --accept-source-agreements
if errorlevel 1 goto :failed
set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" goto :failed
:run_installer
"%PWSH%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-SiolaAutomation.ps1"
if errorlevel 1 goto :failed
echo.
echo Installation preparation finished. Read ..\docs\INSTALACE_CZ.md before enabling LIVE.
pause
exit /b 0
:failed
echo.
echo Installation failed. No LIVE sending was enabled.
pause
exit /b 1
