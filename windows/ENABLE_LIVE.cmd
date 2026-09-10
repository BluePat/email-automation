@echo off
setlocal
set "PWSH=pwsh.exe"
where pwsh.exe >nul 2>&1
if errorlevel 1 set "PWSH=%ProgramFiles%\PowerShell\7\pwsh.exe"
if not exist "%PWSH%" where pwsh.exe >nul 2>&1 || goto :missing
"%PWSH%" -NoLogo -NoProfile -File "%~dp0Enable-SiolaLive.ps1"
set "RC=%errorlevel%"
echo.
pause
exit /b %RC%
:missing
echo PowerShell 7 nebyl nalezen. Spustte znovu INSTALL.cmd.
pause
exit /b 1
