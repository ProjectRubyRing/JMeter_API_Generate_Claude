@echo off
rem ===========================================================================
rem  jmeter_api_tool.bat - launcher for jmeter_api_tool.ps1 (Windows 11)
rem  Runs the PowerShell script with "-ExecutionPolicy Bypass" for this process
rem  only, so it works even when running .ps1 files is restricted.
rem  Usage: jmeter_api_tool.bat help
rem ===========================================================================
setlocal
set "JAT_PS1=%~dp0jmeter_api_tool.ps1"
if not exist "%JAT_PS1%" (
  echo [ERROR] jmeter_api_tool.ps1 not found: "%JAT_PS1%"
  exit /b 1
)
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%JAT_PS1%" %*
exit /b %ERRORLEVEL%
