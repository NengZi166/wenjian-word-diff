@echo off
setlocal
chcp 65001 >nul
title Wenjian - Word document comparison

set "SERVER_SCRIPT=%~dp0serve-offline.ps1"
if not exist "%SERVER_SCRIPT%" goto missing_files

call :run_powershell "%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
if "%ERRORLEVEL%"=="0" exit /b 0

call :run_powershell "%SystemRoot%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
if "%ERRORLEVEL%"=="0" exit /b 0

where pwsh.exe >nul 2>&1
if errorlevel 1 goto powershell_failed
pwsh.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SERVER_SCRIPT%"
if "%ERRORLEVEL%"=="0" exit /b 0

:powershell_failed
echo.
echo Wenjian could not start Windows PowerShell on this computer.
echo Please contact IT or run serve-offline.ps1 with an available PowerShell.
echo.
pause
exit /b 1

:missing_files
echo.
echo Offline site files are incomplete. Please extract the whole ZIP again.
echo.
pause
exit /b 1

:run_powershell
if not exist "%~1" exit /b 1
"%~1" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SERVER_SCRIPT%"
exit /b %ERRORLEVEL%
