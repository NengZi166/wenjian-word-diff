@echo off
setlocal

set "CSC=%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if exist "%CSC%" goto compiler_found

set "CSC=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if exist "%CSC%" goto compiler_found

echo The Windows .NET Framework C# compiler was not found.
exit /b 1

:compiler_found
if not exist "%~dp0..\launcher\bin" mkdir "%~dp0..\launcher\bin"

"%CSC%" /nologo /target:exe /platform:anycpu /optimize+ /out:"%~dp0..\launcher\bin\WenjianLauncher.exe" "%~dp0..\launcher\WenjianLauncher.cs"
if errorlevel 1 exit /b %ERRORLEVEL%

echo Wenjian native launcher is ready.
exit /b 0
