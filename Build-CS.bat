@echo off
title Build RoboSync.exe (C#)
echo.
echo  Compiling RoboSync.exe with built-in C# compiler...
echo  =====================================================

set CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
set SRC=%~dp0RoboSync.cs
set OUT=%~dp0RoboSync.exe

if not exist "%CSC%" (
    echo  ERROR: csc.exe not found at %CSC%
    pause & exit /b 1
)

"%CSC%" /target:winexe /out:"%OUT%" /win32icon:"%~dp0RoboSync.ico" /r:System.Windows.Forms.dll /r:System.Drawing.dll /optimize+ "%SRC%"

if errorlevel 1 (
    echo.
    echo  ERROR: Compilation failed. See output above.
    pause & exit /b 1
)

:: Write version.txt with current git commit SHA (if git is available)
for /f %%i in ('git -C "%~dp0" rev-parse HEAD 2^>nul') do set GIT_SHA=%%i
if defined GIT_SHA (
    echo %GIT_SHA%> "%~dp0version.txt"
    echo  Version: %GIT_SHA%
) else (
    echo  (git not found - version.txt not written)
)

echo.
echo  =====================================================
echo   Done!  RoboSync.exe is ready at:
echo   %OUT%
echo  =====================================================
echo.
pause
