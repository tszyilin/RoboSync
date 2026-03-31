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

"%CSC%" /target:winexe /out:"%OUT%" /optimize+ "%SRC%"

if errorlevel 1 (
    echo.
    echo  ERROR: Compilation failed. See output above.
    pause & exit /b 1
)

echo.
echo  =====================================================
echo   Done!  RoboSync.exe is ready at:
echo   %OUT%
echo  =====================================================
echo.
pause
