@echo off
setlocal EnableExtensions
chcp 65001 >nul
title ClipboardDocxRunner Build
color 0B
cd /d "%~dp0"
echo ============================================================
echo   Building ClipboardDocxRunner.exe ...
echo   Working directory: %CD%
echo ============================================================
python -c "import sys" >nul 2>&1
if %errorlevel% equ 0 (
    python build.py
) else (
    py -3 -c "import sys" >nul 2>&1
    if %errorlevel% equ 0 (
        py -3 build.py
    ) else (
        color 0C
        echo [ERROR] Python 3.10+ not found in PATH.
        pause
        exit /b 1
    )
)
if %errorlevel% neq 0 (
    color 0C
    echo [ERROR] Build failed with code %errorlevel%.
    pause
    exit /b %errorlevel%
)
echo.
echo Press any key to exit...
pause >nul
endlocal
