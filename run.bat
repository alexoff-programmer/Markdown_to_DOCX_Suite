@echo off
setlocal EnableExtensions
cd /d "%~dp0"
python -c "import sys" >nul 2>&1
if %errorlevel% equ 0 (
    python main.py %*
) else (
    py -3 -c "import sys" >nul 2>&1
    if %errorlevel% equ 0 (
        py -3 main.py %*
    ) else (
        echo [ERROR] Python 3.10+ not found in PATH.
        pause
        exit /b 1
    )
)
if %errorlevel% neq 0 pause
endlocal
