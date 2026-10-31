@echo off
title Cancer Research TME Analysis Workbench
cd /d "%~dp0"

echo ====================================================================
echo  Launching Cancer Research TME Analysis Workbench...
echo  Opening browser interface at http://127.0.0.1:7860
echo ====================================================================
echo.

:: Verify Python is available
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python was not found in your system PATH!
    echo Please make sure Python is installed and added to PATH.
    echo.
    pause
    exit /b 1
)

:: Launch the Web application
python run_gui.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] The application terminated with an error code: %errorlevel%
    echo.
    pause
)
