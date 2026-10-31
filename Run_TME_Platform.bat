@echo off
setlocal enabledelayedexpansion
title Cancer Research: TME Analysis Platform Launcher

echo ====================================================================
echo   Cancer Research: TME Surprisal Analysis & ES Platform
echo   Connecting to GitHub Repository & Launching Web Workbench
echo ====================================================================
echo.

:: 1. Verify Python installation
set "PY_CMD="
where python >nul 2>nul
if %errorlevel% equ 0 (
    set "PY_CMD=python"
) else (
    where py >nul 2>nul
    if %errorlevel% equ 0 (
        set "PY_CMD=py"
    )
)

if "%PY_CMD%"=="" (
    echo [ERROR] Python was not found on this machine!
    echo Please install Python 3.10 or higher from:
    echo   https://www.python.org/downloads/
    echo (Make sure to check "Add Python to PATH" during installation)
    echo.
    pause
    exit /b 1
)

:: 2. Target Git repository configuration
set "REPO_URL=https://github.com/natiBirhauz/cancer-research-final-project.git"
set "ZIP_URL=https://github.com/natiBirhauz/cancer-research-final-project/archive/refs/heads/main.zip"

:: Check if already inside the project folder
if exist "%~dp0run_gui.py" (
    set "TARGET_DIR=%~dp0"
    cd /d "!TARGET_DIR!"
    goto :sync_and_run
)

:: Check if a local existing project folder exists right next to this script
if exist "%~dp0פרויקט גמר\run_gui.py" (
    set "TARGET_DIR=%~dp0פרויקט גמר"
    cd /d "!TARGET_DIR!"
    goto :sync_and_run
)

if exist "%~dp0cancer-research-final-project\run_gui.py" (
    set "TARGET_DIR=%~dp0cancer-research-final-project"
    cd /d "!TARGET_DIR!"
    goto :sync_and_run
)

:: If repository does not exist, download or clone from GitHub
set "TARGET_DIR=%~dp0cancer-research-final-project"
echo [1/3] Cloning latest algorithm repository from GitHub...
where git >nul 2>nul
if !errorlevel! equ 0 (
    git clone !REPO_URL! "!TARGET_DIR!"
) else (
    echo [INFO] Git CLI not found. Downloading archive directly from GitHub...
    powershell -Command "Invoke-WebRequest -Uri '!ZIP_URL!' -OutFile '!TARGET_DIR!.zip'; Expand-Archive -Path '!TARGET_DIR!.zip' -DestinationPath '%~dp0' -Force; Remove-Item '!TARGET_DIR!.zip' -Force; if (Test-Path '%~dp0cancer-research-final-project-main') { Rename-Item '%~dp0cancer-research-final-project-main' 'cancer-research-final-project' -Force }"
)

cd /d "!TARGET_DIR!"

:sync_and_run
:: Pull latest updates from Git if Git is installed
if exist ".git" (
    where git >nul 2>nul
    if !errorlevel! equ 0 (
        echo [2/3] Checking for latest algorithm updates from Git...
        git pull --quiet origin main >nul 2>nul
    )
)

:: 3. Verify Python packages
echo [3/3] Verifying dependencies (numpy, pandas, openpyxl, gradio)...
!PY_CMD! -c "import gradio, pandas, openpyxl, numpy" >nul 2>nul
if !errorlevel! neq 0 (
    echo Installing required packages (numpy, pandas, openpyxl, gradio)...
    !PY_CMD! -m pip install -q numpy pandas openpyxl gradio
)

:: 4. Launch Web GUI
echo.
echo ====================================================================
echo   Launching Workbench... Opening browser at http://127.0.0.1:7860
echo ====================================================================
echo.
!PY_CMD! run_gui.py

if !errorlevel! neq 0 (
    echo.
    echo [ERROR] The application terminated unexpectedly with code: !errorlevel!
    echo.
    pause
)
