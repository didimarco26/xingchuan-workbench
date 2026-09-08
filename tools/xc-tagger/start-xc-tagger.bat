@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ============================================
echo   Xingchuan Tagger v3 - starting...
echo   (runs locally; a Chrome window pops up
echo    only when you click "Login to Xingtu")
echo ============================================
echo.

rem 1) Check Node.js
where node >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Node.js not found. Please install Node.js from https://nodejs.org
    echo         After installation, double-click this file again.
    echo.
    pause
    exit /b 1
)

rem 2) First run: install dependencies automatically
if not exist node_modules (
    echo First run: installing dependencies, please wait...
    call npm install
    if errorlevel 1 (
        echo.
        echo [ERROR] Dependency installation failed. Check your network and double-click again.
        echo.
        pause
        exit /b 1
    )
)

rem 3) Ensure Playwright's OWN Chromium is installed (run every time - already installed = a few seconds)
echo Checking Playwright Chromium browser...
call npx playwright install chromium
if errorlevel 1 (
    echo.
    echo [ERROR] Chromium download failed. Check your network and double-click again.
    echo.
    pause
    exit /b 1
)
echo [OK] Chromium ready.

rem 4) Start the tool (stays in background; click "Login to Xingtu" on the web page to scan QR)
echo.
echo ==================================================
echo  Xingchuan Tagger is running v3 (local service).
echo  Run from      : %~dp0
echo  Local service : http://127.0.0.1:7842
echo  Workbench page: https://didimarco26.github.io/xingchuan-workbench/
echo.
echo  3 STEPS:
echo    1. Keep this window open;
echo    2. On the workbench page, click the "Login to Xingtu" button -
echo       a Chrome window pops up, scan the QR code to log in,
echo       the window closes automatically (one time only, no plugin needed);
echo    3. Upload your influencer Excel and click "Start Tagging".
echo.
echo  Keep this window open while using the tool.
echo  Close the window (or press Ctrl+C) to stop.
echo ==================================================
echo.
node xc-tagger.js

echo.
echo Tool exited. Press any key to close.
pause >nul
