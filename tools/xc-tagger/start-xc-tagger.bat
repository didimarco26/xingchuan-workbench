@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

echo ============================================
echo   Xingchuan Tagger - starting...
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

rem 2) First run: install dependencies automatically (Playwright Chromium is downloaded by npm postinstall)
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

rem 3) Ensure Playwright Chromium browser (idempotent; ~150MB download on first run, skipped when installed)
echo Checking Playwright browser...
call npx playwright install chromium

rem 4) Start the tool
echo.
echo Starting xc-tagger on http://127.0.0.1:7842 ...
echo Keep this window open while using the tool. Close the window to stop.
echo.
node xc-tagger.js

echo.
echo Tool exited. Press any key to close.
pause >nul
