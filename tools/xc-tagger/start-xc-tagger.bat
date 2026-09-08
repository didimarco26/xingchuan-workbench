@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ============================================
echo   Xingchuan Tagger v3.5
echo   Node dependencies are bundled - no
echo   network install needed. Uses your system
echo   Chrome / Edge for login (falls back to an
echo   on-page QR code if the browser is blocked).
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

echo [OK] Node.js found. Dependencies are bundled, no network install needed.
echo [INFO] Login uses your system Chrome / Edge; if it is blocked by policy,
echo        an on-page QR code is shown automatically (scan with Douyin app).
echo.

rem 2) Start the tool (stays in background; click "Login to Xingtu" on the web page)
echo ==================================================
echo  Xingchuan Tagger is running v3.5 (local service).
echo  Run from      : %~dp0
echo  Local service : http://127.0.0.1:7842
echo  Workbench page: https://didimarco26.github.io/xingchuan-workbench/
echo.
echo  3 STEPS:
echo    1. Keep this window open;
echo    2. On the workbench page, click the "Login to Xingtu" button -
echo       a Chrome/Edge window opens for QR login; if your company
echo       policy blocks pop-up windows, a QR code is shown right on
echo       the web page instead (scan it with the Douyin phone app);
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
