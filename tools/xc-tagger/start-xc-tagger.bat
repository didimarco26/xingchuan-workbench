@echo off
chcp 65001 >nul 2>&1
setlocal enabledelayedexpansion

set "SRCDIR=%~dp0"
set "RUNDIR=%~dp0"

echo ============================================
echo   Xingchuan Tagger - starting...
echo ============================================
echo.

rem --- 0) Path check: spaces/brackets in the folder path make Chromium --user-data-dir fail
rem     (browser exits immediately, login cookies cannot be saved). This usually happens when
rem     Windows renames a duplicate download to "xc-tagger-win-v1 (2)".
rem     We AUTO-COPY the tool to a clean path (%USERPROFILE%\xc-tagger, fallback C:\xc-tagger)
rem     and run it from there - no manual action needed.
rem     NOTE: goto labels are used instead of ( ... ) blocks because a ")" in the path
rem     would close a parenthesized block early.
set "BADP=0"
echo "%SRCDIR%" | findstr /C:" " >nul 2>&1
if not errorlevel 1 set "BADP=1"
echo "%SRCDIR%" | findstr /C:"(" >nul 2>&1
if not errorlevel 1 set "BADP=1"
echo "%SRCDIR%" | findstr /C:")" >nul 2>&1
if not errorlevel 1 set "BADP=1"
if "%BADP%"=="1" goto :migrate
goto :run

:migrate
set "TARGET=%USERPROFILE%\xc-tagger"
echo "%TARGET%" | findstr /C:" " >nul 2>&1
if not errorlevel 1 set "TARGET=C:\xc-tagger"
echo "%TARGET%" | findstr /C:"(" >nul 2>&1
if not errorlevel 1 set "TARGET=C:\xc-tagger"
echo.
echo ==================================================
echo  [NOTICE] Current folder path contains spaces or brackets:
echo     %SRCDIR%
echo  The browser cannot save login state in such a path.
echo  Auto-copying the tool to a clean path:
echo     %TARGET%
echo ==================================================
echo.
if not exist "%TARGET%" mkdir "%TARGET%"
rem /E = copy subdirs incl. empty; /XD excludes the old (unusable) browser profile
robocopy "%SRCDIR%." "%TARGET%" /E /XD .xc-chrome-profile /NFL /NDL /NJH /NJS /NC /NS /NP
if errorlevel 8 (
  echo.
  echo [ERROR] Auto-copy failed. Please MANUALLY move the xc-tagger folder
  echo         to a path without spaces or brackets, e.g. C:\xc-tagger
  echo         then double-click start-xc-tagger.bat again.
  echo.
  pause
  exit /b 1
)
set "RUNDIR=%TARGET%\"
echo [OK] Copied. Starting from %TARGET%
echo.

:run
cd /d "%RUNDIR%"

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

rem 3) Ensure Playwright Chromium browser (idempotent; ~150MB on first run, skipped when installed)
echo Checking Playwright browser...
call npx playwright install chromium

rem 4) Start the tool
echo.
echo ==================================================
echo  Xingchuan Tagger is running.
echo  Run from      : %RUNDIR%
echo  Local service : http://127.0.0.1:7842
echo  Workbench page: https://didimarco26.github.io/xingchuan-workbench/
echo  Keep this window open while using the tool.
echo  Close the window (or press Ctrl+C) to stop.
echo ==================================================
echo.
node xc-tagger.js

echo.
echo Tool exited. Press any key to close.
pause >nul
