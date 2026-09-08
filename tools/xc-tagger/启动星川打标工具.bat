@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo    星川打标工具启动中...
echo ============================================
echo.

rem 1) 检测 Node.js
where node >nul 2>nul
if errorlevel 1 (
  echo [错误] 未检测到 Node.js，请先安装 Node.js：https://nodejs.org
  echo        安装完成后，重新双击本脚本即可。
  mshta "javascript:new ActiveXObject('WScript.Shell').Popup('未检测到 Node.js，请先安装 Node.js：\nhttps://nodejs.org\n\n安装完成后重新双击本脚本即可。',0,'星川打标工具',48);close()" 2>nul
  echo.
  pause
  exit /b 1
)

rem 2) 首次运行自动安装依赖（含 Playwright Chromium 浏览器）
if not exist node_modules (
  echo 首次运行，正在自动安装依赖（npm install），请耐心等待...
  call npm install
  if errorlevel 1 (
    echo.
    echo [错误] 依赖安装失败，请检查网络后重新双击本脚本。
    pause
    exit /b 1
  )
  echo.
  echo 正在安装 Playwright Chromium 浏览器（仅首次）...
  call npx playwright install chromium
  if errorlevel 1 (
    echo.
    echo [错误] 浏览器安装失败，请检查网络后重新双击本脚本。
    pause
    exit /b 1
  )
)

rem 3) 启动工具
echo.
echo 依赖就绪，启动工具中...
echo （使用过程中请保持本窗口不要关，关闭窗口即停止工具）
echo.
node xc-tagger.js

echo.
echo 工具已退出，按任意键关闭窗口...
pause >nul
