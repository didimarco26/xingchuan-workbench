#!/bin/bash
# ============================================================
# 星川服务商达人打标工具 · Mac 一键启动（v3.1 离线版）
# 双击本文件即可运行。离线包已内置 Node 依赖与 Playwright Chromium
# 浏览器（Intel / Apple Silicon 均含），解压即用、无需联网安装。
# 前提：已安装 Node.js 18+，https://nodejs.org
# 工具后台常驻；在工作台网页点「登录星图」会弹出 Chrome 窗口扫码，登录后自动关窗。
# ============================================================

cd "$(dirname "$0")" || exit 1
RUN_DIR="$(pwd)"

echo "============================================"
echo "   星川打标工具启动中... v3.1（离线版）"
echo "   运行目录：$RUN_DIR"
echo "============================================"
echo ""

# 1) 检测 Node.js
if ! command -v node >/dev/null 2>&1; then
  echo "❌ 未检测到 Node.js，请先安装 Node.js：https://nodejs.org"
  echo "   安装完成后，重新双击本脚本即可。"
  osascript -e 'display dialog "未检测到 Node.js，请先安装 Node.js：" & return & "https://nodejs.org" & return & return & "安装完成后重新双击本脚本即可。" with title "星川打标工具" buttons {"确定"} default button 1 with icon caution' >/dev/null 2>&1
  echo ""
  echo "按任意键关闭窗口..."
  read -n 1
  exit 1
fi

# 2) 使用包内自带的 Playwright Chromium（按芯片架构自动选择，全程不联网下载）
ARCH="$(uname -m)"
if [ "$ARCH" = "arm64" ] || [ "$ARCH" = "aarch64" ]; then
  BROWSERS_DIR="$RUN_DIR/browsers-arm64"
else
  BROWSERS_DIR="$RUN_DIR/browsers-x64"
fi
export PLAYWRIGHT_BROWSERS_PATH="$BROWSERS_DIR"
export PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1

CHROME_BIN="$BROWSERS_DIR/chromium-1148/chrome-mac/Chromium.app/Contents/MacOS/Chromium"
if [ ! -f "$CHROME_BIN" ]; then
  echo "❌ 内置 Chromium 浏览器文件缺失："
  echo "   $CHROME_BIN"
  echo "   请重新下载完整的 Mac 一键启动包，解压时不要跳过/删除任何文件。"
  osascript -e 'display dialog "内置浏览器文件缺失，请重新下载完整的 Mac 一键启动包（解压时保留全部文件）。" with title "星川打标工具" buttons {"确定"} default button 1 with icon caution' >/dev/null 2>&1
  echo ""
  echo "按任意键关闭窗口..."
  read -n 1
  exit 1
fi

# 3) 清除 macOS 隔离属性（浏览器从下载的压缩包解压后可能被 Gatekeeper 拦截弹窗）
xattr -dr com.apple.quarantine "$RUN_DIR" >/dev/null 2>&1 || true

echo "✅ 内置 Chromium 已就绪（$ARCH），无需联网安装。"
echo ""

# 4) 启动工具（后台常驻；点「登录星图」时弹窗扫码）
echo "=================================================="
echo "  星川打标工具已启动 v3.1（本地服务常驻）"
echo "  本地服务：http://127.0.0.1:7842"
echo "  工作台网页：https://didimarco26.github.io/xingchuan-workbench/"
echo ""
echo "  使用 3 步："
echo "   1) 保持本窗口打开；"
echo "   2) 在工作台网页点「🚀 登录星图」，本机弹出 Chrome 窗口，"
echo "      扫码登录后窗口自动关闭（仅需一次，无需任何插件）；"
echo "   3) 上传达人名单 Excel，点「开始打标」等待结果。"
echo ""
echo "  使用过程中请保持本窗口打开，关闭窗口即停止工具"
echo "=================================================="
echo ""
node xc-tagger.js

echo ""
echo "工具已退出，按任意键关闭窗口..."
read -n 1
