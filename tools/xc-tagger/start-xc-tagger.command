#!/bin/bash
# ============================================================
# 星川服务商达人打标工具 · Mac 一键启动
# 双击本文件即可运行（前提：已安装 Node.js 18+，https://nodejs.org）
# 工具后台常驻；在工作台网页点「登录星图」会弹出 Chrome 窗口扫码，登录后自动关窗。
# ============================================================

cd "$(dirname "$0")" || exit 1
RUN_DIR="$(pwd)"

echo "============================================"
echo "   星川打标工具启动中... v3"
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

# 2) 首次运行自动安装依赖
if [ ! -d node_modules ]; then
  echo "⏳ 首次运行，正在自动安装依赖（npm install），请耐心等待..."
  npm install
  if [ $? -ne 0 ]; then
    echo ""
    echo "❌ 依赖安装失败，请检查网络后重新双击本脚本。"
    echo "按任意键关闭窗口..."
    read -n 1
    exit 1
  fi
fi

# 3) 确保 Playwright 自带 Chromium 已安装（每次启动都执行，已装则秒过，未装则自动下载）
echo "⏳ 正在验证 Playwright Chromium 浏览器..."
npx playwright install chromium
if [ $? -ne 0 ]; then
  echo "⚠️ 浏览器下载失败，请检查网络后重新双击本脚本。"
  echo "按任意键关闭窗口..."
  read -n 1
  exit 1
fi
echo "   Chromium 已就绪。"

# 4) 启动工具（后台常驻；点「登录星图」时弹窗扫码）
echo ""
echo "=================================================="
echo "  星川打标工具已启动 v3（本地服务常驻）"
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
