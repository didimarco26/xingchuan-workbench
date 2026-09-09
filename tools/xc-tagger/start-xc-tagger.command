#!/bin/bash
# ============================================================
# 星川服务商达人打标工具 · Mac 一键启动（v3.7.1）
# 双击本文件即可运行。运行依赖（node_modules）已内置，解压即用。
# 登录时优先自动打开系统 Chrome / Edge 扫码；企业环境禁止启动浏览器
# 时，会自动改为在工作台网页里显示登录二维码，手机抖音扫码即可。
# 前提：已安装 Node.js 18+（https://nodejs.org）与 Chrome 或 Edge。
# ============================================================

cd "$(dirname "$0")" || exit 1
RUN_DIR="$(pwd)"

echo "============================================"
echo "   星川打标工具启动中... v3.7.1"
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

# 2) 清除 macOS 隔离属性（解压后内置依赖可能被 Gatekeeper 拦截）
xattr -dr com.apple.quarantine "$RUN_DIR" >/dev/null 2>&1 || true

echo "✅ 运行依赖已内置，无需联网安装。"
echo "ℹ️  登录将优先使用系统 Chrome / Edge；若被企业策略拦截，会自动改用网页二维码。"
echo ""

# 3) 启动工具（后台常驻；点「登录星图」时自动打开系统浏览器或在网页显示二维码）
echo "=================================================="
echo "  星川打标工具已启动 v3.7.1（本地服务常驻）"
echo "  本地服务：http://127.0.0.1:7842"
echo "  工作台网页：https://didimarco26.github.io/xingchuan-workbench/"
echo ""
echo "  使用 3 步："
echo "   1) 保持本窗口打开；"
echo "   2) 在工作台网页点「🚀 登录星图」——自动打开 Chrome/Edge 扫码，"
echo "      或网页直接显示二维码用手机抖音扫码（仅需一次，无需插件）；"
echo "   3) 上传达人名单 Excel，点「开始打标」等待结果。"
echo ""
echo "  使用过程中请保持本窗口打开，关闭窗口即停止工具"
echo "=================================================="
echo ""
node xc-tagger.js

echo ""
echo "工具已退出，按任意键关闭窗口..."
read -n 1
