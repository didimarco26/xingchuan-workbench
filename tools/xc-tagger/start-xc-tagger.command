#!/bin/bash
# ============================================================
# 星川服务商达人打标工具 · Mac 一键启动
# 双击本文件即可运行（前提：已安装 Node.js 18+，https://nodejs.org）
# ============================================================
cd "$(dirname "$0")" || exit 1

echo "============================================"
echo "   星川打标工具启动中..."
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

# 2) 首次运行自动安装依赖（npm postinstall 会自动下载 Playwright Chromium）
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

# 3) 启动工具
echo ""
echo "✅ 依赖就绪，启动工具中..."
echo "（使用过程中请保持本窗口打开，关闭窗口即停止工具）"
echo ""
node xc-tagger.js

echo ""
echo "工具已退出，按任意键关闭窗口..."
read -n 1
