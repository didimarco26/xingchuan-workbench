#!/bin/bash
# ============================================================
# 星川服务商达人打标工具 · Mac 一键启动
# 双击本文件即可运行（前提：已安装 Node.js 18+，https://nodejs.org）
# ============================================================

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
RUN_DIR="$SCRIPT_DIR"

# 0) 路径检测与自动迁移：
#    文件夹路径含空格/括号时（常见于 Mac 给重复下载自动加 " 2" 后缀），
#    Chromium 的 --user-data-dir 会解析失败、浏览器启动即退出，登录态无法保存。
#    此时自动把工具复制到干净路径 ~/xc-tagger 并从那里启动，用户无需手动操作。
if echo "$SCRIPT_DIR" | grep -qE '[ ()]'; then
  TARGET="$HOME/xc-tagger"
  if echo "$TARGET" | grep -qE '[ ()]'; then
    TARGET="/tmp/xc-tagger"   # 极端情况：用户主目录本身也含空格
  fi
  echo ""
  echo "=================================================="
  echo "  [注意] 当前工具所在路径含空格或括号："
  echo "    $SCRIPT_DIR"
  echo "  该路径下浏览器无法保存登录态，正在自动复制到干净路径："
  echo "    $TARGET"
  echo "=================================================="
  echo ""
  mkdir -p "$TARGET" || { echo "❌ 无法创建 $TARGET，请手动把 xc-tagger 文件夹移到不含空格的路径（如 ~/xc-tagger）后重试。"; read -n 1; exit 1; }
  if command -v ditto >/dev/null 2>&1; then
    ditto "$SCRIPT_DIR" "$TARGET" || { echo "❌ 复制失败，请手动移动文件夹后重试。"; read -n 1; exit 1; }
  else
    cp -R "$SCRIPT_DIR/." "$TARGET/" || { echo "❌ 复制失败，请手动移动文件夹后重试。"; read -n 1; exit 1; }
  fi
  rm -rf "$TARGET/.xc-chrome-profile"   # 旧路径下的 profile 本就无法保存登录态，不带走
  RUN_DIR="$TARGET"
  echo "✅ 已复制完成，将从 $RUN_DIR 启动工具。"
  echo ""
fi

cd "$RUN_DIR" || exit 1

echo "============================================"
echo "   星川打标工具启动中..."
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

# 3) 确保 Playwright Chromium 浏览器已安装（幂等：首次下载约 150MB，已安装则秒过）
echo "⏳ 检查 Playwright Chromium 浏览器（首次会下载约 150MB，请耐心等待，之后秒过）..."
npx playwright install chromium || echo "⚠️ 浏览器检查/下载失败，若稍后启动时报浏览器错误，请检查网络后重新双击本脚本。"

# 4) 启动工具
echo ""
echo "=================================================="
echo "  星川打标工具已启动"
echo "  本地服务：http://127.0.0.1:7842"
echo "  工作台网页：https://didimarco26.github.io/xingchuan-workbench/"
echo "  使用过程中请保持本窗口打开，关闭窗口即停止工具"
echo "=================================================="
echo ""
node xc-tagger.js

echo ""
echo "工具已退出，按任意键关闭窗口..."
read -n 1
