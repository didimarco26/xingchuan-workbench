#!/bin/bash
# 星川 KOC 主页作品抓取助手 · Mac 一键启动（v1.0.0）
# 仅监听 127.0.0.1:7844；用系统 Chrome 打开公开移动版主页，不登录、不绕控。
cd "$(dirname "$0")" 2>/dev/null || exit 1

# 去除下载隔离属性（静默失败可忽略）
xattr -dr com.apple.quarantine . 2>/dev/null

# 定位 Node 18+：内置 node-runtime → Apple Silicon brew → Intel brew → PATH
find_node() {
  if [ -x "./node-runtime/bin/node" ]; then echo "./node-runtime/bin/node"; return; fi
  for c in /opt/homebrew/bin/node /usr/local/bin/node; do
    if [ -x "$c" ]; then echo "$c"; return; fi
  fi
  command -v node 2>/dev/null
}
NODE_BIN="$(find_node)"
if [ -z "$NODE_BIN" ]; then
  echo ""
  echo "未检测到 Node.js（需 18+）。"
  echo "请先安装 Node：终端执行 brew install node，或访问 https://nodejs.org 下载 LTS 版后重新双击本文件。"
  echo ""
  read -n1 -r -p "按回车关闭…"; exit 1
fi
echo "使用 Node: $NODE_BIN ($("$NODE_BIN" -v))"
DIR="$(dirname "$NODE_BIN")"
export PATH="$DIR:$PATH"

# 首次运行安装依赖（用系统 Chrome，跳过 Playwright 浏览器下载）
if [ ! -d node_modules ]; then
  echo "首次运行，安装依赖（约 30 秒）…"
  PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=1 npm i --no-audit --no-fund || {
    echo "依赖安装失败，请检查网络后重试。"; read -n1 -r -p "按回车关闭…"; exit 1; }
fi

echo "启动本机抓取助手 http://127.0.0.1:7844 （关闭本窗口即停止）"
exec "$NODE_BIN" xc-koc-fetch.js
