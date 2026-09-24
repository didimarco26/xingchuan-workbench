#!/bin/bash
# ============================================================
# 星川服务商达人打标工具 · Mac 一键启动（v4.1.0）
# ★ 零安装：无需自己安装 Node.js —— 系统没有 Node 时，本脚本会
#   按你的芯片（Apple Silicon / Intel）自动下载官方 Node 运行环境
#   （约 37MB，仅首次运行需要，从国内可直连的火山引擎 TOS 下载，
#    失败自动切换 npmmirror / nodejs.org 官方源），不需要管理员密码。
# 运行依赖（node_modules）已内置，解压即用。
# 登录优先自动打开系统 Chrome / Edge 扫码；企业环境禁止启动浏览器
# 时，会自动改为在工作台网页显示登录二维码，手机抖音扫码即可。
# ============================================================

cd "$(dirname "$0")" || exit 1
RUN_DIR="$(pwd)"
APP_VERSION="4.1.0"
NODE_VERSION="v22.23.2"
NODE_MAJOR_MIN=18
RUNTIME_DIR="$RUN_DIR/node-runtime"

# 下载源（按顺序尝试）。XC_NODE_RUNTIME_URL 可整体覆盖（测试/内网用）。
TOS_BASE="https://magic-builder.tos-cn-beijing.volces.com/xingchuan/node-runtime"
NPMMIRROR_BASE="https://registry.npmmirror.com/-/binary/node/${NODE_VERSION}"
NODEJS_BASE="https://nodejs.org/dist/${NODE_VERSION}"

echo "============================================"
echo "   星川打标工具启动中... v${APP_VERSION}"
echo "   运行目录：$RUN_DIR"
echo "============================================"
echo ""

# 1) 先清除 macOS 隔离属性（解压/下载的文件可能被 Gatekeeper 拦截，
#    必须在执行内置 Node 之前处理）
xattr -dr com.apple.quarantine "$RUN_DIR" >/dev/null 2>&1 || true

# 校验某个 node 可执行文件版本是否 >= NODE_MAJOR_MIN
node_version_ok() {
  local nbin="$1"
  [ -x "$nbin" ] || return 1
  "$nbin" -e 'const m=process.versions.node.split(".")[0];process.exit(Number(m)>=Number(process.env.XC_NODE_MAJOR_MIN||18)?0:1)' >/dev/null 2>&1
}

# 2) 依次找：内置运行环境 -> 常见绝对路径（Homebrew/官方pkg）-> PATH
NODE_BIN=""
CANDIDATES=("$RUNTIME_DIR/node" "/opt/homebrew/bin/node" "/usr/local/bin/node")
PATH_NODE="$(command -v node 2>/dev/null)"
[ -n "$PATH_NODE" ] && CANDIDATES+=("$PATH_NODE")
for cand in "${CANDIDATES[@]}"; do
  if [ -n "$cand" ] && node_version_ok "$cand"; then
    NODE_BIN="$cand"
    break
  fi
done

# 3) 都没有 -> 自动下载与芯片匹配的官方 Node（仅首次）
if [ -z "$NODE_BIN" ]; then
  ARCH_RAW="$(uname -m 2>/dev/null)"
  case "$ARCH_RAW" in
    arm64) ARCH="arm64" ;;
    x86_64) ARCH="x64" ;;
    *) ARCH="x64" ;;
  esac

  echo "ℹ️  未检测到 Node.js（${NODE_MAJOR_MIN}+）。"
  echo "   本工具可自动下载官方 Node ${NODE_VERSION}（${ARCH}，约 37MB，仅首次需要，无需管理员密码）。"
  echo ""
  echo "   8 秒后自动开始下载；或直接按回车立即开始；输入 n 取消。"
  echo "   （取消后也可自行从 https://nodejs.org 安装 Node，再重新双击本脚本）"
  read -t 8 -r ANSWER || ANSWER=""
  if [ "$ANSWER" = "n" ] || [ "$ANSWER" = "N" ]; then
    echo "已取消。安装 Node ${NODE_MAJOR_MIN}+ 后重新双击本脚本即可：https://nodejs.org"
    echo ""
    read -n 1 -r -p "按任意键关闭窗口..."
    exit 1
  fi
  echo ""

  ZIP_NAME="node-mac-${ARCH}-${NODE_VERSION}.zip"
  TARBALL_NAME="node-${NODE_VERSION}-darwin-${ARCH}.tar.gz"
  TMP_ZIP="$RUN_DIR/.xc-node-download"
  rm -rf "$TMP_ZIP"
  mkdir -p "$TMP_ZIP"

  # 源 1：TOS（zip，国内直连）。可用 XC_NODE_RUNTIME_URL 覆盖。
  URL1="${XC_NODE_RUNTIME_URL:-${TOS_BASE}/${ZIP_NAME}}"
  # 源 2/3：npmmirror / nodejs.org（tar.gz）
  URL2="${NPMMIRROR_BASE}/${TARBALL_NAME}"
  URL3="${NODEJS_BASE}/${TARBALL_NAME}"

  DOWNLOAD_OK=""
  echo "➡️  正在从火山引擎 TOS 下载（国内直连）..."
  if curl -fL --retry 2 --connect-timeout 15 --progress-bar -o "$TMP_ZIP/pkg" "$URL1" && [ -s "$TMP_ZIP/pkg" ]; then
    if (cd "$TMP_ZIP" && unzip -q -o pkg) && [ -x "$TMP_ZIP/node-runtime/node" ]; then
      mkdir -p "$RUNTIME_DIR"
      cp "$TMP_ZIP/node-runtime/node" "$RUNTIME_DIR/node"
      chmod +x "$RUNTIME_DIR/node"
      DOWNLOAD_OK="tos"
    fi
  fi

  if [ -z "$DOWNLOAD_OK" ]; then
    echo ""
    echo "➡️  TOS 未成功，切换 npmmirror 镜像下载..."
    if curl -fL --retry 2 --connect-timeout 15 --progress-bar -o "$TMP_ZIP/pkg.tar.gz" "$URL2" && [ -s "$TMP_ZIP/pkg.tar.gz" ]; then
      if (cd "$TMP_ZIP" && tar xzf pkg.tar.gz) && [ -x "$TMP_ZIP/node-${NODE_VERSION}-darwin-${ARCH}/bin/node" ]; then
        mkdir -p "$RUNTIME_DIR"
        cp "$TMP_ZIP/node-${NODE_VERSION}-darwin-${ARCH}/bin/node" "$RUNTIME_DIR/node"
        chmod +x "$RUNTIME_DIR/node"
        DOWNLOAD_OK="npmmirror"
      fi
    fi
  fi

  if [ -z "$DOWNLOAD_OK" ]; then
    echo ""
    echo "➡️  切换 Node.js 官方源下载..."
    if curl -fL --retry 2 --connect-timeout 15 --progress-bar -o "$TMP_ZIP/pkg.tar.gz" "$URL3" && [ -s "$TMP_ZIP/pkg.tar.gz" ]; then
      if (cd "$TMP_ZIP" && tar xzf pkg.tar.gz) && [ -x "$TMP_ZIP/node-${NODE_VERSION}-darwin-${ARCH}/bin/node" ]; then
        mkdir -p "$RUNTIME_DIR"
        cp "$TMP_ZIP/node-${NODE_VERSION}-darwin-${ARCH}/bin/node" "$RUNTIME_DIR/node"
        chmod +x "$RUNTIME_DIR/node"
        DOWNLOAD_OK="nodejs.org"
      fi
    fi
  fi

  rm -rf "$TMP_ZIP"

  if [ -z "$DOWNLOAD_OK" ] || ! node_version_ok "$RUNTIME_DIR/node"; then
    echo ""
    echo "❌ 自动下载未成功（可能网络受限）。请手动安装 Node.js ${NODE_MAJOR_MIN}+："
    echo "   https://nodejs.org （下载 macOS 安装包，双击安装后重新双击本脚本）"
    osascript -e 'display dialog "自动下载运行环境未成功，请手动安装 Node.js：" & return & "https://nodejs.org" & return & return & "安装完成后重新双击本脚本即可。" with title "星川打标工具" buttons {"确定"} default button 1 with icon caution' >/dev/null 2>&1
    echo ""
    read -n 1 -r -p "按任意键关闭窗口..."
    exit 1
  fi
  NODE_BIN="$RUNTIME_DIR/node"
  xattr -dr com.apple.quarantine "$RUNTIME_DIR" >/dev/null 2>&1 || true
  echo ""
  echo "✅ Node 运行环境已就绪（$DOWNLOAD_OK）。"
  echo ""
fi

echo "✅ Node：$("$NODE_BIN" -v 2>/dev/null)（$NODE_BIN）"
echo "✅ 运行依赖已内置，无需联网安装。"
echo "ℹ️  登录优先使用系统 Chrome / Edge；若被企业策略拦截，会自动改用网页二维码。"
echo ""

# 4) 启动工具（后台常驻；点「登录星图」时自动打开系统浏览器或在网页显示二维码）
echo "=================================================="
echo "  星川打标工具已启动 v${APP_VERSION}（本地服务常驻）"
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
XC_NODE_MAJOR_MIN="$NODE_MAJOR_MIN" "$NODE_BIN" xc-tagger.js

echo ""
echo "工具已退出，按任意键关闭窗口..."
read -n 1 -r
