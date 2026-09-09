#!/usr/bin/env node
/* eslint-disable */
/**
 * 星川服务商达人自助打标 · 本地一键工具 (xc-tagger) v3.7.1
 * ------------------------------------------------------------------
 * 用途：服务商在本机运行本工具，它会：
 *   1) 在 127.0.0.1:7842 起一个本地 HTTP 服务（只监听本机，不对外）；
 *   2) 浏览器方案（CDP 优先）：工具自动用【系统 Chrome/Edge】以远程调试端口
 *      启动独立实例并 connectOverCDP 附加，也会先附加已在运行的调试 Chrome
 *      （9222-9225 端口探测）。浏览器真实可见，服务商在窗口里登录。
 *   3) 登录（仅需一次）：网页点「🚀 登录星图」→ POST /login，工具每 3 秒
 *      主动轮询所有标签页 URL（纯 URL 判定：xingtu.cn 下路径落在服务商端
 *      /sup/ 段——含 /sup/ 控制台首页与 /sup/author/ 创作者主页——即成功；
 *      v3.6.2 起登录引导/检测全部走服务商端 /sup/，不再走广告主端 /ad/，
 *      避免两端 session 隔离导致 /sup/ 下无登录态），
 *      扫码/账号密码/验证码均可；成功后写 Cookie 备份、CDP 模式自动关闭窗口，
 *      /health 内存登录位即时生效。超时 5 分钟。
 *   4) 打标：POST /tag 时后台自动重开浏览器（同一配置目录免登录）：
 *      · ★ v3.7.2 主链路改为「服务商详情页直达」：直接冷开服务商端达人详情页
 *        /provider/pages/author/douyin/{达人ID}（服务商 /sup/ 会话下整页 goto 不被
 *        重定向，批量稳定；服务商广场点达人昵称打开的就是这个新 tab 页面），落地后点
 *        顶部 el-tabs「创作能力」tab，拦截 /gw/api/author/get_author_show_items_v2，
 *        取 data.latest_item_info（个人最新 15 条）前 3 条、不足补 latest_star_item_info
 *        （星图商单视频），按 item_id 拼 https://www.douyin.com/video/{id}；
 *        遇滑块验证码自动等待最多 30 秒让服务商手动拖过，超时降级空视频、不阻塞；
 *      · v3.7.2 详情页打不开时依次降级：① 服务商广场站内导航（搜索→点卡片 SPA 跳
 *        /ad/creator/market/detail/{id}，冷开会被踢到 /ad/creator/index，仅作兜底）；
 *        ② 冷开 /sup/author/detail/{id} 至少拿达人资料（该页拿不到创作能力视频）；
 *      · 拦截 XHR/fetch 响应（拦截不到降级解析 DOM），一次拿到达人资料（星川
 *        等级 S0-S5 / 交付项目数 / 星图消耗 / 电商等级 / 粉丝数 / 内容主题标签）
 *        与【前三个视频】的标题/链接/播放/点赞/内容形式，输出 videoAnalysis；
 *        视频标题反哺五大受控标签（人设/内容形式/画面风格/拍摄场景/行业，只补不覆盖）；
 *      · v3.7.0 输出完整达人标签体系：达人人设 / 内容形式 / 画面风格 / 拍摄场景 /
 *        行业 5 大多选标签（对齐存量达人表字段），并新增 视频链接1-3、主要带货类目、
 *        近期爆款内容方向、打标置信度（高/中/低）字段，支持导出/回填存量达人库；
 *      · 达人主页两条路都进不去（ID 无效 / 未入驻星图 / 被重定向）才降级到广场
 *        用昵称/ID 搜索兜底；
 *      · XC_TAGGER_DEBUG=1 时落截图+接口 JSON+DOM 状态到 .xc-debug/
 *        供真机校准；每达人间隔 800ms 防风控。
 *   5) 评分双库口径：
 *      · 存量库（名单「来源」列=存量/库存/已合作，stock）：80 分制四项
 *        （星川等级30/交付20/星图消耗20/电商10），与 tblUXi2raUVtv3et 对齐；
 *      · 增量库（increment）/未在库新达人（new，缺省）：基础 80 分 +
 *        视频分析加分（带货视频+5/个最多15、口播测评种草+5、播放超10万+5，
 *        合计最多+20，总分封顶100）；输出 dataSource/videoBonus/videoAnalysis。
 *   网页接口：GET /health · GET /login/qr · POST /login · POST /parse · POST /tag。
 *   兜底：系统浏览器不可用/调试端口被禁 → 无头扫码模式（网页显示抖音二维码）。
 *   Cookie 只留在服务商本机，不上传、不落库。
 *
 * 运行：解压后双击 start-xc-tagger.command(Mac) / start-xc-tagger.bat(Windows)
 *      （node_modules 已内置，解压即用；需 Node.js 18+ 与 Chrome 或 Edge）。
 *
 * 安全：CORS 仅放行 *.github.io 与本机页面；服务只绑定 127.0.0.1。
 */

'use strict';

const path = require('path');

// 说明：打标使用【无头 Chromium + Cookie 文件】；登录时临时启动【有界面 Chromium】
// 弹窗扫码，登录成功后 Cookie 落盘到 .xc-cookies.json，窗口自动关闭。不使用
// --user-data-dir 持久化目录，工具所在路径含空格/括号也能正常运行。

const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const fs = require('fs');

// ---- 配置 ----------------------------------------------------------------
const PORT = 7842;
const HOST = '127.0.0.1';
// 登录态文件：保存星图 Cookie（Playwright 格式 JSON 数组）。无头 Chromium 启动时加载，
// 检测到已登录会话时自动回写刷新。含登录凭证，严禁提交 / 外传（已在 .gitignore 忽略）。
const COOKIES_FILE = path.resolve(__dirname, '.xc-cookies.json');
// 巨量星图 · 达人广场（广告主侧）。抓取接口与该页同源(www.xingtu.cn)，注入 Cookie 后天然带登录态。
const XINGTU_ORIGIN = 'https://www.xingtu.cn';
// v3.6.2：服务商端控制台首页 = 登录引导页/登录态判定页。
// 服务商账号必须从 /sup/ 端发起登录：未登录访问 /sup/ 会被星图重定向到 SSO，
// 登录成功后 SSO 回跳 /sup/，建立的是【服务商端】会话；/ad/ 广告主端与 /sup/
// 服务商端 session 相互隔离，若从 /ad/ 登录，/sup/author/detail 下仍无登录态。
const SUP_URL = 'https://www.xingtu.cn/sup/';
// v3.6.1：达人广场改为新路径 /pro/ad/pages/market（旧 /ad/creator/square 在普通账号下已重定向/失效）。
// 仅作为昵称搜索兜底页（v3.6.2 起不再用于登录引导/登录态判定）。
const SQUARE_URL = 'https://www.xingtu.cn/pro/ad/pages/market';
// v3.7.1：服务商广场（站内搜索达人 → 点卡片 SPA 跳 /ad/creator/market/detail/{id}，
// 「创作能力」tab 的视频列表 XHR 只在这条站内导航链路下才会触发）。
// 首选 /provider/pages/market，打不开回退 /pro/ad/pages/market。
const PROVIDER_MARKET_URL = 'https://www.xingtu.cn/provider/pages/market';
// 登录直达页（v3.6.2）：直接导航服务商端控制台 /sup/。未登录时星图自动 302 到
// sso.oceanengine.com 统一登录页（redirect_uri 由星图按服务商端编码，role 正确），
// 登录成功后回跳 /sup/——天然建立服务商端会话。无头扫码同理：goto 后落在 SSO 页点抖音图标。
const LOGIN_URL = SUP_URL;
// 固定一个常见桌面 UA，避免无头浏览器被星图风控识别
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36';
const DEBUG_DUMP = process.env.XC_TAGGER_DEBUG === '1'; // 调试：把原始响应落盘

// 仅允许以下来源的网页调用本地服务（安全限制）
const ALLOWED_ORIGINS = [
  'https://didimarco26.github.io',
  'https://xingchuan-advisor.surge.sh',
  'http://localhost:5173',
  'http://127.0.0.1:5173',
  'http://localhost:8000',
  'null', // 部分浏览器 file:// Origin 为 'null'
];

// ---- 打标规则（与存量达人逻辑一致）----------------------------------------
// 星川等级 S（30 分）
const S_LEVEL_SCORE = { S5: 30, S4: 25.5, S3: 21, S2: 15, S1: 9, S0: 4.5 };
// 电商等级 L（10 分）
function scoreEcomLevel(lv) {
  const n = parseInt(String(lv || '').replace(/[^0-9]/g, ''), 10);
  if (isNaN(n)) return 0;
  if (n >= 3) return 10; // L3 / L4 / L5
  if (n === 2) return 7;
  if (n === 1) return 4;
  return 1; // L0
}
// 交付项目数（20 分）
function scoreDeliveries(n) {
  n = Number(n) || 0;
  if (n > 50) return 20;
  if (n >= 21) return 15;
  if (n >= 6) return 10;
  if (n >= 1) return 5;
  return 0;
}
// 星图消耗（20 分，单位：元）
function scoreConsumption(yuan) {
  yuan = Number(yuan) || 0;
  if (yuan > 1000000) return 20;   // >100 万
  if (yuan >= 100000) return 16;  // 10 万 ~ 100 万
  if (yuan >= 10000) return 12;   // 1 万 ~ 10 万
  if (yuan >= 1000) return 8;     // 0.1 万 ~ 1 万
  return 0;
}
// 分层（按总分）
function tierOf(score) {
  if (score >= 60) return { tier: '标杆', medal: '🥇' };
  if (score >= 40) return { tier: '主力', medal: '🥈' };
  if (score >= 20) return { tier: '潜力', medal: '🥉' };
  return { tier: '储备', medal: '⚪' };
}
function fmtWan(yuan) {
  yuan = Number(yuan) || 0;
  if (yuan >= 10000) return (yuan / 10000).toFixed(yuan >= 1000000 ? 0 : 1) + '万';
  return yuan > 0 ? yuan + '元' : '0';
}

// ---- 浏览器会话管理 ---------------------------------------------------------
// 两级方案，自动降级，全程中文提示：
//   ① CDP 模式（优先）：工具自动用【系统 Chrome/Edge】以远程调试端口启动一个
//      独立实例（独立 user-data-dir，与用户日常浏览器互不干扰），再
//      connectOverCDP 附加。浏览器真实可见，服务商在窗口里直接扫码登录。
//      —— Windows 安全策略会立即关闭 Playwright 自己 launch 的 Chromium
//         窗口（exitCode=0），但系统 Chrome 由本工具以普通进程方式拉起，不受影响。
//   ② 无头扫码模式（兜底）：系统浏览器全部不可用 / 调试端口被企业策略禁用时，
//      Playwright 以 headless 启动（优先系统 Chrome/Edge，再兜底内置 Chromium），
//      工具截取星图登录页二维码，服务商在工作台网页上用手机抖音 App 扫码，
//      无头会话完成登录。只要有网就能跑。
const CDP_CANDIDATE_PORTS = [9222, 9223, 9224, 9225];
// 调试浏览器的独立配置目录（与启动参数约定一致；路径允许含中文/空格）
const CHROME_PROFILE_DIR = path.resolve(__dirname, '.xc-chrome-profile');

let _session = null;          // { mode:'cdp'|'headless', browser, context, browserName, browserId, cdpPort, proc, _workPage }
let _sessionAcquiring = null;

/**
 * 探测本机可用浏览器（纯函数，按优先级返回【已存在】的候选列表，便于单测注入）。
 * 优先级：系统 Chrome → Chrome Beta → Chrome Canary → Microsoft Edge → 其他 Chromium。
 * @param {{platform?:string, env?:object, exists?:(p:string)=>boolean, which?:(n:string)=>?string}} opt
 */
function browserCandidates(opt = {}) {
  const platform = opt.platform || process.platform; // win32 | darwin | linux
  const env = opt.env || process.env;
  const exists = opt.exists || ((p) => { try { return !!p && fs.existsSync(p); } catch (_) { return false; } });
  const out = [];
  const pushFirst = (id, name, kind, paths) => {
    for (const p of paths) { if (p && exists(p)) { out.push({ id, name, kind, exe: p }); return; } }
  };
  if (platform === 'win32') {
    const pf = env.ProgramFiles || 'C:\\Program Files';
    const pf86 = env['ProgramFiles(x86)'] || env.ProgramFilesX86 || 'C:\\Program Files (x86)';
    const local = env.LOCALAPPDATA || '';
    const win = (...a) => a.filter(Boolean).join('\\');
    pushFirst('chrome', '系统 Google Chrome', 'chrome', [
      win(pf, 'Google', 'Chrome', 'Application', 'chrome.exe'),
      win(pf86, 'Google', 'Chrome', 'Application', 'chrome.exe'),
      win(local, 'Google', 'Chrome', 'Application', 'chrome.exe'),
    ]);
    pushFirst('chrome-beta', 'Google Chrome Beta', 'chrome', [
      win(pf, 'Google', 'Chrome Beta', 'Application', 'chrome.exe'),
      win(pf86, 'Google', 'Chrome Beta', 'Application', 'chrome.exe'),
    ]);
    pushFirst('chrome-canary', 'Google Chrome Canary', 'chrome', [
      win(local, 'Google', 'Chrome SxS', 'Application', 'chrome.exe'),
    ]);
    pushFirst('edge', 'Microsoft Edge', 'edge', [
      win(pf86, 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
      win(pf, 'Microsoft', 'Edge', 'Application', 'msedge.exe'),
    ]);
    pushFirst('chromium', 'Chromium', 'chromium', [
      win(local, 'Chromium', 'Application', 'chrome.exe'),
    ]);
  } else if (platform === 'darwin') {
    const home = env.HOME || '';
    const mac = (app, bin) => [
      `/Applications/${app}.app/Contents/MacOS/${bin}`,
      home ? `${home}/Applications/${app}.app/Contents/MacOS/${bin}` : '',
    ];
    pushFirst('chrome', '系统 Google Chrome', 'chrome', mac('Google Chrome', 'Google Chrome'));
    pushFirst('chrome-beta', 'Google Chrome Beta', 'chrome', mac('Google Chrome Beta', 'Google Chrome Beta'));
    pushFirst('chrome-canary', 'Google Chrome Canary', 'chrome', mac('Google Chrome Canary', 'Google Chrome Canary'));
    pushFirst('edge', 'Microsoft Edge', 'edge', mac('Microsoft Edge', 'Microsoft Edge'));
    pushFirst('chromium', 'Chromium', 'chromium', mac('Chromium', 'Chromium'));
  } else {
    // Linux / 其他：PATH 查找
    const which = opt.which || ((name) => {
      try {
        const r = require('child_process').execSync(`command -v ${name} 2>/dev/null`, { timeout: 3000 }).toString().trim();
        return r || null;
      } catch (_) { return null; }
    });
    const bins = [
      ['chrome', '系统 Google Chrome', 'chrome', ['google-chrome-stable', 'google-chrome']],
      ['chrome-beta', 'Google Chrome Beta', 'chrome', ['google-chrome-beta']],
      ['chrome-canary', 'Google Chrome Canary', 'chrome', ['google-chrome-unstable']],
      ['edge', 'Microsoft Edge', 'edge', ['microsoft-edge-stable', 'microsoft-edge']],
      ['chromium', 'Chromium', 'chromium', ['chromium-browser', 'chromium']],
    ];
    for (const [id, name, kind, cmds] of bins) {
      for (const c of cmds) { const p = which(c); if (p) { out.push({ id, name, kind, exe: p }); break; } }
    }
  }
  return out;
}

// 端口是否空闲（纯逻辑包装，isFree 可注入便于单测）
function netTcpPortFree(port) {
  return new Promise((resolve) => {
    try {
      const srv = require('net').createServer();
      let done = false;
      const finish = (v) => { if (!done) { done = true; try { srv.close(); } catch (_) {} resolve(v); } };
      srv.once('error', () => finish(false));
      srv.once('listening', () => finish(true));
      srv.listen(port, '127.0.0.1');
      setTimeout(() => finish(false), 3000);
    } catch (_) { resolve(false); }
  });
}
async function pickFreePort(ports, isFree) {
  const check = isFree || netTcpPortFree;
  for (const p of ports) { if (await check(p)) return p; }
  return null;
}

// 探测 CDP 调试端口是否就绪（GET /json/version）
async function cdpProbe(port, timeoutMs = 20000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const r = await fetch(`http://${HOST}:${port}/json/version`, { signal: AbortSignal.timeout(1500) });
      if (r.ok) return await r.json();
    } catch (_) { /* 浏览器还在启动，下轮再试 */ }
    await sleep(500);
  }
  return null;
}

// 单次探测某端口是否已有 CDP 调试服务（不轮询，快速失败），返回 /json/version 信息或 null
async function cdpProbeOnce(port, timeoutMs = 1200) {
  try {
    const r = await fetch(`http://${HOST}:${port}/json/version`, { signal: AbortSignal.timeout(timeoutMs) });
    if (r.ok) return await r.json();
  } catch (_) { /* 端口无服务 */ }
  return null;
}

/**
 * ★ v3.3：附加到【已经在运行】的调试 Chrome/Edge（端口 9222-9225 逐个探测）。
 * 典型场景：上一次工具运行以 detached 方式拉起的 Chrome 在工具退出后仍活着，
 * 且持有 .xc-chrome-profile 配置锁与登录态；新 spawn 的实例会把命令行转发给它后
 * 立即退出，导致旧版误判「启动失败」并静默降级无头。直接复用该实例即可拿到
 * 用户刚登录的登录态。connectImpl/ probeImpl 可注入便于单测。
 */
async function connectExistingCdp(playwright, ports, probeImpl) {
  const list = ports || CDP_CANDIDATE_PORTS;
  const probe = probeImpl || cdpProbeOnce;
  for (const port of list) {
    let info = null;
    try { info = await probe(port); } catch (_) { info = null; }
    if (!info) continue;
    try {
      const browser = await playwright.chromium.connectOverCDP(`http://${HOST}:${port}`, { timeout: 6000 });
      const ctxs = browser.contexts ? browser.contexts() : [];
      const context = ctxs[0] || await browser.newContext();
      const brand = (info && info.Browser) ? String(info.Browser).split('/')[0] : 'Chrome/Edge';
      console.log(`   ✅ 已附加到正在运行的调试浏览器（CDP 端口 ${port}，${brand}），直接复用其登录态。`);
      return {
        mode: 'cdp', browser, context, cdpPort: port, proc: null, reused: true, _workPage: null,
        browserId: 'existing',
        browserName: `已运行的系统 ${/Edg/i.test(brand) ? 'Edge' : 'Chrome'}（复用窗口）`,
      };
    } catch (e) {
      console.log(`   ⚠️ 端口 ${port} 有调试服务但附加失败：${friendlyErr(e)}`);
    }
  }
  return null;
}

/**
 * 用指定浏览器启动一个带远程调试端口的独立实例。
 * spawn 以参数数组传参（不经 shell），路径含中文/空格也安全。
 * 成功返回 {proc, port, info}；失败（进程退出 / 端口无响应）返回 null 并清理子进程。
 */
async function launchSystemBrowser(cand, port, profileDir) {
  const { spawn } = require('child_process');
  const args = [
    `--remote-debugging-port=${port}`,
    `--user-data-dir=${profileDir}`,
    '--no-first-run',
    '--no-default-browser-check',
    '--disable-default-apps',
    '--disable-translate',
    SUP_URL,
  ];
  let proc = null;
  let settled = false;
  return new Promise((resolve) => {
    const startedAt = Date.now();
    try {
      proc = spawn(cand.exe, args, { detached: true, stdio: 'ignore' });
      proc.unref();
      proc.on('error', () => { if (!settled) { settled = true; resolve(null); } });
      proc.on('exit', () => {
        // 启动早期就退出（被安全策略/企业策略拦截）→ 判定失败
        if (!settled && Date.now() - startedAt < 25000) { settled = true; resolve(null); }
      });
    } catch (e) {
      settled = true; return resolve(null);
    }
    cdpProbe(port, 22000).then((info) => {
      if (settled) return;
      settled = true;
      if (info) resolve({ proc, port, info });
      else { try { proc.kill(); } catch (_) {} resolve(null); }
    });
  });
}

// ① CDP 模式：先附加已运行的调试浏览器；没有再逐个候选浏览器【启动→连接】，全部失败返回 null
async function tryCdpSession() {
  const playwright = require('playwright');

  // 0) 优先复用已在运行的调试 Chrome/Edge（含上一次运行残留、已登录的窗口）
  try {
    const existing = await connectExistingCdp(playwright);
    if (existing) return existing;
  } catch (e) {
    console.log('   ⚠️ 探测已有调试浏览器异常：' + friendlyErr(e));
  }

  const cands = browserCandidates();
  if (!cands.length) {
    console.log('ℹ️  未探测到系统 Chrome/Edge，将使用无头扫码模式。');
    return null;
  }
  for (const cand of cands) {
    // 每个候选浏览器都重新挑端口：上一轮失败可能留下了半启动实例占着端口
    const port = await pickFreePort(CDP_CANDIDATE_PORTS);
    if (!port) {
      console.log('ℹ️  调试端口 9222-9225 均被占用，尝试附加已有实例…');
    } else {
      console.log(`🌐 尝试启动 ${cand.name}（调试端口 ${port}）…`);
      let launched = null;
      try { launched = await launchSystemBrowser(cand, port, CHROME_PROFILE_DIR); } catch (_) { launched = null; }
      if (launched) {
        try {
          const browser = await playwright.chromium.connectOverCDP(`http://${HOST}:${port}`, { timeout: 8000 });
          const ctxs = browser.contexts ? browser.contexts() : [];
          const context = ctxs[0] || await browser.newContext();
          console.log(`   ✅ 已连接 ${cand.name}（CDP 端口 ${port}）。`);
          return { mode: 'cdp', browser, context, browserName: cand.name, browserId: cand.id, cdpPort: port, proc: launched.proc, _workPage: null };
        } catch (e) {
          console.log(`   ✗ 连接 ${cand.name} 失败：${friendlyErr(e)}，尝试下一个浏览器。`);
          try { launched.proc.kill(); } catch (_) {}
        }
      } else {
        console.log(`   ✗ ${cand.name} 启动失败或调试端口无响应（常见于配置目录被已有调试浏览器占用）。`);
      }
    }
    // spawn 失败/端口被占：很可能命令行已被转发给一个持锁的旧实例 → 再试附加已有调试浏览器
    try {
      const existing = await connectExistingCdp(playwright);
      if (existing) return existing;
    } catch (_) { /* 继续下一候选 */ }
  }
  console.log('ℹ️  CDP 模式不可用，将使用无头扫码模式。');
  return null;
}

// ② 无头扫码模式：Playwright headless launch（优先系统浏览器的可执行文件，
//    系统浏览器全无时兜底 Playwright 内置 Chromium）
async function tryHeadlessSession() {
  const playwright = require('playwright');
  const baseArgs = ['--no-sandbox', '--disable-blink-features=AutomationControlled', '--disable-dev-shm-usage'];
  const cands = browserCandidates();
  for (const cand of cands) {
    try {
      const browser = await playwright.chromium.launch({
        headless: true,
        executablePath: cand.exe,
        args: baseArgs,
      });
      console.log(`🌐 无头模式：使用 ${cand.name}（后台运行，登录请用网页上的二维码扫码）。`);
      const context = await browser.newContext({ userAgent: UA, viewport: { width: 1280, height: 900 } });
      return { mode: 'headless', browser, context, browserName: cand.name + '（无头）', browserId: cand.id, cdpPort: null, proc: null, _workPage: null };
    } catch (e) {
      console.log(`   ✗ 无头启动 ${cand.name} 失败：${friendlyErr(e)}`);
    }
  }
  // 兜底：Playwright 内置 Chromium（若未下载会给出明确指引）
  try {
    const browser = await playwright.chromium.launch({ headless: true, args: baseArgs });
    const context = await browser.newContext({ userAgent: UA, viewport: { width: 1280, height: 900 } });
    console.log('🌐 无头模式：使用 Playwright 内置 Chromium。');
    return { mode: 'headless', browser, context, browserName: 'Playwright 内置 Chromium（无头）', browserId: 'bundled', cdpPort: null, proc: null, _workPage: null };
  } catch (e) {
    throw new Error(
      '本机未找到可用浏览器，且内置 Chromium 未就绪。请安装 Google Chrome 或 Microsoft Edge 后重试；' +
      '或在工具目录执行 npx playwright install chromium 后重启工具。原始错误：' + friendlyErr(e)
    );
  }
}

async function sessionAlive(sess) {
  try { return !!(sess && sess.browser && sess.browser.isConnected && sess.browser.isConnected()); }
  catch (_) { return false; }
}

/**
 * 获取（或复用）浏览器会话：CDP 优先，失败自动降级无头扫码。
 * 浏览器只启动一次；连接断开后自动重建。
 */
async function acquireSession() {
  if (_session && await sessionAlive(_session)) return _session;
  if (_sessionAcquiring) return _sessionAcquiring;
  _sessionAcquiring = (async () => {
    let sess = null;
    try { sess = await tryCdpSession(); } catch (e) { console.log('CDP 模式异常：' + friendlyErr(e)); }
    if (!sess) sess = await tryHeadlessSession();
    sess.browser.on('disconnected', () => {
      if (_session === sess) {
        console.log('⚠️  浏览器连接已断开，下次操作将自动重新启动。');
        _session = null;
      }
    });
    // 无头模式：注入备份 Cookie
    if (sess.mode === 'headless') {
      const saved = readCookieFile();
      if (saved.length) {
        const n = await addCookiesSafe(sess.context, saved);
        if (n > 0) console.log('✅ 已注入备份星图 Cookie（' + n + ' 条）。');
      }
    }
    _session = sess;
    return sess;
  })();
  try { return await _sessionAcquiring; } finally { _sessionAcquiring = null; }
}

/**
 * 在浏览器上下文里找一个可复用的标签页（v3.3：避免重复开 SSO 登录页）。
 * 优先级：已在星图服务商端 /sup/ 控制台的页 → 任意星图页 → 停在 SSO/登录页的页 → null。
 * v3.6.2：优先匹配 /sup/（服务商端），不再优先 /ad/（广告主端 session 与 /sup/ 隔离）。
 */
function findReusablePage(ctx) {
  const pages = ctx.pages ? ctx.pages() : [];
  const urlOf = (p) => { try { return p.url() || ''; } catch (_) { return ''; } };
  const alive = pages.filter(p => { try { return p && !p.isClosed(); } catch (_) { return false; } });
  return alive.find(p => { const u = urlOf(p); return /xingtu\.cn/.test(u) && /\/sup(?:\/|$)/.test(u); })
    || alive.find(p => /xingtu\.cn/.test(urlOf(p)))
    || alive.find(p => /oceanengine\.com|sso|login/i.test(urlOf(p)))
    || null;
}

/**
 * 获取会话内的星图工作标签页：复用已打开的星图/SSO 页，没有才新开；
 * 并确保最终落在星图服务商端 /sup/ 控制台（v3.6.2：打标主链路 /sup/author/detail
 * 与同源接口都依赖服务商端会话；不再收敛到广告主端 /ad/ 广场）。
 */
async function sessionWorkPage(sess) {
  if (sess._workPage && !sess._workPage.isClosed()) return sess._workPage;
  const ctx = sess.context;
  let page = findReusablePage(ctx);
  if (!page) page = await ctx.newPage();
  try { await page.setViewportSize({ width: 1280, height: 860 }); } catch (_) {}
  let url = '';
  try { url = page.url() || ''; } catch (_) {}
  // 不在星图服务商端 /sup/ 区域（SSO 登录页 / 营销首页 / about:blank）→ 导航到 /sup/ 控制台；
  // 已登录会直接渲染控制台（URL 留在 /sup/），未登录会被重定向到 SSO（checkLoggedIn 据此判定）。
  if (!(/xingtu\.cn/.test(url) && /\/sup(?:\/|$)/.test(url))) {
    try {
      await page.goto(SUP_URL, { waitUntil: 'domcontentloaded', timeout: 45000 });
      await page.waitForTimeout(2500); // 等 JS 渲染，避免误判登录态
    } catch (e) {
      console.log('⚠️ 打开星图异常（稍后自动重试）：' + friendlyErr(e));
    }
  } else {
    try { await page.bringToFront(); } catch (_) {}
  }
  sess._workPage = page;
  page.on('close', () => { if (sess._workPage === page) sess._workPage = null; });
  return page;
}

// ---- 无头模式：登录二维码抓取 ----------------------------------------------
let _qr = null; // { status:'starting'|'waiting'|'loggedin'|'expired', qrDataUrl, kind, updatedAt }
const QR_SELECTORS = [
  'img[src*="qrcode"]', 'img[src*="qr-code"]', 'img[src*="qrCode"]', 'img[src*="qr"]',
  'img[src*="douyin"]', 'img[src*="oauth"]',
  '[class*="qrcode"] img', '[class*="qr-code"] img', '[class*="qrCode"] img',
  '[class*="login-qr"] img', '[class*="scan"] img', '[class*="douyin"] img', '[class*="oauth"] img',
  'canvas[class*="qr"]', '[class*="qrcode"] canvas', '[class*="scan"] canvas', 'canvas',
];
async function captureLoginQr(page) {
  for (const sel of QR_SELECTORS) {
    try {
      const el = await page.$(sel);
      if (!el) continue;
      const box = await el.boundingBox().catch(() => null);
      // 二维码是较大的方形图（>=120px），过滤掉小图标（抖音/头条 logo 仅 24px）
      if (box && box.width >= 120 && box.height >= 120) {
        const buf = await el.screenshot({ type: 'png' });
        return { qrDataUrl: 'data:image/png;base64,' + buf.toString('base64'), kind: 'element' };
      }
    } catch (_) { /* 选择器不匹配，试下一个 */ }
  }
  // 兜底：截取右侧登录卡片区域（SSO 页二维码出现在点「其他方式-抖音」后的右侧卡片中部）
  try {
    const vp = page.viewportSize() || { width: 1360, height: 950 };
    const buf = await page.screenshot({
      type: 'png',
      clip: {
        x: Math.round(vp.width * 0.585), y: Math.round(vp.height * 0.30),
        width: Math.round(vp.width * 0.30), height: Math.round(vp.height * 0.42),
      },
    });
    return { qrDataUrl: 'data:image/png;base64,' + buf.toString('base64'), kind: 'clip' };
  } catch (_) { return null; }
}
async function refreshQr(page) {
  const shot = await captureLoginQr(page);
  if (shot) {
    _qr = { status: 'waiting', ...shot, updatedAt: Date.now() };
  } else if (_qr && _qr.status === 'waiting' && Date.now() - _qr.updatedAt > 60000) {
    _qr = { ..._qr, status: 'expired' };
  }
}

/**
 * 无头模式：导航到登录页并切到「抖音扫码」。
 * v3.6.2：goto 服务商端 /sup/（LOGIN_URL），未登录会被星图 302 到 SSO 统一登录页
 * （redirect_uri 回跳 /sup/，服务商角色正确）；落到 SSO 页后点「其他方式-抖音」图标
 * （.icon.douyin），右侧卡片即加载抖音扫码二维码。返回 true 表示已进入扫码流程。
 */
async function gotoDouyinQr(page) {
  try {
    await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
  } catch (e) {
    console.log('⚠️ 打开登录页超时（继续尝试）：' + friendlyErr(e));
  }
  await page.waitForTimeout(6000); // 等 SSO 页渲染
  // 点击「其他方式 - 抖音」图标（优先选择器，失败则坐标兜底）
  let clicked = false;
  try {
    const dy = await page.$('.icon.douyin, [class*="douyin"]');
    if (dy) { await dy.click({ timeout: 8000 }); clicked = true; }
  } catch (_) { /* 选择器点击失败，走坐标兜底 */ }
  if (!clicked) {
    try {
      const box = await page.evaluate(() => {
        const el = document.querySelector('.icon.douyin') ||
          Array.from(document.querySelectorAll('img,span,i,div')).find(e =>
            /douyin/i.test((e.className || '') + ' ' + (e.src || '')) && e.getBoundingClientRect().width > 15);
        if (!el) return null;
        const r = el.getBoundingClientRect();
        return { x: r.x + r.width / 2, y: r.y + r.height / 2 };
      });
      if (box) { await page.mouse.click(box.x, box.y); clicked = true; }
    } catch (_) { /* 坐标兜底也失败 */ }
  }
  console.log(clicked ? '🟢 已切换到抖音扫码登录，等待二维码加载…' : '⚠️ 未找到抖音扫码入口，展示登录页截图兜底。');
  await page.waitForTimeout(5000); // 等二维码加载
  return clicked;
}

// ---- Cookie 登录态（备份文件，主要登录态在浏览器 profile 内）----------------
function readCookieFile() {
  try {
    if (!fs.existsSync(COOKIES_FILE)) return [];
    const arr = JSON.parse(fs.readFileSync(COOKIES_FILE, 'utf8'));
    return Array.isArray(arr) ? arr : [];
  } catch (_) { return []; }
}
function writeCookieFile(cookies) {
  try {
    if (Array.isArray(cookies) && cookies.length) {
      fs.writeFileSync(COOKIES_FILE, JSON.stringify(cookies, null, 2), 'utf8');
    }
  } catch (_) { /* 落盘失败不影响主流程 */ }
}
// 逐条注入 Cookie，单条失败不影响其余，返回成功条数
async function addCookiesSafe(ctx, cookies) {
  let ok = 0;
  for (const ck of cookies) {
    try { await ctx.addCookies([ck]); ok++; } catch (_) { /* 跳过非法 / 过期项 */ }
  }
  return ok;
}

// ★ v3.4：纯 URL 登录判定（不读页面内容）。
// 规则（v3.6.2 修订）：URL 在 xingtu.cn 域名下，且路径落在【服务商端 /sup/ 段】即视为已登录。
// 依据：未登录访问任意 /sup/ 页（控制台首页 /sup/、创作者主页 /sup/author/detail/{id}）
// 都会被星图重定向到 sso.oceanengine.com 登录页（URL 离开 xingtu.cn）；只有服务商端会话
// 建立后才会停在 /sup/。v3.6.2 起【不再】把 /ad/（广告主端，含 /pro/ad/pages/market 广场）
// 当作登录成功——广告主端与服务商端 session 隔离，/ad/ 已登录不代表 /sup/ 有登录态，
// 旧逻辑正是「窗口里 /ad/ 已登录、/sup/author/detail 却一直未检索到」的根因。
// 另排除仍在登录链路上的 URL（SSO/护照/redirect_uri 参数/login 路径），双保险。
function isLoggedInUrl(url) {
  if (!url) return false;
  let host = '', path = '';
  try {
    const p = new URL(String(url));
    host = p.hostname || '';
    path = p.pathname || '';
  } catch (_) { return false; }
  if (!/(^|\.)xingtu\.cn$/i.test(host)) return false;      // 主机名必须是星图域名（含子域）
  // v3.6.2：登录区域 = 服务商端 /sup/ 段（含控制台首页 /sup/ 与创作者主页 /sup/author/detail/{id}）。
  // 广告主端 /ad/（含达人广场 /pro/ad/pages/market）不再判为登录成功。
  if (!/^\/sup(?:\/|$)/i.test(path)) return false;
  if (/\/(sso|passport|login)(?:\/|$)/i.test(path)) return false; // 仍在登录链路
  return true;
}

// 页面级判定（供 /tag 等在 sessionWorkPage 导航到广场后调用）：只看该页 URL。
async function checkLoggedIn(page) {
  try {
    if (!page) return false;
    try { if (page.isClosed && page.isClosed()) return false; } catch (_) { return false; }
    return isLoggedInUrl(page.url() || '');
  } catch { return false; }
}

// ---- 登录（在已 CDP 连接的系统 Chrome 里打开星图标签页，等待扫码/账号密码登录）----
// /login 正在等待登录时为 true，/health 据此让前端显示「等待登录中…」。
let _loginPending = false;
// v3.4：内存登录位——登录成功后置 true（/health 即时返回 loggedIn，无需启动/导航浏览器）；
// 进程重启后由启动预检重新探测。
let _loggedIn = false;
const LOGIN_WAIT_MS = 300000; // 登录最长等待 5 分钟
const LOGIN_POLL_MS = 3000;   // v3.4：每 3 秒主动轮询一次所有标签页 URL

/**
 * 统一登录流程（自动适配浏览器会话模式）：
 * - CDP 模式：系统 Chrome/Edge 窗口已打开星图页，服务商在窗口里扫码；
 * - 无头模式：工具截取登录页二维码写入 _qr，前端 GET /login/qr 展示，
 *   服务商在网页上用手机抖音 App 扫码。
 * ★ v3.3：① 复用浏览器里已有的星图/SSO 标签页，不重复开页；
 *   ② 轮询时遍历【所有标签页】，任一标签进入星图服务商端 /sup/ 控制台即判成功
 *   （用户可能在启动时自带的那个标签页里登录，旧版只盯新开的页会漏判）；
 *   ③ 登录后落在营销首页等非 /sup/ 页时，自动导航到 /sup/ 控制台做最终复核；
 *   ④ 成功后工作页统一收敛到 /sup/ 控制台（服务商端会话），并全量回写 Cookie。
 * 返回 {ok, loggedIn, cancelled?, needInstall?, mode?, browserName?, message?}。
 * @param {{aborted:boolean}} signal 前端取消标志（请求中断时置 true）
 */
async function runLogin(signal) {
  let sess = null;
  try {
    sess = await acquireSession();
  } catch (e) {
    return { ok: false, loggedIn: false, needInstall: true, message: friendlyErr(e) };
  }

  // 登录成功收尾：把工作页收敛到服务商端 /sup/ 控制台、回写 Cookie、复位二维码状态
  const finishLogin = async (loggedPage) => {
    let work = loggedPage;
    try {
      const u = (() => { try { return work.url() || ''; } catch (_) { return ''; } })();
      if (!isLoggedInUrl(u)) {
        await work.goto(SUP_URL, { waitUntil: 'domcontentloaded', timeout: 45000 }).catch(() => {});
        await work.waitForTimeout(2500).catch(() => {});
      }
    } catch (_) { /* 导航失败不影响登录态判定 */ }
    sess._workPage = work;
    try { work.on('close', () => { if (sess._workPage === work) sess._workPage = null; }); } catch (_) {}
    writeCookieFile(await sess.context.cookies().catch(() => []));
    _qr = { status: 'loggedin', qrDataUrl: null, updatedAt: Date.now() };
  };

  // v3.4：枚举浏览器【全部上下文】（CDP 下可能含默认上下文之外的目标），
  // 供标签页扫描使用；兜底至少包含会话主上下文。
  const allContexts = () => {
    const out = [];
    try { if (sess.browser && typeof sess.browser.contexts === 'function') out.push(...sess.browser.contexts()); } catch (_) {}
    if (sess.context && !out.includes(sess.context)) out.push(sess.context);
    return out;
  };

  // v3.4：每 3 秒主动轮询——遍历所有上下文的所有标签页，【只看 URL、不读页面内容】：
  // 任一标签页落在 xingtu.cn 的服务商端 /sup/ 控制台即判定登录成功。
  const findLoggedInPage = async () => {
    for (const ctx of allContexts()) {
      let pages = [];
      try { pages = (ctx && typeof ctx.pages === 'function') ? ctx.pages() : []; } catch (_) { continue; }
      for (const pg of pages) {
        try {
          if (!pg || pg.isClosed()) continue;
          if (isLoggedInUrl(pg.url() || '')) return pg;
        } catch (_) { /* 页面跳转中，下轮再测 */ }
      }
    }
    return null;
  };

  // v3.4：登录成功后关闭可见浏览器窗口（仅 CDP 模式；无头模式无窗口、保留供打标）。
  // 登录态已持久化在 .xc-chrome-profile 配置目录（重开免登录）+ Cookie 备份文件。
  const closeLoginBrowser = async () => {
    if (sess.mode !== 'cdp') return;
    try { await sess.browser.close(); } catch (_) { /* CDP Browser.close，关闭整个浏览器 */ }
    try { if (sess.proc && !sess.proc.killed) sess.proc.kill(); } catch (_) { /* 兜底杀进程 */ }
    await sleep(1200);
    if (_session === sess) _session = null;
    sess._workPage = null;
    console.log('🪟 登录窗口已自动关闭（登录态已保存在本机，打标时会在后台自动重开浏览器）。');
  };

  // 已登录快速返回（任一标签页在服务商端 /sup/ 控制台）。此路径说明登录态此前已建立，
  // 保留现有浏览器会话供打标直接复用，不关窗、不重开。
  let already = await findLoggedInPage();
  if (already) {
    await finishLogin(already);
    _loggedIn = true;
    return { ok: true, loggedIn: true, mode: sess.mode, browserName: sess.browserName, message: '已是登录状态' };
  }

  // 取登录驱动页：优先复用已停在 SSO/星图页的标签（用户可见、且与启动时打开的是同一个），没有才新开
  let page = sess._workPage;
  if (!page || page.isClosed()) page = findReusablePage(sess.context);
  if (!page) {
    page = await sess.context.newPage().catch(() => null);
    if (!page) return { ok: false, loggedIn: false, message: '无法打开浏览器标签页，请重试。', mode: sess.mode };
    try { await page.setViewportSize({ width: 1360, height: 950 }); } catch (_) {}
  }
  sess._workPage = page;

  // 未登录：导航到服务商端 /sup/ 控制台（未登录会自动跳到 SSO，登录后回跳 /sup/）
  if (sess.mode === 'cdp') {
    try {
      const u = (() => { try { return page.url(); } catch (_) { return ''; } })();
      if (!/sso\.oceanengine\.com|\/sup(?:\/|$)/.test(u)) await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await page.bringToFront();
    } catch (_) { /* 导航失败也让用户在可见窗口里手动操作 */ }
    console.log('🟢 请在已打开的 ' + sess.browserName + ' 窗口中登录巨量星图（可抖音扫码或账号密码/手机验证码登录）…');
  } else {
    console.log('🟢 无头模式：正在打开抖音扫码登录，请在工作台网页用手机抖音 App 扫码…');
    _qr = { status: 'starting', qrDataUrl: null, updatedAt: Date.now() };
    await gotoDouyinQr(page);
  }

  const deadline = Date.now() + LOGIN_WAIT_MS;
  let lastReload = Date.now();
  let lastSquareCheck = 0;
  while (Date.now() < deadline) {
    if (signal && signal.aborted) {
      console.log('ℹ️  用户取消了登录等待。');
      return { ok: false, loggedIn: false, cancelled: true, mode: sess.mode, message: '已取消登录' };
    }
    if (!(await sessionAlive(sess))) {
      return { ok: false, loggedIn: false, message: '浏览器连接中断，请重新点「登录星图」（工具会自动重启浏览器）。', mode: sess.mode };
    }
    // 无头模式：持续刷新二维码；停留超过 100 秒未扫码则重新进抖音扫码换新码
    if (sess.mode === 'headless') {
      await refreshQr(page);
      if (Date.now() - lastReload > 100000 && _qr && _qr.status !== 'loggedin') {
        await gotoDouyinQr(page).catch(() => {});
        lastReload = Date.now();
      }
    }
    // ① 所有上下文的所有标签页里找已登录的（纯 URL 判定）
    let loggedPage = await findLoggedInPage();
    // ② CDP 落地页复核：某标签已到 xingtu.cn 但不在 /sup/（营销首页/角色页等），
    //    导航该页到服务商端 /sup/ 控制台——已登录会渲染控制台（URL 留在 /sup/），未登录会被踢回 SSO。
    //    每 8 秒最多一次，避免打断用户正在操作的登录页（SSO/护照页不导航）。
    if (!loggedPage && sess.mode === 'cdp' && Date.now() - lastSquareCheck > 8000) {
      lastSquareCheck = Date.now();
      let landing = null;
      for (const ctx of allContexts()) {
        let pages = [];
        try { pages = (ctx && typeof ctx.pages === 'function') ? ctx.pages() : []; } catch (_) { continue; }
        landing = pages.find(p => {
          try {
            if (!p || p.isClosed()) return false;
            const u = p.url() || '';
            // 星图域名下、尚未进入服务商端 /sup/ 控制台、且不在 SSO/护照/登录链路 → 落地页，导航到 /sup/ 复核
            return /xingtu\.cn/i.test(u) && !isLoggedInUrl(u) &&
              !/sso\.oceanengine\.com|passport|redirect_uri=|\/(sso|login)(?:\/|$)/i.test(u);
          } catch (_) { return false; }
        }) || null;
        if (landing) break;
      }
      if (landing) {
        try {
          await landing.goto(SUP_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
          await landing.waitForTimeout(2500);
          loggedPage = await findLoggedInPage();
        } catch (_) { /* 下轮再试 */ }
      }
    }
    if (loggedPage) {
      await finishLogin(loggedPage);
      _loggedIn = true;
      console.log('✅ 登录成功，Cookie 已备份到 .xc-cookies.json。');
      // v3.4：CDP 模式登录成功后自动关闭浏览器窗口（无头模式无窗口，保留供打标）
      await closeLoginBrowser();
      return {
        ok: true, loggedIn: true, mode: sess.mode, browserName: sess.browserName,
        windowClosed: sess.mode === 'cdp',
        message: sess.mode === 'cdp'
          ? '登录成功，浏览器窗口已自动关闭（打标时会在后台自动重开，无需再操作）。'
          : '登录成功',
      };
    }
    await sleep(LOGIN_POLL_MS);
  }
  return {
    ok: false, loggedIn: false, mode: sess.mode,
    message: sess.mode === 'headless'
      ? '登录超时（5 分钟未扫码成功）：请用手机抖音 App 扫描网页上的二维码后重试。'
      : '登录超时（5 分钟未检测到登录成功）：请在打开的浏览器窗口里完成巨量星图登录（扫码或账号密码）后重试。',
  };
}

// ---- 星图接口抓取（在页面上下文内 fetch，同源自动带 Cookie）----------------
async function xgFetch(page, apiPath, body) {
  const url = XINGTU_ORIGIN + apiPath;
  return await page.evaluate(async ({ url, body }) => {
    const resp = await fetch(url, {
      method: 'POST',
      credentials: 'include',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    const text = await resp.text();
    let json = null;
    try { json = JSON.parse(text); } catch (_) { json = { _raw: text.slice(0, 2000) }; }
    return { status: resp.status, json };
  }, { url, body });
}

// 深度遍历：在响应里收集所有“看起来像达人”的对象
function deepCollectAuthors(node, out, depth) {
  if (!node || depth > 8 || out.length > 200) return;
  if (Array.isArray(node)) { node.forEach(n => deepCollectAuthors(n, out, depth + 1)); return; }
  if (typeof node === 'object') {
    const keys = Object.keys(node);
    const looksAuthor =
      (/nick_?name|author_?name|name/i.test(keys.join(','))) &&
      (/(author|uid|user|star|fans|follower)/i.test(keys.join(',')));
    if (looksAuthor) out.push(node);
    for (const k of keys) {
      const v = node[k];
      if (v && typeof v === 'object') deepCollectAuthors(v, out, depth + 1);
    }
  }
}

// 从一个达人对象里按候选字段名优先级取值
function pick(obj, patterns) {
  for (const pat of patterns) {
    for (const k of Object.keys(obj || {})) {
      if (pat.test(k)) {
        const v = obj[k];
        if (v != null && v !== '' && typeof v !== 'object') return v;
      }
    }
  }
  return '';
}
function normText(v) { return String(v == null ? '' : v).trim(); }

// 粉丝量级档位（与存量/星川达人库「万粉标签」口径一致）
function fansTierOf(fans) {
  fans = Number(fans) || 0;
  if (fans >= 10000000) return '千万粉';
  if (fans >= 1000000) return '百万粉';
  if (fans >= 500000) return '50-100万粉';
  if (fans >= 100000) return '10-50万粉';
  if (fans >= 10000) return '1-10万粉';
  if (fans > 0) return '万粉以下';
  return '';
}

// ---- 存量达人库受控词表（与飞书多维表格 tblUXi2raUVtv3et 多选字段选项一致）-----
// 人设标签/内容形式/行业标签在存量库里都是固定选项的多选字段，打标输出直接映射到
// 这些受控词，方便结果 CSV 直接回填/导入存量库；匹配不到就留空，绝不臆造。
const STOCK_PERSONA_RULES = [
  ['美妆护肤达人', /美妆|护肤|彩妆|口红|面膜|香水|粉底/],
  ['母婴亲子达人', /母婴|亲子|宝宝|育儿|孕|婴|奶爸|奶妈/],
  ['服饰穿搭达人', /穿搭|服饰|服装|男装|女装|鞋靴|鞋|搭配|时尚/],
  ['美食测评达人', /美食|零食|吃播|吃货|探店|饮品|餐饮|料理|好吃|(?<!宠物)食品(?!猫|狗)/],
  ['家居清洁达人', /家居|家清|清洁|收纳|家务|居家好物/],
  ['剧情娱乐达人', /剧情|搞笑|娱乐|段子|情景剧|综艺|演绎/],
  ['专业测评达人', /测评|评测|开箱|体验|实验|实测/],
  ['女性种草达人', /种草|爱用物|好物分享|好物推荐/],
  ['生活好物推荐官', /生活好物|生活技巧|生活小窍|实用好物/],
  // v3.7.0 补齐存量库人设字段剩余三个可选值
  ['高潜带货达人', /带货|爆款|出单|转化|小黄车|橱窗|热销|卖爆/],
  ['素人真实分享', /素人|真实分享|无广|自用分享|亲测|随手拍|普通人/],
  ['垂类达人', /垂类|垂直领域|深耕|专注(?!力)/],
];
const STOCK_FORM_RULES = [
  ['好物评测', /评测|测评/],
  ['对比测评', /对比/],
  ['开箱展示', /开箱/],
  ['教程教学', /教程|教学|攻略|怎么|技巧|方法/],
  ['清单合集', /清单|合集|盘点/],
  ['剧情植入', /剧情|情景剧|段子|演绎/],
  ['口播讲解', /口播|讲解|解说/],
  ['直播切片', /直播/],
  ['产品种草', /种草/],
  ['生活记录', /生活|vlog|日常|记录/],
];
// v3.7.0 新增：画面风格标签（存量库多选字段，10 个固定选项；依据视频标题/内容主题推断）
const STOCK_STYLE_RULES = [
  ['真人出镜', /真人出镜|本人出镜|出镜|博主本人|露脸/],
  ['口播字幕', /口播|字幕|讲解|解说|配音|旁白/],
  ['测评实验感', /测评|评测|实测|实验|测试|横评|数据|检测/],
  ['剧情化镜头', /剧情|情景剧|段子|反转|演绎|分镜/],
  ['实拍近景', /实拍|近景|第一视角|第一人称|沉浸式|近距离|特写拍摄/],
  ['产品特写', /特写|细节|质地|上手|质感|微距/],
  ['前后对比', /前后对比|对比图|before|after|变化|逆袭|七天|一个月/],
  ['生活化自然光', /自然光|生活化|日常感|居家拍|vlog感|生活记录/],
  ['高饱和展示', /高饱和|鲜艳|高颜值|ins风|高级感|大片|色彩/],
  ['简洁干净', /极简|简洁|干净|白底|纯色背景|简约/],
];
// v3.7.0 新增：拍摄场景标签（存量库多选字段，10 个固定选项）
const STOCK_SCENE_RULES = [
  ['美妆护肤', /化妆台|化妆间|妆容|护肤|美妆|上脸|试色|梳妆台/],
  ['服饰穿搭', /试衣间|穿搭|试穿|ootd|换装|衣帽间/],
  ['厨房餐桌', /厨房|餐桌|做饭|料理|烹饪|烘焙|美食制作/],
  ['居家场景', /客厅|卧室|房间|家里|居家|书房|沙发/],
  ['浴室洗护', /浴室|洗澡|沐浴|洗护|洗发水|身体乳|卫生间/],
  ['户外出行', /户外|街拍|出行|旅行|外景|公园|路边/],
  ['办公室通勤', /办公室|通勤|职场|工位|上班/],
  ['母婴亲子', /母婴|亲子|宝宝|育儿|带娃|儿童房/],
  ['宠物互动', /宠物|撸猫|遛狗|萌宠|猫舍|狗/],
  ['货架/橱窗展示', /橱窗|货架|柜台|门店|店铺|陈列|直播间|专柜/],
];
const STOCK_INDUSTRY_RULES = [
  ['美妆', /美妆|彩妆|口红|面膜|香水|粉底/],
  ['个人护理', /护肤|个护|洗护|洗发|沐浴|牙膏|护发|身体乳/],
  ['母婴用品', /母婴|婴|孕|宝宝|育儿/],
  ['服装', /服装|穿搭|男装|女装|鞋/],
  ['食品饮料', /(?<!宠物)食品(?!猫|狗)|零食|美食|吃/],
  ['水饮冲调', /饮品|饮料|冲调|咖啡|奶茶|水饮/],
  ['生鲜食品', /生鲜/],
  ['家清纸品', /家清|清洁|纸巾|洗衣|纸品/],
  ['家居家装', /家居|家装|家具|收纳/],
  ['3C数码家电', /3c|数码|家电|电器|手机|电脑/],
  ['运动户外', /运动|户外|健身|瑜伽/],
  ['宠物生活', /宠物|猫|狗/],
  ['健康滋补', /滋补|保健|养生|营养/],
  ['珠宝配饰', /珠宝|配饰|首饰/],
  ['图书教育', /图书|教育|课程/],
  ['本地生活', /本地生活|探店/],
];
// 按关键词命中受控词表，保持词表顺序，最多取 cap 个
function matchControlled(text, rules, cap) {
  const out = [];
  const s = String(text || '').toLowerCase();
  if (!s) return out;
  for (const [name, re] of rules) {
    if (re.test(s) && !out.includes(name)) out.push(name);
    if (out.length >= cap) break;
  }
  return out;
}

// 把星图原始达人对象映射为打标所需指标（字段名按存量达人库口径对齐）
function mapAuthor(raw) {
  const blob = JSON.stringify(raw);
  // ID
  const id = normText(pick(raw, [/^author_id$/i, /star_author_id/i, /^uid$/i, /author_uid/i, /sec_?author/i, /^id$/i])) ||
    (blob.match(/(\d{15,20})/) || [])[1] || '';
  // 昵称
  const name = normText(pick(raw, [/nick_?name/i, /author_?name/i, /^name$/i]));
  // 粉丝
  const fans = Number(pick(raw, [/fans_?count/i, /follower/i, /^fans$/i, /fans_num/i])) || 0;
  const fansTier = fansTierOf(fans);
  // 星川等级 S0-S5：优先取形如 S5 的字段值
  let sLevel = '';
  const sMatch = blob.match(/"[^"]*(?:xc|xinchuan|星川|star_?level|s_?level)[^"]*"\s*:\s*"?(S[0-5])"?/i);
  if (sMatch) sLevel = sMatch[1].toUpperCase();
  if (!sLevel) { const anyS = blob.match(/\b(S[0-5])\b/); if (anyS) sLevel = anyS[1]; }
  // 电商等级 L0-L5
  let lLevel = '';
  const lMatch = blob.match(/"[^"]*(?:ecom|电商|electronic|l_?level)[^"]*"\s*:\s*"?(L[0-5])"?/i);
  if (lMatch) lLevel = lMatch[1].toUpperCase();
  // 交付项目数 / 星图项目数
  const deliveries = Number(pick(raw, [/project_?count/i, /item_?count/i, /cooperate_?count/i, /合作_?数|项目_?数|接单/i, /deliver/i, /trade_?count/i, /order_?count/i])) || 0;
  // 星图消耗（元）
  const consumption = Number(pick(raw, [/consume|consumption|spend|星图_?消耗|消耗|gmv|amount|cost/i])) || 0;
  // —— 存量达人库标签字段（尽力提取；星图接口返回则带出，无则留空，不臆造）——
  // 原始人设/内容形式/类目文本先从接口字段尽力提取，再映射到存量库受控词表
  // （达人人设标签/内容形式标签/行业标签均为固定选项多选字段），匹配不到则留空。
  const personaRaw = normText(pick(raw, [/persona|人设|标签|tag|label/i])).split(/[、,，/|]/).map(s => s.trim()).filter(Boolean);
  const formsRaw = normText(pick(raw, [/content_?type|内容_?形式|形式|video_?type|material/i])).split(/[、,，/|]/).map(s => s.trim()).filter(Boolean);
  const category = normText(pick(raw, [/category|cate|类目|行业|industry|vertical/i])).slice(0, 60);
  const evidence = [].concat(personaRaw, formsRaw, [category]).join(' ');
  const persona = matchControlled(evidence, STOCK_PERSONA_RULES, 3);
  const forms = matchControlled(evidence, STOCK_FORM_RULES, 4);
  const industry = matchControlled(evidence, STOCK_INDUSTRY_RULES, 2);
  // v3.7.0：画面风格 / 拍摄场景标签（同样映射存量库受控词表，主页资料信号弱时可留空，后续视频反哺）
  const style = matchControlled(evidence, STOCK_STYLE_RULES, 3);
  const scene = matchControlled(evidence, STOCK_SCENE_RULES, 3);
  return { id, name, fans, fansTier, sLevel, lLevel, deliveries, consumption, persona, forms, style, scene, industry, category, raw };
}

// 按昵称搜索达人
async function searchByName(page, kw) {
  const payloads = [
    { keyword: kw, query: kw, page: 1, size: 10, offset: 0, limit: 10 },
    { search_keyword: kw, keyword: kw, page_no: 1, page_size: 10 },
  ];
  for (const body of payloads) {
    try {
      const r = await xgFetch(page, '/gw/api/gsearch/search_for_author_square', body);
      if (DEBUG_DUMP) fs.writeFileSync(path.join(__dirname, `debug_search_${Date.now()}.json`), JSON.stringify(r.json, null, 2));
      const authors = [];
      deepCollectAuthors(r.json, authors, 0);
      const mapped = authors.map(mapAuthor).filter(a => a.name);
      if (mapped.length) {
        // 昵称完全匹配优先
        mapped.sort((a, b) => (b.name === kw ? 1 : 0) - (a.name === kw ? 1 : 0));
        return mapped[0];
      }
    } catch (_) { /* 试下一种 payload */ }
  }
  return null;
}

// v3.6：打标主链路不再走批量资料接口——直接用达人 ID 进创作者主页抓取
// （fetchCreatorHome，XHR 拦截 + DOM 降级），昵称搜索仅在主页打不开时兜底。

// 对单个达人打标（输出字段与存量达人库口径对齐）
function scoreAuthor(auth, opts = {}) {
  const dataSource = opts.dataSource === 'stock' || opts.dataSource === 'increment' ? opts.dataSource : 'new';
  const videos = Array.isArray(opts.videoAnalysis) ? opts.videoAnalysis : [];
  const sScore = S_LEVEL_SCORE[auth.sLevel] != null ? S_LEVEL_SCORE[auth.sLevel] : 0;
  const dScore = scoreDeliveries(auth.deliveries);
  const cScore = scoreConsumption(auth.consumption);
  const lScore = scoreEcomLevel(auth.lLevel);
  const baseScore = Math.round((sScore + dScore + cScore + lScore) * 10) / 10;
  // v3.5 双库口径：存量库保持 80 分制；增量库/新达人叠加视频内容分析加分（最多 +20，总分封顶 100）
  let videoBonus = 0;
  let bonusWhy = [];
  if (dataSource !== 'stock' && videos.length) {
    const vb = scoreVideoBonus(videos);
    videoBonus = vb.bonus;
    bonusWhy = vb.why;
  }
  const score = Math.min(100, Math.round((baseScore + videoBonus) * 10) / 10);
  const { tier, medal } = tierOf(score);
  // 标签：星川等级 + 粉丝量级 + 消耗档 + 电商等级 + 交付经验 + 人设/内容形式（存量库标签风格）
  const tags = [];
  if (auth.sLevel) tags.push(auth.sLevel + ' 星川');
  if (auth.fansTier) tags.push(auth.fansTier);
  if (auth.consumption > 100000) tags.push('高星图消耗');
  else if (auth.consumption > 10000) tags.push('中星图消耗');
  if (auth.lLevel && /L[3-5]/.test(auth.lLevel)) tags.push(auth.lLevel + ' 电商');
  if (auth.deliveries > 50) tags.push('交付经验丰富');
  (auth.persona || []).slice(0, 2).forEach(t => { if (t && t.length <= 12 && !tags.includes(t)) tags.push(t); });
  (auth.forms || []).slice(0, 1).forEach(t => { if (t && t.length <= 8 && !tags.includes(t)) tags.push(t); });
  videos.slice(0, 2).forEach(v => { if (v && v.contentType && !tags.includes(v.contentType)) tags.push(v.contentType); });
  if (videoBonus > 0) tags.push(`视频+${videoBonus}`);
  if (!tags.length) tags.push('待培育');
  const dsText = dataSource === 'stock' ? '存量库' : dataSource === 'increment' ? '增量库' : '未在库(新达人)';
  let reason =
    `星川${auth.sLevel || '未分级'}(${sScore}分) · 星图消耗${fmtWan(auth.consumption)}(${cScore}分) · ` +
    `交付${auth.deliveries}个项目(${dScore}分) · 电商${auth.lLevel || '—'}(${lScore}分)` +
    (auth.fansTier ? ` · ${auth.fansTier}` : '') +
    (auth.industry && auth.industry.length ? ` · 行业${auth.industry.join('/')}` : '') +
    (auth.category ? ` · 类目${auth.category}` : '') +
    ` · 基础分${baseScore}`;
  if (videos.length) {
    reason += `\n视频分析: ` + videos.slice(0, 3).map(v => `[${(v.title || '').slice(0, 15)}]`).join('');
    if (videoBonus > 0) reason += `\n视频加分: +${videoBonus}（${bonusWhy.join('；')}）`;
  }
  reason += `\n数据来源: ${dsText}`;
  return { score, tier, medal, tags, reason, videoBonus, dataSource, baseScore };
}

// ---- v3.5：达人主页视频分析（进达人详情页抓前三个视频，辅助内容方向判断与加分）-------
// v3.6.1：创作者主页改走 /sup/author/detail/{id}（旧 /ad/creator/detail/{id} 在普通星图账号下
// 会被重定向到广场/SSO，导致「未检索到」）；XHR 拦截关键字与 DOM 降级解析均与路径无关，自动适用。
const CREATOR_DETAIL_URL = (id) => `https://www.xingtu.cn/sup/author/detail/${id}`;
// v3.7.0：星图主页链接（对外展示 / 复制用）。服务商登录后进服务商达人广场，点达人进入的
// 落地页为市场详情页 /ad/creator/market/detail/{id}，「创作能力」tab 下有该达人作品视频列表。
// 注意：该路径是 SPA 站内路由——浏览器整页冷 goto 会被星图重定向到 /ad/creator/index（需广告主
// 资质），只有在服务商广场页上下文里站内跳转才能正常渲染。故抓取主链路仍用 CREATOR_DETAIL_URL
// （/sup/author/detail，服务商端会话直开）；市场详情页 + 创作能力视频的站内导航抓取待真机校准。
const MARKET_DETAIL_URL = (id) => `https://www.xingtu.cn/ad/creator/market/detail/${id}`;
// v3.7.2（实测定稿）：服务商端达人详情页 /provider/pages/author/douyin/{id}。
// 在服务商 /sup/ 会话下【可整页冷开、不被重定向】（广告主端 /ad/creator/market/detail 冷开会被踢到
// /ad/creator/index，必须站内 SPA 跳转，批量极不稳定）。服务商达人广场点达人昵称即在新 tab 打开此页。
// 页内顶部 el-tabs 有「达人概览/商业能力/创作能力/…」；点「创作能力」会拉
// /gw/api/author/get_author_show_items_v2，data.latest_item_info（个人最新 15 条）+
// data.latest_star_item_info（星图商单 15 条）即创作能力视频列表。此 URL 同时作为星图主页链接。
const PROVIDER_AUTHOR_URL = (id) => `https://www.xingtu.cn/provider/pages/author/douyin/${id}`;

// "12.5万"/"1.2亿"/"3w"/纯数字 → 数字
function parseMediaCount(t) {
  if (typeof t === 'number') return isFinite(t) ? t : 0;
  const s = String(t == null ? '' : t).replace(/[,，\s]/g, '');
  const m = s.match(/(\d+(?:\.\d+)?)\s*(亿|万|[wW])?/);
  if (!m) return 0;
  let n = parseFloat(m[1]);
  if (m[2] === '亿') n *= 1e8;
  else if (m[2] === '万' || m[2] === 'w' || m[2] === 'W') n *= 1e4;
  return Math.round(n);
}

// 视频内容形式推断（仅用视频标题，匹配不到留空、不臆造）
function inferVideoContentType(title) {
  const t = title || '';
  if (/测评|评测|开箱|实测|横评|对比|值不值|避坑|红黑榜|真的好用/.test(t)) return '测评';
  if (/种草|好物|推荐|清单|必买|爱用|安利|闭眼入|盘点|合集/.test(t)) return '种草';
  if (/口播|知识|科普|教程|攻略|干货|解说|怎么选|方法|小课堂/.test(t)) return '口播';
  if (/剧情|段子|搞笑|反转|情景剧/.test(t)) return '剧情';
  if (/vlog|Vlog|VLOG|日常|记录生活/.test(t)) return 'vlog';
  return '';
}

// 带货/产品植入信号词（标题启发式；API 返回商品字段时直接判 true）
const VIDEO_SELL_KW = /同款|链接|下单|到手|橱窗|好物|种草|推荐|购买|小黄车|价格|多少钱|划算|旗舰店|正品|囤货|福利|专场|直播|带货|优惠|券|平替|新品|品牌/;
function videoIsSelling(v) {
  return !!(v && (v.hasProduct || (v.title && VIDEO_SELL_KW.test(v.title))));
}

// v3.7.0：由视频 ID / 分享字段拼出可点击视频链接（星图作品即抖音视频，优先抖音视频页；
// 有显式 share_url / href 时优先归一化使用）。
function buildVideoUrl(id, explicit) {
  const s = String(explicit || '').trim();
  if (/^https?:\/\//i.test(s)) return s;
  if (/^\/\//.test(s)) return 'https:' + s;
  const vid = String(id || '').trim();
  if (/^\d{6,}$/.test(vid)) return `https://www.douyin.com/video/${vid}`;
  return '';
}

// 深度遍历星图接口 JSON，收集“像视频/作品”的对象（宽松匹配，字段名待真机校准）
function deepCollectVideos(node, out, depth) {
  if (!node || depth > 10 || out.length >= 15) return;
  if (Array.isArray(node)) { for (const n of node) deepCollectVideos(n, out, depth + 1); return; }
  if (typeof node === 'object') {
    const stat = node.statistics || node.stats || node.stat || (node.video && node.video.statistics) || {};
    const vid = node.aweme_id ?? node.item_id ?? node.video_id ?? node.awemeId ?? node.itemId ?? stat.aweme_id;
    const title = normText(node.desc ?? node.title ?? node.item_title ?? node.video_title ?? node.name);
    const hasPlayLike = node.play_count != null || node.playCount != null || node.play_cnt != null ||
      node.digg_count != null || node.diggCount != null || stat.play_count != null || stat.digg_count != null ||
      (node.video && typeof node.video === 'object');
    const looksVideo = (vid && /^\d{6,}$/.test(String(vid)) && title) || (title && hasPlayLike);
    const isUser = node.follower_count != null || node.fans_count != null || node.followerCount != null; // 排除达人对象
    if (looksVideo && !isUser) {
      const plays = parseMediaCount(node.play_count ?? node.playCount ?? node.play_cnt ?? stat.play_count ?? stat.playCount ?? 0);
      const likes = parseMediaCount(node.digg_count ?? node.diggCount ?? stat.digg_count ?? stat.diggCount ?? node.like_count ?? 0);
      const hasProduct = !!(node.product_info || node.products || node.goods || node.goods_list ||
        node.with_goods || node.shop_goods || node.ecommerce_info || node.product_related ||
        node.anchor_info || node.promotions || node.channels ||
        (node.video && (node.video.product || node.video.goods || node.video.product_info)));
      // v3.7.0：带出视频链接（share_url / share_link / 播放地址；兜底按抖音 ID 拼）
      const shareUrl = node.share_url ?? node.share_link ?? node.share_url_link ?? node.play_url ??
        node.aweme_url ?? node.video_url ?? node.url ??
        (node.video && (node.video.share_url || node.video.play_url)) ?? '';
      out.push({ id: String(vid || ''), url: buildVideoUrl(vid, shareUrl), title: title.slice(0, 80), plays, likes, hasProduct });
    }
    for (const k of Object.keys(node)) {
      const v = node[k];
      if (v && typeof v === 'object') deepCollectVideos(v, out, depth + 1);
    }
  }
}

// 浏览器侧 DOM 降级：抓视频/作品卡片文本（纯浏览器 JS，不引用 Node 侧变量）
function domVideoCardsEval() {
  const out = [];
  const seenText = new Set();
  const push = (el) => {
    const text = (el.innerText || '').replace(/\s+/g, ' ').trim();
    if (text.length < 8 || text.length > 400) return;
    const key = text.slice(0, 24);
    if (seenText.has(key)) return;
    seenText.add(key);
    const a = el.matches && el.matches('a') ? el : el.querySelector('a[href*="video"], a[href*="aweme"]');
    out.push({ text, href: a ? a.href : '' });
  };
  const sels = [
    'a[href*="/video/"]', 'a[href*="aweme"]',
    '[class*="video" i]', '[class*="work" i]', '[class*="aweme" i]',
    '[class*="content-card" i]', '[class*="creator-card" i]',
  ];
  for (const s of sels) { try { document.querySelectorAll(s).forEach(push); } catch (_) {} }
  return out.slice(0, 12);
}

/**
 * v3.6：浏览器侧——读取达人主页状态与 DOM 资料（纯浏览器 JS，不引用 Node 侧变量）。
 * 返回：{ url, onDetail, notFound, headText, chips, name }
 *  - onDetail：最终 URL 仍停在创作者详情页（v3.6.1：/sup/author/detail/{数字}；旧 /ad/creator/detail/{数字} 兼容），被重定向到广场/SSO 则 false
 *  - notFound：正文含「达人不存在/未入驻/404」等失效文案
 *  - headText：正文前 3000 字（Node 侧正则提取粉丝/等级/项目/消耗）
 *  - chips：页面上可见的短标签（内容主题/类目），过滤数字与操作按钮
 *  - name：尽力从昵称语义节点/h1/h2 取达人昵称
 */
function domHomeStateEval() {
  const out = { url: '', onDetail: false, notFound: false, headText: '', chips: [], name: '' };
  try {
    out.url = location.href;
    // v3.7.2：onDetail 首选服务商端达人详情页 /provider/pages/author/douyin/{id}（可冷开，主链路）；
    // 兼容服务商广场站内跳转的市场详情 /ad/creator/market/detail/{id}、/sup/author/detail/{id}、旧 /ad/creator/detail/{id}。
    out.onDetail = /\/provider\/pages\/author\/douyin\/\d+/.test(out.url) ||
      /\/ad\/creator\/market\/detail\/\d+/.test(out.url) ||
      /\/(?:sup\/author|ad\/creator)\/detail\/\d+/.test(out.url);
    const text = (document.body && document.body.innerText || '').replace(/\s+/g, ' ').trim();
    out.headText = text.slice(0, 3000);
    out.notFound = /达人不存在|创作者不存在|达人已注销|账号已注销|页面不存在|找不到相关|未入驻|暂无权限|没有权限|内容不存在|\b404\b/.test(text.slice(0, 1000));
    // 昵称：优先语义化 class，再退 h1/h2；过滤导航/操作文案与纯数字
    const nameSels = [
      '[class*="nickname" i]', '[class*="nick-name" i]', '[class*="userName" i]',
      '[class*="user-name" i]', '[class*="authorName" i]', '[class*="author-name" i]',
      'h1', 'h2',
    ];
    for (const s of nameSels) {
      let el = null;
      try { el = document.querySelector(s); } catch (_) { continue; }
      if (!el) continue;
      const t = (el.innerText || '').replace(/\s+/g, ' ').trim();
      if (t && t.length >= 2 && t.length <= 30 && !/^\d+$/.test(t) &&
          !/登录|注册|搜索|首页|广场|收藏|分享|评论|私信|关注|粉丝|获赞/.test(t)) {
        out.name = t;
        break;
      }
    }
    // 内容主题/类目标签 chips
    const seen = new Set();
    const chipSels = ['[class*="tag" i]', '[class*="label" i]', '[class*="chip" i]', '[class*="topic" i]'];
    for (const s of chipSels) {
      let els = [];
      try { els = Array.from(document.querySelectorAll(s)); } catch (_) { continue; }
      for (const el of els) {
        try {
          if (el.offsetParent === null) continue; // 不可见
          const t = (el.innerText || '').replace(/\s+/g, ' ').trim();
          if (!t || t.length < 2 || t.length > 12) continue;
          if (/粉丝|关注|获赞|点赞|播放|登录|注册|收藏|分享|评论|私信|下单|预约|合作|咨询|客服|首页|广场|\d|￥|¥/.test(t)) continue;
          if (seen.has(t)) continue;
          seen.add(t);
          out.chips.push(t);
          if (out.chips.length >= 40) break;
        } catch (_) { /* skip */ }
      }
      if (out.chips.length >= 40) break;
    }
  } catch (_) { /* 页面异常时返回空状态，调用方按打不开处理 */ }
  return out;
}

// v3.6：Node 侧——从主页 DOM 状态正则提取达人资料（字段名待真机校准，宽松匹配）
function parseDomProfile(state, authorId) {
  if (!state || !state.onDetail || state.notFound) return null;
  const text = String(state.headText || '');
  if (text.length < 20) return null;
  // 粉丝数：125.6万粉丝 / 1.2亿粉丝 / 8632粉丝
  let fans = 0;
  let m = text.match(/([\d.]+)\s*([亿万wW]?)\s*粉丝(?:数|量)?/);
  if (m) fans = parseMediaCount(m[1] + (m[2] || ''));
  // 星川等级 S0-S5：优先「星川…S4」语境，兜底独立 S 等级徽标
  let sLevel = '';
  m = text.match(/星川[\s\S]{0,12}?(S[0-5])(?![0-9])/i);
  if (m) sLevel = m[1].toUpperCase();
  if (!sLevel) { m = text.match(/(?:^|[^A-Za-z0-9])(S[0-5])(?![0-9])/); if (m) sLevel = m[1].toUpperCase(); }
  // 电商等级 L0-L5
  let lLevel = '';
  m = text.match(/电商[\s\S]{0,12}?(L[0-5])(?![0-9])/i);
  if (m) lLevel = m[1].toUpperCase();
  if (!lLevel) { m = text.match(/(?:^|[^A-Za-z0-9])(L[0-5])(?![0-9])/); if (m) lLevel = m[1].toUpperCase(); }
  // 交付/合作项目数
  let deliveries = 0;
  m = text.match(/(?:合作项目|交付项目|项目数|累计合作|累计交付|接单|完成项目|合作过)[^0-9]{0,8}(\d+(?:\.\d+)?)/);
  if (m) deliveries = parseInt(m[1], 10) || 0;
  if (!deliveries) {
    m = text.match(/(\d+(?:\.\d+)?)\s*个?\s*(?:合作项目|交付项目|个项目|项目数|条合作)/);
    if (m) deliveries = parseInt(m[1], 10) || 0;
  }
  // 星图消耗（元）：页面常显示「125.3万」「1.2亿」
  let consumption = 0;
  m = text.match(/(?:星图|累计)?消耗[^0-9]{0,8}(\d+(?:\.\d+)?)\s*([亿万wW]?)\s*(?:元)?/);
  if (m) consumption = parseMediaCount(m[1] + (m[2] || ''));
  // 内容主题标签：chips 优先（短标签信号最强），正文补充；映射到存量库受控词
  const chips = Array.isArray(state.chips) ? state.chips : [];
  const evidence = chips.join(' ') + ' ' + text.slice(0, 1200);
  const persona = matchControlled(evidence, STOCK_PERSONA_RULES, 2);
  const forms = matchControlled(evidence, STOCK_FORM_RULES, 2);
  const industry = matchControlled(evidence, STOCK_INDUSTRY_RULES, 1);
  const style = matchControlled(evidence, STOCK_STYLE_RULES, 2);
  const scene = matchControlled(evidence, STOCK_SCENE_RULES, 2);
  const name = normText(state.name).slice(0, 40);
  const hasSignal = fans > 0 || !!sLevel || deliveries > 0 || consumption > 0 || chips.length >= 3;
  if (!hasSignal) return null;
  return {
    id: String(authorId || ''),
    name,
    fans,
    fansTier: fansTierOf(fans),
    sLevel, lLevel, deliveries, consumption,
    persona, forms, style, scene, industry,
    category: chips[0] ? String(chips[0]).slice(0, 20) : '',
    raw: { _dom: true, chips: chips.slice(0, 20), headText: text.slice(0, 500) },
  };
}

// v3.6：主页 XHR 可能挖到多个达人对象（含登录账号自身/推荐位），按 ID 精确匹配优先、
// 资料完整度兜底，避免把广告主账号或推荐达人误当目标。
function authorFullness(a) {
  let n = 0;
  if (a.sLevel) n += 4;
  if (a.deliveries > 0) n += 3;
  if (a.consumption > 0) n += 3;
  if (a.fans > 0) n += 2;
  if (a.name) n += 1;
  if ((a.persona || []).length || (a.forms || []).length) n += 1;
  return n;
}
function pickHomeAuthor(rawAuthors, authorId) {
  const mapped = (Array.isArray(rawAuthors) ? rawAuthors : [])
    .map(mapAuthor).filter(a => a && (a.id || a.name));
  if (!mapped.length) return null;
  const idStr = String(authorId || '');
  if (idStr) {
    const exact = mapped.find(a => String(a.id) === idStr);
    if (exact) return exact;
    const loose = mapped.find(a => a.id &&
      (String(a.id).includes(idStr) || idStr.includes(String(a.id))));
    if (loose) return loose;
  }
  const ranked = mapped.slice().sort((x, y) => authorFullness(y) - authorFullness(x));
  return authorFullness(ranked[0]) >= 3 ? ranked[0] : null;
}

// v3.6：合并两份达人资料——base 优先（搜索接口字段通常更全），extra 只补空字段；
// 标签类列表取并集。用于「昵称搜索命中 + 主页补充资料/视频」场景。
function mergeAuth(base, extra) {
  if (!extra) return base;
  const out = Object.assign({}, base);
  const union = (a, b) => Array.from(new Set([...(a || []), ...(b || [])]));
  if (!out.name && extra.name) out.name = extra.name;
  if (!out.sLevel && extra.sLevel) out.sLevel = extra.sLevel;
  if (!out.lLevel && extra.lLevel) out.lLevel = extra.lLevel;
  if ((!out.fans || out.fans === 0) && extra.fans) out.fans = extra.fans;
  if ((!out.deliveries || out.deliveries === 0) && extra.deliveries) out.deliveries = extra.deliveries;
  if ((!out.consumption || out.consumption === 0) && extra.consumption) out.consumption = extra.consumption;
  if (!out.category && extra.category) out.category = extra.category;
  out.fansTier = fansTierOf(out.fans || 0);
  out.persona = union(base.persona, extra.persona);
  out.forms = union(base.forms, extra.forms);
  out.style = union(base.style, extra.style);   // v3.7.0 画面风格
  out.scene = union(base.scene, extra.scene);   // v3.7.0 拍摄场景
  out.industry = union(base.industry, extra.industry);
  out.raw = extra.raw || base.raw;
  return out;
}

// ---- v3.7.1：服务商广场站内导航，进「创作能力」抓视频列表 XHR ----------------
// 浏览器侧：判断当前是否弹出滑块/安全验证码（需要服务商手动拖）
function captchaPresentEval() {
  try {
    const kw = /拖动|滑块|按住|完成验证|安全验证|向右滑动|拼图|验证码|请完成下方验证/;
    const vis = (e) => { try { return e && e.offsetParent !== null && e.offsetWidth > 60 && e.offsetHeight > 30; } catch (_) { return false; } };
    const els = Array.from(document.querySelectorAll('div,span,button,section'));
    const hit = els.find(e => vis(e) && kw.test((e.innerText || '') + (e.title || '')));
    const ifr = document.querySelector('iframe[src*="captcha" i],iframe[src*="verify" i],iframe[id*="captcha" i]');
    return !!(hit || (ifr && vis(ifr)));
  } catch (_) { return false; }
}
// 遇到滑块验证码时最多等 timeoutMs（默认 30s）等服务商手动拖过；验证码消失即返回
async function waitIfCaptcha(page, timeoutMs = 30000) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const has = await page.evaluate(captchaPresentEval).catch(() => false);
    if (!has) return false;
    await page.waitForTimeout(1500).catch(() => {});
  }
  return true; // 超时仍在验证码
}

/**
 * v3.7.1 站内导航链路（关键：不能冷开详情页）：
 *   广场页 → 搜索达人 ID → 站内点达人卡片（SPA 跳 /ad/creator/market/detail/{id}）
 *   → 点「创作能力」tab → 点前 3 个视频封面触发详情弹窗接口（站内弹窗，URL 不变）。
 * 全程响应由调用方 vp.on('response') 拦截进 apiHits。任何一步失败都返回 reached:false，
 * 由调用方降级（不阻塞打标）。
 */
async function enterCreatorDetailViaMarket(vp, authorId) {
  const id = String(authorId || '').trim();
  if (!id) return { reached: false };
  // ① 进服务商广场（首选 /provider/pages/market，失败回退 /pro/ad/pages/market）
  let opened = false;
  for (const u of [PROVIDER_MARKET_URL, SQUARE_URL]) {
    try {
      await vp.goto(u, { waitUntil: 'domcontentloaded', timeout: 45000 });
      await vp.waitForTimeout(2500);
      await waitIfCaptcha(vp, 30000);
      const hasSearch = await vp.evaluate(() => {
        const vis = (e) => { try { return e && e.offsetParent !== null && !e.disabled; } catch (_) { return false; } };
        return Array.from(document.querySelectorAll('input')).some(i => vis(i));
      }).catch(() => false);
      if (hasSearch) { opened = true; break; }
    } catch (_) { /* 试下一个广场入口 */ }
  }
  if (!opened) return { reached: false };

  // ② 搜索框输入达人 ID 并回车
  const searched = await vp.evaluate((qid) => {
    const vis = (e) => { try { return e && e.offsetParent !== null && !e.disabled; } catch (_) { return false; } };
    const inputs = Array.from(document.querySelectorAll('input'));
    const inp = inputs.find(i => vis(i) && /搜索|达人|昵称|账号|关键词|星图/.test(i.placeholder || '')) ||
      inputs.find(i => vis(i) && /search/i.test((i.className || '') + (i.getAttribute('aria-label') || ''))) ||
      inputs.find(i => vis(i));
    if (!inp) return 'no-input';
    const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
    setter.call(inp, String(qid));
    inp.dispatchEvent(new Event('input', { bubbles: true }));
    inp.dispatchEvent(new Event('change', { bubbles: true }));
    try { inp.focus(); } catch (_) {}
    inp.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
    inp.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true }));
    return 'ok';
  }, id).catch(() => 'err');
  if (searched !== 'ok') {
    // 回车不行 → 点「搜索」按钮兜底
    await vp.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('button,[role="button"],span,div,a'));
      const b = btns.find(e => e.offsetParent !== null && /^搜索$/.test((e.innerText || '').trim()));
      if (b) try { b.click(); } catch (_) {}
    }).catch(() => {});
  }
  await vp.waitForTimeout(5000);
  await waitIfCaptcha(vp, 30000);

  // ③ 站内点击达人卡片（优先命中 /market/detail/{id} 的链接）
  const clicked = await vp.evaluate((qid) => {
    const anchors = Array.from(document.querySelectorAll('a[href]'));
    const exact = anchors.find(a => new RegExp('/market/detail/' + qid + '(?:\\?|/|$)').test(a.href)) ||
      anchors.find(a => new RegExp('/detail/' + qid + '(?:\\?|/|$)').test(a.href));
    const anyDetail = anchors.find(a => /\/ad\/creator\/(?:market\/)?detail\/\d+/.test(a.href));
    const target = exact || anyDetail;
    if (target) { target.scrollIntoView({ block: 'center' }); try { target.click(); } catch (_) {} return target.href; }
    // 无 <a>：找带该 ID 的可点击卡片兜底
    const cards = Array.from(document.querySelectorAll('[class*="creator" i],[class*="author" i],[class*="card" i],[class*="result" i]'));
    const card = cards.find(c => c.offsetParent !== null && new RegExp(qid).test(c.getAttribute('data-id') || c.outerHTML || ''));
    if (card) { card.scrollIntoView({ block: 'center' }); try { card.click(); } catch (_) {} return 'card:' + qid; }
    return '';
  }, id).catch(() => '');
  if (!clicked) return { reached: false };

  // ④ 等 SPA 路由到市场详情
  try {
    await vp.waitForFunction((qid) => new RegExp('/ad/creator/(?:market/)?detail/' + qid + '(?:\\?|/|$)').test(location.href), id, { timeout: 15000 });
  } catch (_) { /* 路由可能不同，继续尝试点 tab */ }
  await vp.waitForTimeout(3000);
  await waitIfCaptcha(vp, 30000);

  // ⑤ 点「创作能力」tab（作品表现/创作案例区域）
  await vp.evaluate(() => {
    const kw = /创作能力|作品表现|创作案例|^作品$|TA的视频|视频内容/;
    const els = Array.from(document.querySelectorAll('div,span,a,button,li,[role="tab"]'));
    const hit = els.find(e => {
      const t = (e.innerText || '').trim();
      return t && t.length <= 10 && kw.test(t) && e.offsetParent !== null;
    });
    if (hit) { hit.scrollIntoView({ block: 'center' }); try { hit.click(); } catch (_) {} }
  }).catch(() => {});
  await vp.waitForTimeout(3500);
  await waitIfCaptcha(vp, 30000);

  // ⑥ 依次点前 3 个视频封面 → 触发详情弹窗接口（站内弹窗，URL 不变），每个抓完关闭
  for (let k = 0; k < 3; k++) {
    const n = await vp.evaluate((idx) => {
      const vis = (e) => { try { return e && e.offsetParent !== null && e.offsetWidth > 60 && e.offsetHeight > 60; } catch (_) { return false; } };
      const cand = Array.from(document.querySelectorAll(
        'a[href*="video"],a[href*="aweme"],[class*="video-item" i],[class*="work-item" i],[class*="aweme" i],[class*="video-card" i],[class*="cover" i],[class*="video" i]'
      )).filter(vis);
      const el = cand[idx];
      if (!el) return 0;
      el.scrollIntoView({ block: 'center' });
      try { el.click(); } catch (_) {}
      return 1;
    }, k).catch(() => 0);
    if (!n) break;
    await vp.waitForTimeout(2500);
    await waitIfCaptcha(vp, 30000);
    // 关闭弹窗（×/关闭按钮 或 Esc）
    await vp.evaluate(() => {
      const btns = Array.from(document.querySelectorAll('[class*="close" i],[aria-label*="close" i],[class*="mask" i],button,span,div'));
      const c = btns.find(b => b.offsetParent !== null &&
        (/close|关闭|modal|dialog|mask/i.test(b.className || '') || /^(关闭|×|✕|X|x)$/.test((b.innerText || '').trim())));
      if (c) try { c.click(); } catch (_) {}
    }).catch(() => {});
    await vp.keyboard.press('Escape').catch(() => {});
    await vp.waitForTimeout(800).catch(() => {});
  }

  const onDetail = await vp.evaluate((qid) =>
    new RegExp('/ad/creator/(?:market/)?detail/' + qid + '(?:\\?|/|$)').test(location.href), id).catch(() => false);
  return { reached: !!onDetail };
}

/**
 * v3.7.2 主链路（实测定稿，替代不稳定的广场站内导航）：直接冷开服务商端达人详情页
 *   /provider/pages/author/douyin/{id} —— 服务商 /sup/ 会话下整页 goto 不重定向、直接渲染，
 *   无需广场搜索/点卡片/弹新 tab，批量稳定。落地后点顶部 el-tabs「创作能力」tab，
 *   页面即拉 get_author_show_items_v2（响应已由调用方 vp.on('response') 拦截进 apiHits）。
 * 任一步失败返回 reached:false，调用方再降级冷开 /sup/author/detail 至少拿资料，不阻塞打标。
 */
async function enterProviderAuthorPage(vp, authorId) {
  const id = String(authorId || '').trim();
  if (!id) return { reached: false, creative: false };
  try {
    await vp.goto(PROVIDER_AUTHOR_URL(id), { waitUntil: 'domcontentloaded', timeout: 45000 });
  } catch (_) { return { reached: false, creative: false }; }
  await vp.waitForTimeout(7000).catch(() => {});
  await waitIfCaptcha(vp, 30000);
  // 点顶部「创作能力」tab（Element UI .el-tabs__item，精确匹配文案；宽松兜底含“创作能力”的 tab）
  const creative = await vp.evaluate(() => {
    const tabs = Array.from(document.querySelectorAll('.el-tabs__item,[role="tab"],[class*="tab" i]'));
    let t = tabs.find((e) => (e.innerText || '').trim() === '创作能力');
    if (!t) t = tabs.find((e) => /创作能力/.test((e.innerText || '').trim()) && (e.innerText || '').trim().length <= 8);
    if (t) { try { t.scrollIntoView({ block: 'center' }); } catch (_) {} try { t.click(); } catch (_) {} return true; }
    return false;
  }).catch(() => false);
  if (creative) {
    try { await vp.waitForResponse((r) => /get_author_show_items_v2/.test(r.url()), { timeout: 12000 }); }
    catch (_) { /* 接口可能已返回或在飞行中，下面 apiHits 兜底 */ }
  }
  await vp.waitForTimeout(1800).catch(() => {});
  await waitIfCaptcha(vp, 30000);
  const reached = await vp.evaluate((qid) =>
    new RegExp('/provider/pages/author/douyin/' + qid + '(?:\\?|/|$)').test(location.href), id).catch(() => false);
  return { reached: !!reached, creative: !!creative };
}

// v3.7.2：从创作能力 tab 的 get_author_show_items_v2 响应里精确取前三个视频。
// data.latest_item_info = 个人最新视频（默认「全部」列表前 15 条，创作能力代表内容，优先）；
// data.latest_star_item_info = 星图商单视频（个人不足 3 条时补足）。视频链接统一按抖音视频页拼。
function pickShowItemVideos(apiHits) {
  const out = [];
  const seen = new Set();
  const push = (it) => {
    if (!it || typeof it !== 'object') return;
    const id = String(it.item_id || it.video_id || it.aweme_id || '').trim();
    if (!/^\d{15,}$/.test(id) || seen.has(id)) return;
    seen.add(id);
    const title = normText(it.item_title || it.title || it.desc).slice(0, 80);
    const plays = parseMediaCount(it.play != null ? it.play : (it.play_count || 0));
    const likes = parseMediaCount(it.like != null ? it.like : (it.digg_count || 0));
    const hasProduct = !!(it.goods_list || it.goods || it.product_info || it.with_goods || it.has_product) ||
      (!!title && VIDEO_SELL_KW.test(title));
    out.push({ id, url: buildVideoUrl(id, ''), title, plays, likes, hasProduct });
  };
  for (const h of apiHits) {
    if (!/get_author_show_items_v2/.test(h.url || '')) continue;
    const d = (h.json && h.json.data) || {};
    const personal = Array.isArray(d.latest_item_info) ? d.latest_item_info : [];
    const star = Array.isArray(d.latest_star_item_info) ? d.latest_star_item_info : [];
    personal.forEach(push);              // 个人最新视频优先
    if (out.length < 3) star.forEach(push); // 不足补星图商单视频
    break;
  }
  return out.slice(0, 3);
}

/**
 * v3.6 打标主链路（v3.7.2 改为冷开服务商详情页直取创作能力视频）：
 *   先进服务商广场搜索达人 ID → 站内点卡片到 /ad/creator/market/detail/{id} →
 *   点「创作能力」tab，拦截视频列表/详情 XHR，一次拿到：
 *   ① 达人资料：星川等级 / 交付项目数 / 星图消耗 / 电商等级 / 粉丝数 / 内容主题标签
 *      ——优先拦截 XHR/fetch 响应挖达人对象（ID 精确匹配），拦截不到降级 DOM；
 *   ② 前三个视频：链接 / 标题 / 播放 / 点赞 / 内容形式 / 带货信号（XHR 优先、DOM 降级）。
 *   站内导航失败时回退冷开 /sup/author/detail/{id} 至少拿资料；两条路都失败返回
 *   { ok:false }，调用方降级走昵称搜索，绝不阻塞打标。
 * debug=true（XC_TAGGER_DEBUG=1）时落截图 + 接口 JSON + DOM 状态到 .xc-debug/。
 */
async function fetchCreatorHome(page, authorId, { debug = false } = {}) {
  const fail = (reason) => ({ ok: false, auth: null, videos: [], via: '', reason });
  if (!authorId) return fail('no-id');
  const ctx = page.context ? page.context() : page;
  let vp = null;
  const apiHits = [];
  const dbgDir = path.join(process.cwd(), '.xc-debug');
  try {
    vp = await ctx.newPage();
    try { await vp.setViewportSize({ width: 1360, height: 900 }); } catch (_) {}
    vp.on('response', async (resp) => {
      try {
        const u = resp.url();
        if (!/xingtu\.cn|oceanengine\.com/i.test(u)) return;
        const rt = resp.request().resourceType();
        if (rt !== 'xhr' && rt !== 'fetch') return;
        if (!/video|aweme|work|post|creator|author|content|detail|feed|user|info|stat|profile|tag|label|portrait/i.test(u)) return;
        const j = await resp.json();
        apiHits.push({ url: u, json: j });
      } catch (_) { /* 非 JSON */ }
    });
    // v3.7.2：主链路 = 冷开服务商端达人详情页 /provider/pages/author/douyin/{id}（不重定向、批量稳定），
    // 落地后自动/点「创作能力」tab 拉 get_author_show_items_v2；response 监听已挂在 vp 上。
    let nav = await enterProviderAuthorPage(vp, authorId).catch(() => ({ reached: false, creative: false }));
    // 等最后一批视频接口回来
    await vp.waitForResponse((r) => {
      try {
        const rt = r.request().resourceType();
        if (rt !== 'xhr' && rt !== 'fetch') return false;
        return /get_author_show_items_v2|show_items|homepage_videos|video|aweme|work|post|feed|content/i.test(r.url());
      } catch (_) { return false; }
    }, { timeout: 6000 }).catch(() => {});
    await vp.waitForTimeout(1200).catch(() => {});

    // ---- 主页状态：重定向/未入驻 判定 + DOM 资料 ----
    let domState = await vp.evaluate(domHomeStateEval).catch(() => null);
    // 服务商详情页没进成 → 依次降级：① 广场站内导航；② 冷开 /sup/author/detail/{id} 至少拿资料（视频可能为空，不阻塞）
    if (!(domState && domState.onDetail) && nav && !nav.reached) {
      const m2 = await enterCreatorDetailViaMarket(vp, authorId).catch(() => ({ reached: false }));
      domState = await vp.evaluate(domHomeStateEval).catch(() => domState);
      if (!(domState && domState.onDetail) && m2 && !m2.reached) {
        try {
          await vp.goto(CREATOR_DETAIL_URL(authorId), { waitUntil: 'domcontentloaded', timeout: 30000 }).catch(() => {});
          await vp.waitForTimeout(2500).catch(() => {});
          domState = await vp.evaluate(domHomeStateEval).catch(() => domState);
        } catch (_) {}
      }
    }
    const stateKnown = !!domState;
    const onDetail = !!(domState && domState.onDetail);
    const notFound = !!(domState && domState.notFound);
    if (debug) {
      try {
        fs.mkdirSync(dbgDir, { recursive: true });
        const ts = Date.now();
        await vp.screenshot({ path: path.join(dbgDir, `creator_${authorId}_${ts}.png`), fullPage: false }).catch(() => {});
        fs.writeFileSync(path.join(dbgDir, `creator_${authorId}_${ts}.json`),
          JSON.stringify({ domState, nav, apiHits: apiHits.map((h) => ({ url: h.url, json: h.json })) }).slice(0, 3000000), 'utf8');
      } catch (_) {}
    }
    // 站内/冷开两条路都不在详情页，或页面明示不存在 → 主页失败，交调用方兜底
    if (stateKnown && (!onDetail || notFound)) return fail(notFound ? 'notfound' : 'redirect');

    // ---- ① 资料：XHR 挖达人对象（ID 精确匹配优先）----
    const rawAuthors = [];
    for (const h of apiHits) deepCollectAuthors(h.json, rawAuthors, 0);
    let auth = pickHomeAuthor(rawAuthors, authorId);
    let via = auth ? 'xhr' : '';
    // ---- ①b 资料：DOM 降级 / 补全 ----
    const domAuth = parseDomProfile(domState, authorId);
    if (domAuth) {
      auth = auth ? mergeAuth(auth, domAuth) : domAuth;
      if (!via) via = 'dom';
    }
    if (!auth) return fail(stateKnown ? 'no-data' : 'error');

    // ---- ② 视频：首选创作能力 tab 的 get_author_show_items_v2（个人最新视频前3，口径最准）----
    let vids = pickShowItemVideos(apiHits);
    // ②b 兜底：其余 XHR 深度遍历挖视频
    if (vids.length === 0) {
      for (const h of apiHits) deepCollectVideos(h.json, vids, 0);
    }
    const seen = new Set();
    vids = vids.filter((v) => {
      const key = v.id || v.title.slice(0, 20);
      if (!key || seen.has(key)) return false;
      seen.add(key); return true;
    });
    // ---- ②b 视频：DOM 降级 ----
    if (vids.length === 0) {
      const cards = await vp.evaluate(domVideoCardsEval).catch(() => []);
      vids = cards.map((c) => {
        const title = c.text
          .replace(/\d+(?:\.\d+)?\s*(亿|万|[wW])?\s*(次播放|播放|次观看|观看|赞|评论|分享|收藏|点赞)?/g, ' ')
          .replace(/\s+/g, ' ').trim().slice(0, 60);
        const pm = c.text.match(/(\d+(?:\.\d+)?)\s*(亿|万|[wW])?\s*(次播放|播放|次观看)/);
        return {
          id: '', url: buildVideoUrl('', c.href), title: title || '',
          plays: pm ? parseMediaCount(pm[0]) : 0, likes: 0,
          hasProduct: /同款|链接|下单|到手|橱窗|好物|种草|购买|小黄车|价格|划算|直播|带货|优惠/.test(c.text),
          fromDom: true,
        };
      }).filter((v) => v.title && v.title.length >= 4).slice(0, 3);
    }
    const videos = vids.slice(0, 3).map((v) => ({
      title: v.title,
      url: v.url || buildVideoUrl(v.id, ''),
      plays: v.plays || 0,
      likes: v.likes || 0,
      contentType: inferVideoContentType(v.title),
      isSelling: videoIsSelling(v),
    }));

    if (debug) console.log(`  🐞 主页抓取(${authorId})：资料来源=${via}，视频=${videos.length}个`);
    return { ok: true, auth, videos, via, reason: 'ok' };
  } catch (e) {
    if (debug) console.log('  ⚠️ 达人主页访问失败(' + authorId + ')：' + friendlyErr(e));
    return fail('error');
  } finally {
    if (vp) await vp.close().catch(() => {});
  }
}

// 用视频标题/内容形式推断的标签【补充】（不覆盖）达人 forms/persona/industry/style/scene 受控词
function enrichAuthFromVideos(auth, videos) {
  if (!auth || !Array.isArray(videos) || !videos.length) return;
  const text = videos.map((v) => v.title || '').filter(Boolean).join(' ');
  if (!text) return;
  const merge = (oldList, add) => Array.from(new Set([...(oldList || []), ...add]));
  auth.forms = merge(auth.forms, matchControlled(text, STOCK_FORM_RULES, 2));
  auth.persona = merge(auth.persona, matchControlled(text, STOCK_PERSONA_RULES, 2));
  auth.industry = merge(auth.industry, matchControlled(text, STOCK_INDUSTRY_RULES, 2));
  // v3.7.0：画面风格 / 拍摄场景主要由视频内容反哺
  auth.style = merge(auth.style, matchControlled(text, STOCK_STYLE_RULES, 3));
  auth.scene = merge(auth.scene, matchControlled(text, STOCK_SCENE_RULES, 3));
  // 内容形式种子：口播/测评类视频天然对应「真人出镜+口播字幕 / 测评实验感」风格
  const forms = videos.map((v) => v.contentType || '').filter(Boolean);
  const seedStyle = [];
  if (forms.includes('口播')) seedStyle.push('真人出镜', '口播字幕');
  if (forms.includes('测评')) seedStyle.push('测评实验感');
  if (forms.includes('剧情')) seedStyle.push('剧情化镜头');
  if (forms.includes('种草')) seedStyle.push('生活化自然光');
  if (seedStyle.length) auth.style = merge(auth.style, seedStyle);
}

/**
 * v3.5 增量库/新达人口径：前三个视频内容分析加分（最多 +20，总分封顶 100）。
 *  · 带货/产品植入视频：+5/个，最多 +15
 *  · 内容形式为口播/测评/种草：+5
 *  · 任一视频播放量 >10 万：+5
 * 存量库口径（dataSource='stock'）不加分，保持 80 分制。
 */
function scoreVideoBonus(videos) {
  const list = Array.isArray(videos) ? videos : [];
  let bonus = 0;
  const why = [];
  const sellingCnt = list.filter((v) => v && v.isSelling).length;
  const sellBonus = Math.min(sellingCnt * 5, 15);
  if (sellBonus > 0) { bonus += sellBonus; why.push(`${sellingCnt}条带货/产品植入视频+${sellBonus}`); }
  if (list.some((v) => v && /测评|口播|种草/.test(v.contentType || ''))) { bonus += 5; why.push('口播/测评/种草内容形式+5'); }
  if (list.some((v) => v && (Number(v.plays) || 0) >= 100000)) { bonus += 5; why.push('视频播放量超10万+5'); }
  bonus = Math.min(bonus, 20);
  return { bonus, why };
}

// v3.7.0：构建达人输出级洞察——视频链接 / 主要带货类目 / 近期爆款内容方向 / 打标置信度。
// 全部基于已抓到的客观数据推断，证据不足时留空、不臆造。
function buildCreatorInsights(auth, videos, via) {
  const vids = Array.isArray(videos) ? videos : [];
  // ① 前三条视频链接（去空、去重，最多 3 条）
  const videoLinks = [];
  for (const v of vids) {
    const u = (v && v.url) || '';
    if (u && !videoLinks.includes(u)) videoLinks.push(u);
    if (videoLinks.length >= 3) break;
  }
  // ② 主要带货类目：优先行业受控词第一个；否则用主页 category 文本；再兜底从带货视频标题抽
  let mainCategory = '';
  if (Array.isArray(auth.industry) && auth.industry.length) mainCategory = auth.industry[0];
  else if (auth.category) mainCategory = String(auth.category).slice(0, 12);
  if (!mainCategory) {
    const sellTitle = vids.filter(v => v && v.isSelling).map(v => v.title || '').join(' ');
    const ind = matchControlled(sellTitle, STOCK_INDUSTRY_RULES, 1);
    if (ind.length) mainCategory = ind[0];
  }
  // ③ 近期爆款内容方向：取播放最高的一条视频，按「行业 + 内容形式」概括
  let hotDirection = '';
  const ranked = vids.slice().filter(v => v && v.title).sort((a, b) => (b.plays || 0) - (a.plays || 0));
  if (ranked.length) {
    const top = ranked[0];
    const parts = [];
    const ind = (auth.industry && auth.industry[0]) || '';
    if (ind) parts.push(ind);
    if (top.contentType) parts.push({ 测评: '测评', 种草: '种草', 口播: '口播讲解', 剧情: '剧情' }[top.contentType] || top.contentType);
    const selling = top.isSelling ? '带货' : '';
    if (selling) parts.push(selling);
    const dirText = parts.join('·') || '内容创作';
    const playTxt = top.plays >= 10000 ? `（代表作播放${top.plays >= 100000000 ? (top.plays / 1e8).toFixed(1) + '亿' : top.plays >= 10000 ? (top.plays / 1e4).toFixed(1) + '万' : top.plays}）` : '';
    hotDirection = `${dirText}${playTxt}`;
  }
  // ④ 打标置信度：依据资料来源硬指标与视频丰富度
  //   高：主页 XHR/DOM 命中且（有星川等级或消耗/交付硬指标）且至少 2 条视频
  //   中：命中主页且有任一硬指标或至少 1 条视频；或昵称搜索命中且资料较全
  //   低：仅昵称兜底、无视频、硬指标缺失
  const hardSignals = (auth.sLevel ? 1 : 0) + (auth.consumption > 0 ? 1 : 0) + (auth.deliveries > 0 ? 1 : 0) + (auth.fans > 0 ? 1 : 0);
  let confidence = '低';
  if (/home|xhr|dom/.test(String(via)) && hardSignals >= 2 && vids.length >= 2) confidence = '高';
  else if ((/home|xhr|dom/.test(String(via)) && (hardSignals >= 1 || vids.length >= 1)) || (hardSignals >= 3)) confidence = '中';
  else if (vids.length >= 2 && hardSignals >= 1) confidence = '中';
  return { videoLinks, mainCategory, hotDirection, confidence };
}


function parseListBuffer(buf, filename) {
  const wb = XLSX.read(buf, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
  if (!rows.length) return { items: [], headers: [] };
  const headers = Object.keys(rows[0]);
  const idKey = headers.find(h => /id|达人id|星图id|uid|编号/i.test(h));
  // 昵称列：跳过 ID 列（「星图达人ID」也含「达人」二字，不能误选）
  const nameKey = headers.find(h => h !== idKey && /昵称|名字|达人名|账号名|name|主播|达人/i.test(h)) ||
    headers.find(h => h !== idKey);
  // v3.5：数据来源列（存量/增量），用于双库口径打分；可缺省（默认按新达人）
  const srcKey = headers.find(h => h !== idKey && /来源|数据源|库别|所属库|名单类型/i.test(h));
  const normSrc = (v) => {
    const s = normText(v).toLowerCase();
    if (/存量|库存|已合作|stock/.test(s)) return 'stock';
    if (/增量|incr|increment/.test(s)) return 'increment';
    return '';
  };
  const items = [];
  rows.forEach((r, i) => {
    let id = idKey ? normText(r[idKey]) : '';
    let name = nameKey ? normText(r[nameKey]) : '';
    // 单列兜底：纯数字当 ID，否则当昵称
    if (!id && !name) {
      const v = normText(Object.values(r)[0]);
      if (/^\d{12,}$/.test(v)) id = v; else name = v;
    }
    const idDigits = (id.match(/\d{12,}/) || [])[0] || '';
    if (idDigits) id = idDigits;
    const src = srcKey ? normSrc(r[srcKey]) : '';
    if (id || name) items.push({ id, name, row: i + 2, src });
  });
  return { items, headers };
}

// ---- HTTP 服务 ------------------------------------------------------------
function startServer() {
  const app = express();
  app.use(cors({
    origin(origin, cb) {
      // 同源 / curl / 无 Origin（服务端到服务端）一律放行；浏览器跨域严格白名单
      if (!origin) return cb(null, true);
      if (ALLOWED_ORIGINS.includes(origin)) return cb(null, true);
      if (/^https:\/\/([a-z0-9-]+\.)?github\.io$/.test(origin)) return cb(null, true); // 任意 github.io 预览域
      if (/^http:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(origin)) return cb(null, true);
      return cb(null, false);
    },
  }));
  app.use(express.json({ limit: '5mb' }));
  app.use(express.raw({ type: () => true, limit: '15mb' })); // 接收上传文件原始字节

  // 健康检查 / 登录状态 / 浏览器模式
  // v3.4：不在 /health 里导航浏览器（前端每 2.5s 轮询，sessionWorkPage 会反复 goto /sup/）。
  // 登录态来源：①内存登录位 _loggedIn（登录成功即置位）；②对【已存活】会话做零副作用
  // URL 扫描（只遍历所有上下文所有标签页的 URL，不 evaluate、不导航），命中 xingtu.cn/sup/
  // （服务商端控制台，v3.6.2）也置位。登录成功后 CDP 浏览器已关闭、_session 为 null，此时靠内存位返回已登录。
  app.get('/health', async (_req, res) => {
    let loggedIn = !!_loggedIn;
    let browserInfo = { mode: null, ready: false, name: '', cdpPort: null };
    try {
      if (_session && await sessionAlive(_session)) {
        browserInfo = { mode: _session.mode, ready: true, name: _session.browserName, cdpPort: _session.cdpPort || null };
        if (!loggedIn) {
          let ctxs = [];
          try { ctxs = (_session.browser && typeof _session.browser.contexts === 'function') ? _session.browser.contexts() : []; } catch (_) { ctxs = []; }
          if (!ctxs.includes(_session.context)) ctxs.push(_session.context);
          outer: for (const ctx of ctxs) {
            let pages = [];
            try { pages = (ctx && typeof ctx.pages === 'function') ? ctx.pages() : []; } catch (_) { continue; }
            for (const pg of pages) {
              try {
                if (pg && !pg.isClosed() && isLoggedInUrl(pg.url() || '')) { loggedIn = true; break outer; }
              } catch (_) { /* 跳转中，下轮再测 */ }
            }
          }
        }
      }
    } catch (_) { /* 探测失败按未就绪处理 */ }
    if (loggedIn) _loggedIn = true;
    const cookieCount = readCookieFile().filter(c => /xingtu/i.test(c.domain || '')).length;
    res.json({
      ok: true, loggedIn, version: '3.7.1', port: PORT,
      loginUrl: SUP_URL,
      marketUrl: PROVIDER_MARKET_URL,
      cookiesFile: '.xc-cookies.json',
      cookieCount,
      loginPending: _loginPending,
      browser: browserInfo,
      loginNote: loggedIn ? '' :
        '请点工作台「🚀 登录星图」按钮：工具会自动打开系统 Chrome/Edge 窗口，扫码或账号密码登录均可，' +
        '登录成功后窗口会自动关闭；若电脑禁止弹窗，则在网页上显示二维码，用手机抖音 App 扫码登录。',
    });
  });

  // 登录二维码（无头模式专用）：前端登录等待中每 ~2.5s 轮询
  app.get('/login/qr', (_req, res) => {
    if (!_session) return res.json({ ok: true, mode: null, status: 'starting' });
    if (_session.mode === 'cdp') {
      return res.json({ ok: true, mode: 'cdp', status: _loginPending ? 'waiting' : 'idle', browserName: _session.browserName });
    }
    res.json({
      ok: true, mode: 'headless',
      status: (_qr && _qr.status) || 'starting',
      qr: (_qr && _qr.qrDataUrl) || null,
      kind: (_qr && _qr.kind) || null,
      browserName: _session.browserName,
    });
  });

  // 登录：POST /login
  // - 自动启动/复用浏览器会话（CDP 系统浏览器优先，无头扫码兜底）；
  // - 每 3 秒主动轮询所有上下文所有标签页的 URL（纯 URL 判定，不读页面内容），最多等 5 分钟；
  // - CDP 模式检测到登录成功后自动关闭浏览器窗口（登录态在配置目录，打标时后台自动重开）；
  // - 前端取消（请求中断）/ 浏览器断开 / 超时，都会中止并返回中文提示。
  app.post('/login', async (req, res) => {
    if (_loginPending) {
      return res.status(409).json({ ok: false, pending: true, error: '已有登录任务在等待中，请完成登录或取消后重试' });
    }
    let aborted = false;
    req.on('close', () => { aborted = true; }); // 前端点「取消」或离开页面
    _loginPending = true;
    try {
      const r = await runLogin({ get aborted() { return aborted; } });
      if (r.needInstall) {
        return res.status(503).json({ ok: false, needInstall: true, error: r.message });
      }
      res.json({
        ok: !!r.ok, loggedIn: !!r.loggedIn, pending: false, cancelled: !!r.cancelled,
        mode: r.mode || null, browserName: r.browserName || '', windowClosed: !!r.windowClosed,
        message: r.message || (r.ok ? '登录成功' : '登录失败，请重试'),
      });
    } catch (e) {
      res.status(500).json({ ok: false, error: '登录流程异常：' + friendlyErr(e) });
    } finally {
      _loginPending = false;
    }
  });

  // 解析名单文件：POST /parse?fn=xxx.xlsx ，body=文件字节
  app.post('/parse', (req, res) => {
    try {
      const fn = decodeURIComponent(req.query.fn || 'list.xlsx');
      const buf = Buffer.isBuffer(req.body) ? req.body : Buffer.from(req.body || '');
      if (!buf.length) return res.status(400).json({ ok: false, error: '未收到文件内容' });
      const { items, headers } = parseListBuffer(buf, fn);
      if (!items.length) return res.status(422).json({ ok: false, error: '未从表格中识别到达人ID或昵称列，请确认表头含「达人ID/星图ID」或「昵称/达人名」' });
      res.json({ ok: true, count: items.length, headers, items: items.slice(0, 5000) });
    } catch (e) {
      res.status(500).json({ ok: false, error: '文件解析失败：' + friendlyErr(e) });
    }
  });

  // 打标：POST /tag  body={items:[{id?,name?}]}
  app.post('/tag', async (req, res) => {
    try {
      let items = (req.body && req.body.items) || [];
      if (!Array.isArray(items) || !items.length) return res.status(400).json({ ok: false, error: 'items 为空' });
      items = items.slice(0, 1000); // 单次上限，保护本机
      let sess;
      try {
        sess = await acquireSession();
      } catch (e) {
        return res.status(503).json({ ok: false, needInstall: true, error: friendlyErr(e) });
      }
      const page = await sessionWorkPage(sess);
      // v3.5.1：优先信任内存登录位 _loggedIn（登录成功时已置 true）；
      // 仅当内存位为 false 时才用 checkLoggedIn 做 URL 二次探测，避免
      // sessionWorkPage 导航跳转瞬间误判为未登录。
      let loggedIn = !!_loggedIn;
      if (!loggedIn) {
        loggedIn = await checkLoggedIn(page);
        // 补扫所有标签页 URL（与 /health 逻辑保持一致）
        if (!loggedIn && _session && await sessionAlive(_session)) {
          let ctxs = [];
          try { ctxs = (_session.browser && typeof _session.browser.contexts === 'function') ? _session.browser.contexts() : []; } catch (_) { ctxs = []; }
          if (!ctxs.includes(_session.context)) ctxs.push(_session.context);
          outer2: for (const ctx of ctxs) {
            let pages = [];
            try { pages = (ctx && typeof ctx.pages === 'function') ? ctx.pages() : []; } catch (_) { continue; }
            for (const pg of pages) {
              try {
                if (pg && !pg.isClosed() && isLoggedInUrl(pg.url() || '')) { loggedIn = true; break outer2; }
              } catch (_) { /* skip */ }
            }
          }
        }
      }
      if (!loggedIn) {
        return res.status(401).json({
          ok: false, loggedIn: false, mode: sess.mode,
          error: sess.mode === 'headless'
            ? '星图未登录：请点工作台「🚀 登录星图」按钮，用手机抖音 App 扫描网页上显示的二维码登录后再打标'
            : '星图未登录：请点工作台「🚀 登录星图」按钮，在自动打开的 Chrome/Edge 窗口里扫码或账号密码登录（成功后窗口会自动关闭）后再打标',
        });
      }
      _loggedIn = true;

      const results = [];
      const debugVideo = process.env.XC_TAGGER_DEBUG === '1';
      if (debugVideo) console.log('🐞 XC_TAGGER_DEBUG=1：达人主页抓取将落截图、接口 JSON 与 DOM 状态到 .xc-debug/');
      // v3.6 主链路：① 有 ID 直接进达人主页抓资料+视频；
      //             ② 主页打不开（ID 无效/未入驻/重定向）才用昵称搜达人广场兜底；
      //             ③ 两条路都失败才报「未检索到」。
      const homeCache = new Map(); // authorId -> 主页抓取结果（避免同一 ID 重复访问）
      const homeFor = async (id) => {
        const key = String(id);
        if (!homeCache.has(key)) {
          let r;
          try { r = await fetchCreatorHome(page, key, { debug: debugVideo }); }
          catch (_) { r = { ok: false, auth: null, videos: [], via: '', reason: 'error' }; }
          homeCache.set(key, r);
          await sleep(800); // 防风控：达人主页访问间隔
        }
        return homeCache.get(key);
      };
      for (const it of items) {
        let auth = null;
        let videoAnalysis = [];
        let via = '';
        // ① 主链路：用达人 ID 直进星图主页
        if (it.id) {
          const home = await homeFor(it.id);
          if (home && home.ok && home.auth) {
            auth = home.auth;
            videoAnalysis = home.videos || [];
            via = 'home';
          }
        }
        // ② 兜底：昵称搜索达人广场（主页打不开 / 名单只有昵称）
        if (!auth) {
          const kw = it.name || it.id;
          if (kw) {
            let searched = null;
            try { searched = await searchByName(page, String(kw)); } catch (_) { searched = null; }
            await sleep(350);
            if (searched) {
              auth = searched;
              via = 'search';
              // 搜索命中后仍进一次主页：抓视频 + 用主页资料补空字段（同 ID 走缓存不重复访问）
              if (searched.id) {
                const home = await homeFor(searched.id);
                if (home) {
                  videoAnalysis = home.videos || [];
                  if (home.auth) auth = mergeAuth(searched, home.auth);
                }
              }
            }
          }
        }
        // ③ 两条路都失败
        if (!auth) {
          results.push({
            id: it.id || '', name: it.name || '(未命名)', found: false, score: 0, tier: '储备', medal: '⚪',
            sLevel: '', lLevel: '', deliveries: 0, consumption: 0, fans: 0, fansTier: '',
            persona: [], forms: [], style: [], scene: [], industry: [], category: '',
            videoLinks: [], mainCategory: '', hotDirection: '', confidence: '低',
            homeUrl: it.id ? PROVIDER_AUTHOR_URL(it.id) : '',
            tags: ['未检索到'],
            reason: '未能通过 ID 进入主页，昵称搜索也未命中，请核实达人是否入驻星图',
            dataSource: 'new', videoBonus: 0, videoAnalysis: [],
          });
          continue;
        }
        // v3.5：双库口径（名单「来源」列：存量=stock/增量=increment；缺省按新达人 new）
        const dataSource = it.src === 'stock' || it.src === 'increment' ? it.src : 'new';
        enrichAuthFromVideos(auth, videoAnalysis);
        const sc = scoreAuthor(auth, { dataSource, videoAnalysis });
        // v3.7.0：视频链接 / 主要带货类目 / 爆款内容方向 / 置信度
        const insights = buildCreatorInsights(auth, videoAnalysis, via);
        const reason = sc.reason +
          (via === 'home' ? '\n获取方式: ID直进达人主页' : '\n获取方式: 昵称搜索兜底') +
          (insights.mainCategory ? `\n主要带货类目: ${insights.mainCategory}` : '') +
          (insights.hotDirection ? `\n近期爆款方向: ${insights.hotDirection}` : '') +
          ((auth.style || []).length ? `\n画面风格: ${auth.style.join(' ')}` : '') +
          ((auth.scene || []).length ? `\n拍摄场景: ${auth.scene.join(' ')}` : '') +
          `\n打标置信度: ${insights.confidence}`;
        results.push({
          id: auth.id || it.id || '', name: auth.name || it.name || '', found: true,
          score: sc.score, tier: sc.tier, medal: sc.medal,
          sLevel: auth.sLevel, lLevel: auth.lLevel,
          deliveries: auth.deliveries, consumption: auth.consumption,
          fans: auth.fans, fansTier: auth.fansTier || '',
          persona: auth.persona || [], forms: auth.forms || [],
          style: auth.style || [], scene: auth.scene || [],
          industry: auth.industry || [],
          category: auth.category || '',
          videoLinks: insights.videoLinks,
          mainCategory: insights.mainCategory,
          hotDirection: insights.hotDirection,
          confidence: insights.confidence,
          homeUrl: PROVIDER_AUTHOR_URL(auth.id || it.id || ''),
          tags: sc.tags, reason,
          dataSource: sc.dataSource, videoBonus: sc.videoBonus, videoAnalysis,
          via,
        });
      }
      // 按分降序
      results.sort((a, b) => b.score - a.score);
      const summary = { 标杆: 0, 主力: 0, 潜力: 0, 储备: 0 };
      results.forEach(r => { summary[r.tier] = (summary[r.tier] || 0) + 1; });
      res.json({ ok: true, loggedIn: true, total: results.length, summary, results });
    } catch (e) {
      res.status(500).json({ ok: false, error: '打标失败：' + friendlyErr(e) });
    }
  });

  app.listen(PORT, HOST, () => {
    console.log('\n==================================================');
    console.log('  星川服务商达人自助打标 · 本地工具已启动 v3.7.1');
    console.log(`  本地服务：http://${HOST}:${PORT}`);
    console.log('  工作台网页：https://didimarco26.github.io/xingchuan-workbench/');
    console.log('--------------------------------------------------');
    console.log('  使用 3 步：');
    console.log('   1) 保持本窗口打开；');
    console.log('   2) 在工作台网页点「🚀 登录星图」——工具会自动打开系统');
    console.log('      Chrome/Edge 窗口，扫码或账号密码登录均可，登录成功后');
    console.log('      窗口会自动关闭（登录态保存在本机，仅需一次）；若电脑');
    console.log('      禁止弹窗，网页会显示二维码，用手机抖音 App 扫码登录；');
    console.log('   3) 上传达人名单 Excel，点「开始打标」等待结果。');
    console.log('==================================================\n');
    // 启动后自动准备浏览器会话（CDP 模式会自动打开 Chrome/Edge 窗口）
    setTimeout(() => {
      acquireSession().then(async (sess) => {
        const page = await sessionWorkPage(sess);
        const in_ = await checkLoggedIn(page);
        if (in_) _loggedIn = true;
        console.log(in_
          ? `✅ 检测到星图登录态有效，可直接使用（浏览器：${sess.browserName}）。\n`
          : `ℹ️  尚未登录：请到工作台点「🚀 登录星图」，在弹出的浏览器窗口里扫码或账号密码登录（当前浏览器：${sess.browserName}）。\n`);
      }).catch(e => console.log('⚠️  浏览器自动启动失败：' + friendlyErr(e) + '\n   点「登录星图」时会自动重试。\n'));
    }, 800);
  });
}

// 直接运行时启动服务；被 require（单测）时仅导出纯逻辑函数
if (require.main && require.main.filename === __filename) {
  startServer();
}

module.exports = {
  browserCandidates,
  pickFreePort,
  netTcpPortFree,
  checkLoggedIn,
  isLoggedInUrl,
  captureLoginQr,
  connectExistingCdp,
  cdpProbeOnce,
  findReusablePage,
  tierOf,
  scoreConsumption,
  scoreDeliveries,
  scoreEcomLevel,
  S_LEVEL_SCORE,
  QR_SELECTORS,
  CDP_CANDIDATE_PORTS,
  CHROME_PROFILE_DIR,
  SQUARE_URL,
  SUP_URL,
  LOGIN_URL,
  LOGIN_WAIT_MS,
  LOGIN_POLL_MS,
  // v3.5
  scoreAuthor,
  scoreVideoBonus,
  inferVideoContentType,
  videoIsSelling,
  buildVideoUrl,
  buildCreatorInsights,
  STOCK_STYLE_RULES,
  STOCK_SCENE_RULES,
  parseMediaCount,
  deepCollectVideos,
  enrichAuthFromVideos,
  CREATOR_DETAIL_URL,
  // v3.6
  parseDomProfile,
  pickHomeAuthor,
  authorFullness,
  mergeAuth,
  // v3.7.1
  PROVIDER_MARKET_URL,
  MARKET_DETAIL_URL,
  enterCreatorDetailViaMarket,
  waitIfCaptcha,
  captchaPresentEval,
  // v3.7.2
  PROVIDER_AUTHOR_URL,
  enterProviderAuthorPage,
  pickShowItemVideos,
};

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
function friendlyErr(e) { return (e && (e.message || String(e))) || '未知错误'; }
