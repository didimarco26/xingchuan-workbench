#!/usr/bin/env node
/* eslint-disable */
/**
 * 星川服务商达人自助打标 · 本地一键工具 (xc-tagger) v3.2
 * ------------------------------------------------------------------
 * 用途：服务商在本机运行本工具，它会：
 *   1) 在 127.0.0.1:7842 起一个本地 HTTP 服务（只监听本机，不对外）；
 *   2) 浏览器方案：启动脚本（start-xc-tagger / xc-open-chrome）先用
 *      【系统 Chrome/Edge】以远程调试端口启动一个独立浏览器实例：
 *        --remote-debugging-port=9222 --user-data-dir=<工具目录>/.xc-chrome-profile
 *      工具本身【绝不 launch 浏览器】，只用 connectOverCDP 附加到这个实例
 *      （Windows 安全策略会立即关闭 Playwright 自己弹出的 Chromium 窗口，
 *        exitCode=0；系统 Chrome 由启动脚本正常拉起，不受此限制）。
 *   3) 登录（仅需一次）：网页点「🚀 登录星图」→ POST /login，工具在已连接的
 *      Chrome 里打开巨量星图达人广场标签页，服务商在该窗口扫码/验证码登录；
 *      工具每 2 秒检测一次，成功后 Cookie 写入 .xc-cookies.json 备份
 *      （登录态主体保存在 .xc-chrome-profile 浏览器配置目录，长期免登录）。
 *      浏览器窗口不自动关闭，由用户自行关闭；超时 180 秒 / 点「取消」中止。
 *   4) 打标：POST /tag 复用同一条 CDP 连接，在同一个 Chrome 里开后台标签页
 *      抓取（标签页可见，可看到工具在工作；关闭浏览器窗口即停止）。
 *   5) 网页（星川决策工作台·服务商 GitHub 版）调用本地接口：
 *        GET  /health   → 探测服务、CDP 连接与登录状态（含 loginPending）
 *        POST /login    → 在 Chrome 中打开星图页并等待扫码登录
 *        POST /parse    → 解析上传的 Excel/CSV 达人名单
 *        POST /tag      → 用星图登录态抓取达人信息并按存量逻辑分层打标
 *   若 9222 端口连不上（Chrome 未用调试端口启动），/login 返回
 *      needOpenChrome:true，提示用户双击包里的 xc-open-chrome 脚本补开。
 *   Cookie 只留在服务商本机，不上传、不落库。
 *
 * 运行：解压后双击 start-xc-tagger.command(Mac) / start-xc-tagger.bat(Windows)
 *      （脚本会自动打开调试 Chrome 并启动本服务；需已安装 Node.js 18+ 与
 *        Google Chrome 或 Edge）。
 *
 * 安全：CORS 仅放行 didimarco26.github.io 与本机页面；服务只绑定 127.0.0.1。
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
const SQUARE_URL = 'https://www.xingtu.cn/ad/creator/square';
// 巨量引擎统一登录直达页（role=1 客户侧）：未登录访问达人广场会被踢到此 SSO。
// 直接导航到它可省去「营销首页→点登录→选客户角色」几步点击，真机/无头都更稳；
// 登录成功后 SSO 自动回跳 redirect_uri（达人广场）。
const LOGIN_URL = 'https://sso.oceanengine.com/xingtu/login?redirect_uri=' +
  encodeURIComponent('/ad/creator/square') + '&role=1';
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
    SQUARE_URL,
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

// ① CDP 模式：逐个候选浏览器尝试 启动→连接，全部失败返回 null
async function tryCdpSession() {
  const cands = browserCandidates();
  if (!cands.length) {
    console.log('ℹ️  未探测到系统 Chrome/Edge，将使用无头扫码模式。');
    return null;
  }
  const port = await pickFreePort(CDP_CANDIDATE_PORTS);
  if (!port) {
    console.log('ℹ️  调试端口 9222-9225 均被占用，将使用无头扫码模式。');
    return null;
  }
  const playwright = require('playwright');
  for (const cand of cands) {
    console.log(`🌐 尝试启动 ${cand.name}（调试端口 ${port}）…`);
    let launched = null;
    try { launched = await launchSystemBrowser(cand, port, CHROME_PROFILE_DIR); } catch (_) { launched = null; }
    if (!launched) { console.log(`   ✗ ${cand.name} 启动失败或调试端口无响应，尝试下一个浏览器。`); continue; }
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
  }
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
 * 获取会话内的星图工作标签页：复用已打开的星图页，没有则新开并跳转达人广场。
 */
async function sessionWorkPage(sess) {
  if (sess._workPage && !sess._workPage.isClosed()) return sess._workPage;
  const ctx = sess.context;
  const pages = ctx.pages ? ctx.pages() : [];
  let page = pages.find(p => { try { return /xingtu\.cn/.test(p.url()); } catch (_) { return false; } }) || null;
  if (!page) page = await ctx.newPage();
  try { await page.setViewportSize({ width: 1280, height: 860 }); } catch (_) {}
  let url = '';
  try { url = page.url() || ''; } catch (_) {}
  if (!/xingtu\.cn/.test(url)) {
    try {
      await page.goto(SQUARE_URL, { waitUntil: 'domcontentloaded', timeout: 45000 });
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
 * 无头模式：导航到 SSO 登录页并切到「抖音扫码」。
 * 直达 LOGIN_URL（省去营销首页点登录/选角色），再点「其他方式-抖音」图标
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

// 是否已登录：必须真正进入星图「创作者/广告主」区域才算已登录——
//   · URL 含 redirect_uri / passport / sso / login（被踢去登录）→ 未登录；
//   · URL 不在 xingtu.cn 或不含 /ad/ 路径（停在营销首页 www.xingtu.cn/）→ 未登录；
//   · 页面空白（加载失败/网络异常）→ 未登录（宁可让用户重新扫码，不可误判已登录）；
//   · 页面出现扫码/手机号登录入口且无达人广场特征 → 未登录。
// 已登录时自动把最新 Cookie 回写文件（刷新有效期），实现长期免登录。
async function checkLoggedIn(page) {
  try {
    const url = page.url() || '';
    if (/redirect_uri|passport|sso|login/i.test(url)) return false;
    if (!/xingtu\.cn\//.test(url)) return false;
    if (!/\/ad\//.test(url)) return false;
    const loggedIn = await page.evaluate(() => {
      const t = document.body ? document.body.innerText : '';
      if (t.trim().length < 40) return false; // 空白页/加载中，不能判定已登录
      // 出现明显「扫码登录 / 手机号登录」入口且无达人广场特征，判为未登录
      const hasLoginBtn = /扫码登录|登录巨量星图|手机号登录|验证码登录|免费登录/.test(t) && !/达人广场|找达人|达人榜单/.test(t);
      return !hasLoginBtn;
    });
    if (loggedIn) {
      try {
        const cookies = await page.context().cookies();
        if (Array.isArray(cookies) && cookies.length) {
          fs.writeFileSync(COOKIES_FILE, JSON.stringify(cookies, null, 2), 'utf8');
        }
      } catch (_) { /* 回写失败不影响判定 */ }
    }
    return loggedIn;
  } catch { return false; }
}

// ---- 登录（在已 CDP 连接的系统 Chrome 里打开星图标签页，等待扫码）----------
// /login 正在等待扫码时为 true，/health 据此让前端显示「等待扫码中…」。
let _loginPending = false;
const LOGIN_WAIT_MS = 180000; // 扫码最长等待 3 分钟

/**
 * 统一登录流程（自动适配浏览器会话模式）：
 * - CDP 模式：系统 Chrome/Edge 窗口已打开星图页，服务商在窗口里扫码；
 * - 无头模式：工具截取登录页二维码写入 _qr，前端 GET /login/qr 展示，
 *   服务商在网页上用手机抖音 App 扫码。
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
  // 取一个工作标签页（复用已打开的；登录成功后它就是登录态广场页，供打标复用）
  let page = sess._workPage;
  if (!page || page.isClosed()) {
    page = await sess.context.newPage().catch(() => null);
    if (!page) return { ok: false, loggedIn: false, message: '无法打开浏览器标签页，请重试。', mode: sess.mode };
    sess._workPage = page;
    try { await page.setViewportSize({ width: 1360, height: 950 }); } catch (_) {}
  }

  // 已登录快速返回（页面在 /ad/ 广场区域）
  if (await checkLoggedIn(page)) {
    writeCookieFile(await sess.context.cookies().catch(() => []));
    _qr = { status: 'loggedin', qrDataUrl: null, updatedAt: Date.now() };
    return { ok: true, loggedIn: true, mode: sess.mode, browserName: sess.browserName, message: '已是登录状态' };
  }

  // 未登录：导航到直达 SSO 登录页
  if (sess.mode === 'cdp') {
    try {
      const u = (() => { try { return page.url(); } catch (_) { return ''; } })();
      if (!/sso\.oceanengine\.com|\/ad\//.test(u)) await page.goto(LOGIN_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await page.bringToFront();
    } catch (_) { /* 导航失败也让用户在可见窗口里手动操作 */ }
    console.log('🟢 请在已打开的 ' + sess.browserName + ' 窗口中登录巨量星图（可抖音扫码或手机验证码）…');
  } else {
    console.log('🟢 无头模式：正在打开抖音扫码登录，请在工作台网页用手机抖音 App 扫码…');
    _qr = { status: 'starting', qrDataUrl: null, updatedAt: Date.now() };
    await gotoDouyinQr(page);
  }

  const deadline = Date.now() + LOGIN_WAIT_MS;
  let lastReload = Date.now();
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
    let logged = false;
    try { logged = await checkLoggedIn(page); } catch (_) { /* 页面跳转中，下轮再测 */ }
    if (logged) {
      writeCookieFile(await sess.context.cookies().catch(() => []));
      _qr = { status: 'loggedin', qrDataUrl: null, updatedAt: Date.now() };
      sess._workPage = page; // 登录态广场页，打标直接复用
      console.log('✅ 登录成功，Cookie 已备份到 .xc-cookies.json。');
      return { ok: true, loggedIn: true, mode: sess.mode, browserName: sess.browserName, message: '登录成功' };
    }
    await sleep(2200);
  }
  return {
    ok: false, loggedIn: false, mode: sess.mode,
    message: sess.mode === 'headless'
      ? '登录超时（3 分钟未扫码成功）：请用手机抖音 App 扫描网页上的二维码后重试。'
      : '登录超时（3 分钟未检测到登录成功）：请在打开的浏览器窗口里完成巨量星图扫码后重试。',
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
  return { id, name, fans, fansTier, sLevel, lLevel, deliveries, consumption, persona, forms, industry, category, raw };
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

// 按 ID 批量取达人信息（50 个/批）
async function getByIds(page, ids) {
  const result = new Map();
  const payloads = [
    { author_ids: ids },
    { star_author_ids: ids },
    { author_id_list: ids },
    { ids },
  ];
  for (const body of payloads) {
    try {
      const r = await xgFetch(page, '/gw/api/aggregator/multi_get_author_info', body);
      if (DEBUG_DUMP) fs.writeFileSync(path.join(__dirname, `debug_ids_${Date.now()}.json`), JSON.stringify(r.json, null, 2));
      const authors = [];
      deepCollectAuthors(r.json, authors, 0);
      for (const a of authors.map(mapAuthor).filter(x => x.id || x.name)) {
        const key = a.id || a.name;
        if (key && !result.has(key)) result.set(key, a);
      }
      if (result.size) return result;
    } catch (_) { /* 试下一种 payload */ }
  }
  return result;
}

// 对单个达人打标（输出字段与存量达人库口径对齐）
function scoreAuthor(auth) {
  const sScore = S_LEVEL_SCORE[auth.sLevel] != null ? S_LEVEL_SCORE[auth.sLevel] : 0;
  const dScore = scoreDeliveries(auth.deliveries);
  const cScore = scoreConsumption(auth.consumption);
  const lScore = scoreEcomLevel(auth.lLevel);
  const score = Math.round((sScore + dScore + cScore + lScore) * 10) / 10;
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
  if (!tags.length) tags.push('待培育');
  const reason =
    `星川${auth.sLevel || '未分级'}(${sScore}分) · 星图消耗${fmtWan(auth.consumption)}(${cScore}分) · ` +
    `交付${auth.deliveries}个项目(${dScore}分) · 电商${auth.lLevel || '—'}(${lScore}分)` +
    (auth.fansTier ? ` · ${auth.fansTier}` : '') +
    (auth.industry && auth.industry.length ? ` · 行业${auth.industry.join('/')}` : '') +
    (auth.category ? ` · 类目${auth.category}` : '');
  return { score, tier, medal, tags, reason };
}

// ---- 名单解析（Excel / CSV）-----------------------------------------------
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
    if (id || name) items.push({ id, name, row: i + 2 });
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
  app.get('/health', async (_req, res) => {
    let loggedIn = false;
    let browserInfo = { mode: null, ready: false, name: '', cdpPort: null };
    try {
      if (_session && await sessionAlive(_session)) {
        browserInfo = { mode: _session.mode, ready: true, name: _session.browserName, cdpPort: _session.cdpPort || null };
        const page = await sessionWorkPage(_session);
        loggedIn = await checkLoggedIn(page);
      }
    } catch (_) { /* 探测失败按未登录处理 */ }
    const cookieCount = readCookieFile().filter(c => /xingtu/i.test(c.domain || '')).length;
    res.json({
      ok: true, loggedIn, version: '3.2.0', port: PORT,
      loginUrl: SQUARE_URL,
      cookiesFile: '.xc-cookies.json',
      cookieCount,
      loginPending: _loginPending,
      browser: browserInfo,
      loginNote: loggedIn ? '' :
        '请点工作台「🚀 登录星图」按钮：工具会自动打开系统 Chrome/Edge 窗口扫码；' +
        '若电脑无可用浏览器，则在网页上显示二维码，用手机抖音 App 扫码登录。',
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

  // 扫码登录：POST /login
  // - 自动启动/复用浏览器会话（CDP 系统浏览器优先，无头扫码兜底）；
  // - 每 2.2 秒检测一次登录态，最多等 180 秒；
  // - 前端取消（请求中断）/ 浏览器断开 / 超时，都会中止并返回中文提示。
  app.post('/login', async (req, res) => {
    if (_loginPending) {
      return res.status(409).json({ ok: false, pending: true, error: '已有登录任务在等待扫码，请完成登录或取消后重试' });
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
        mode: r.mode || null, browserName: r.browserName || '',
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
      const loggedIn = await checkLoggedIn(page);
      if (!loggedIn) {
        return res.status(401).json({
          ok: false, loggedIn: false, mode: sess.mode,
          error: sess.mode === 'headless'
            ? '星图未登录：请点工作台「🚀 登录星图」按钮，用手机抖音 App 扫描网页上显示的二维码登录后再打标'
            : '星图未登录：请点工作台「🚀 登录星图」按钮，在自动打开的 Chrome/Edge 窗口里扫码登录后再打标',
        });
      }

      const results = [];
      // 1) 先按 ID 批量取（50/批）
      const withId = items.filter(x => x.id);
      const byName = items.filter(x => !x.id && x.name);
      const idMap = new Map();
      for (let i = 0; i < withId.length; i += 50) {
        const batch = withId.slice(i, i + 50).map(x => String(x.id));
        const got = await getByIds(page, batch);
        for (const [k, v] of got) idMap.set(String(k), v);
        await sleep(400);
      }
      // 2) 有 ID 但批量接口没取到的，以及只有昵称的，走搜索
      for (const it of items) {
        let auth = null;
        if (it.id && idMap.has(String(it.id))) auth = idMap.get(String(it.id));
        if (!auth) {
          const kw = it.name || it.id;
          if (kw) { try { auth = await searchByName(page, String(kw)); } catch (_) { auth = null; } await sleep(350); }
        }
        if (!auth) {
          results.push({ id: it.id || '', name: it.name || '(未命名)', found: false, score: 0, tier: '储备', medal: '⚪', sLevel: '', lLevel: '', deliveries: 0, consumption: 0, tags: ['未检索到'], reason: '未在星图达人广场检索到，请核对昵称/ID 或该达人是否入驻星图' });
          continue;
        }
        const sc = scoreAuthor(auth);
        results.push({
          id: auth.id || it.id || '', name: auth.name || it.name || '', found: true,
          score: sc.score, tier: sc.tier, medal: sc.medal,
          sLevel: auth.sLevel, lLevel: auth.lLevel,
          deliveries: auth.deliveries, consumption: auth.consumption,
          fans: auth.fans, fansTier: auth.fansTier || '',
          persona: auth.persona || [], forms: auth.forms || [], industry: auth.industry || [],
          category: auth.category || '',
          tags: sc.tags, reason: sc.reason,
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
    console.log('  星川服务商达人自助打标 · 本地工具已启动 v3.2.0');
    console.log(`  本地服务：http://${HOST}:${PORT}`);
    console.log('  工作台网页：https://didimarco26.github.io/xingchuan-workbench/');
    console.log('--------------------------------------------------');
    console.log('  使用 3 步：');
    console.log('   1) 保持本窗口打开；');
    console.log('   2) 在工作台网页点「🚀 登录星图」——工具会自动打开系统');
    console.log('      Chrome/Edge 窗口扫码；若无可用浏览器，网页会显示二维码，');
    console.log('      用手机抖音 App 扫码登录（登录态保存在本机，仅需一次）；');
    console.log('   3) 上传达人名单 Excel，点「开始打标」等待结果。');
    console.log('==================================================\n');
    // 启动后自动准备浏览器会话（CDP 模式会自动打开 Chrome/Edge 窗口）
    setTimeout(() => {
      acquireSession().then(async (sess) => {
        const page = await sessionWorkPage(sess);
        const in_ = await checkLoggedIn(page);
        console.log(in_
          ? `✅ 检测到星图登录态有效，可直接使用（浏览器：${sess.browserName}）。\n`
          : `ℹ️  尚未登录：请到工作台点「🚀 登录星图」扫码（当前浏览器：${sess.browserName}）。\n`);
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
  captureLoginQr,
  tierOf,
  scoreConsumption,
  scoreDeliveries,
  scoreEcomLevel,
  S_LEVEL_SCORE,
  QR_SELECTORS,
  CDP_CANDIDATE_PORTS,
  CHROME_PROFILE_DIR,
  SQUARE_URL,
  LOGIN_URL,
};

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
function friendlyErr(e) { return (e && (e.message || String(e))) || '未知错误'; }
