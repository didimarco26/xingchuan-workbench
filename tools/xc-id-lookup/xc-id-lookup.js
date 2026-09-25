#!/usr/bin/env node
/* eslint-disable */
/**
 * 星川服务商达人信息互查 · 本地一键工具 (xc-id-lookup) v1.0.0
 * ------------------------------------------------------------------
 * 用途：服务商在本机运行本工具，上传达人 Excel（列可以是
 *   【抖音号 / 星图ID / 达人名称】中的任意一列或多列，不必齐全），
 *   工具复用本机已登录星图的浏览器，按行批量检索，为每一行补全三元组：
 *     抖音号 + 星图ID + 达人名称
 *   完成后可导出完整对照表（Excel / CSV）。本工具【只做身份补全，不打分评级】。
 *
 * 三路检索优先级（命中即止）：
 *   ① 星图ID → 直接进达人主页 /xingtu/author/{id}/，取 unique_id（抖音号）与昵称；
 *   ② 抖音号 → POST /gw/api/gsearch/basic_search_authors，seach_type=101 精确检索；
 *   ③ 达人名称 → POST /gw/api/gsearch/basic_search_authors，seach_type=0 模糊检索，取第一条。
 * 模糊结果过滤：有名称时第一条结果的名称须与给定名称一致才采用；无名称时信任排序取第一条。
 *
 * 浏览器架构（与 xc-tagger 一致）：
 *   1) 在 127.0.0.1:7843 起本地 HTTP 服务（只监听本机，不对外），自带网页界面；
 *   2) CDP 优先：用系统 Chrome/Edge 以远程调试端口启动独立实例并 connectOverCDP
 *      附加，也会先附加已在运行的调试浏览器（9222-9225）；窗口真实可见，服务商在窗口里登录；
 *   3) 登录仅需一次：网页点「🚀 登录星图」，纯 URL 判定（xingtu.cn 下落点 /sup//provider/），
 *      扫码/账号密码/验证码均可；成功写 Cookie 备份、CDP 模式自动关窗；超时 5 分钟；
 *   4) 兜底：系统浏览器不可用/调试端口被禁 → 无头扫码模式（网页显示抖音二维码）。
 *
 * 运行：解压后双击 start-xc-id-lookup.command(Mac) / start-xc-id-lookup.bat(Windows)。
 * 安全：CORS 仅放行 github.io 与本机页面；服务只绑定 127.0.0.1；Cookie 只留本机，不上传。
 */

'use strict';

const path = require('path');
const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const fs = require('fs');

// ---- 配置 ----------------------------------------------------------------
const VERSION = '1.0.0';
const PORT = 7843;
const HOST = '127.0.0.1';
// 登录态文件与独立配置目录（与 xc-tagger 各自独立，可共存于同一台电脑）
const COOKIES_FILE = path.resolve(__dirname, '.xc-id-cookies.json');
const CHROME_PROFILE_DIR = path.resolve(__dirname, '.xc-id-chrome-profile');
const PUBLIC_DIR = path.join(__dirname, 'public');
const DEBUG_DUMP = process.env.XC_ID_DEBUG === '1';
const DEBUG_DIR = path.resolve(__dirname, '.xc-id-debug');

const XINGTU_ORIGIN = 'https://www.xingtu.cn';
// 服务商端控制台首页 = 登录引导/判定页（服务商账号必须从 /sup/ 端建立会话）
const SUP_URL = 'https://www.xingtu.cn/sup/';
// 纯服务商账号控制台实际落在 /provider/（广场/详情页）
const PROVIDER_MARKET_URL = 'https://www.xingtu.cn/provider/pages/market';
const LOGIN_URL = SUP_URL;
// 达人主页（星图ID 直达）：取 unique_id（抖音号）与昵称
const AUTHOR_URL = (id) => `${XINGTU_ORIGIN}/xingtu/author/${id}`;
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36';

// 仅允许以下来源的网页调用本地服务（安全限制）
const ALLOWED_ORIGINS = [
  'https://didimarco26.github.io',
  'https://xingchuan-advisor.surge.sh',
  'http://localhost:5173',
  'http://127.0.0.1:5173',
  'http://localhost:8000',
  'null', // 部分浏览器 file:// Origin 为 'null'
];

function debugLog(...a) {
  if (DEBUG_DUMP) console.log(...a);
}
function dumpDebug(name, obj) {
  if (!DEBUG_DUMP) return;
  try {
    if (!fs.existsSync(DEBUG_DIR)) fs.mkdirSync(DEBUG_DIR, { recursive: true });
    fs.writeFileSync(path.join(DEBUG_DIR, `${Date.now()}_${name}.json`), JSON.stringify(obj, null, 2));
  } catch (_) { /* 落盘失败不影响主流程 */ }
}

// ---- 浏览器会话管理（CDP 优先，无头扫码兜底）------------------------------
const CDP_CANDIDATE_PORTS = [9222, 9223, 9224, 9225];

let _session = null;          // { mode:'cdp'|'headless', browser, context, browserName, browserId, cdpPort, proc, _workPage }
let _sessionAcquiring = null;

/**
 * 探测本机可用浏览器（纯函数，按优先级返回已存在的候选，便于单测注入）。
 * 优先级：系统 Chrome → Beta → Canary → Edge → 其他 Chromium。
 */
function browserCandidates(opt = {}) {
  const platform = opt.platform || process.platform;
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

// 单次探测某端口是否已有 CDP 调试服务（快速失败）
async function cdpProbeOnce(port, timeoutMs = 1200) {
  try {
    const r = await fetch(`http://${HOST}:${port}/json/version`, { signal: AbortSignal.timeout(timeoutMs) });
    if (r.ok) return await r.json();
  } catch (_) { /* 端口无服务 */ }
  return null;
}

// 附加到【已经在运行】的调试 Chrome/Edge（9222-9225 逐个探测），直接复用其登录态
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
      console.log(`   ✅ 已附加正在运行的调试浏览器（CDP 端口 ${port}，${brand}），直接复用登录态。`);
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
 * 用指定浏览器启动带远程调试端口的独立实例。
 * spawn 以参数数组传参（不经 shell），路径含中文/空格也安全。
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

// ① CDP 模式：先附加已运行调试浏览器；没有再逐个候选【启动→连接】，全失败返回 null
async function tryCdpSession() {
  const playwright = require('playwright');

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
        console.log(`   ✗ ${cand.name} 启动失败或调试口无响应（常见于配置目录被已有调试浏览器占用）。`);
      }
    }
    // spawn 失败/端口被占 → 很可能命令行已转发给持锁旧实例 → 再试附加
    try {
      const existing = await connectExistingCdp(playwright);
      if (existing) return existing;
    } catch (_) { /* 继续下一候选 */ }
  }
  console.log('ℹ️  CDP 模式不可用，将使用无头扫码模式。');
  return null;
}

// ② 无头扫码模式：headless launch（优先系统浏览器，全无时兜底内置 Chromium）
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
      console.log(`🌐 无头模式：使用 ${cand.name}（后台运行，登录请扫网页上的二维码）。`);
      const context = await browser.newContext({ userAgent: UA, viewport: { width: 1280, height: 900 } });
      return { mode: 'headless', browser, context, browserName: cand.name + '（无头）', browserId: cand.id, cdpPort: null, proc: null, _workPage: null };
    } catch (e) {
      console.log(`   ✗ 无头启动 ${cand.name} 失败：${friendlyErr(e)}`);
    }
  }
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

// 获取（或复用）浏览器会话：CDP 优先，失败自动降级无头；断开自动重建
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
 * 在浏览器上下文里找可复用的标签页（避免重复开 SSO 页）。
 * 优先级：/sup/ 控制台页 → 任意星图页 → SSO/登录页 → null。
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

// 获取会话内的星图工作标签页：复用已有页，没有才新开；并确保落在 /sup//provider/ 登录区
async function sessionWorkPage(sess) {
  if (sess._workPage && !sess._workPage.isClosed()) return sess._workPage;
  const ctx = sess.context;
  let page = findReusablePage(ctx);
  if (!page) page = await ctx.newPage();
  try { await page.setViewportSize({ width: 1280, height: 860 }); } catch (_) {}
  let url = '';
  try { url = page.url() || ''; } catch (_) {}
  if (!(/xingtu\.cn/.test(url) && /\/(?:sup|provider)(?:\/|$)/.test(url))) {
    await convergeLoggedInPage(page).catch(() => {});
  } else {
    try { await page.bringToFront(); } catch (_) {}
  }
  sess._workPage = page;
  page.on('close', () => { if (sess._workPage === page) sess._workPage = null; });
  return page;
}

// ---- 通用工具 -------------------------------------------------------------
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function friendlyErr(e) {
  try {
    const s = (e && (e.message || e.toString())) || String(e);
    return s.split('\n')[0].slice(0, 220);
  } catch (_) { return String(e).slice(0, 220); }
}

// 清洗输入：去空格、去首尾 @（抖音号/名称比较前统一走这里）
function cleanHandle(v) {
  if (v == null) return '';
  return String(v).replace(/\s+/g, '').replace(/^@+/, '').replace(/@+$/, '').trim();
}
// 名称比较（忽略大小写、空格、@）
function nameEqual(a, b) {
  return cleanHandle(a).toLowerCase() === cleanHandle(b).toLowerCase();
}
function normText(v) {
  if (v == null) return '';
  return String(v).trim();
}
// 纯数字串（星图ID）
function keepDigits(v) {
  const m = normText(v).match(/\d+/);
  return m ? m[0] : '';
}

// ---- Cookie 备份 / 恢复 ----------------------------------------------------
function readCookieFile() {
  try {
    const arr = JSON.parse(fs.readFileSync(COOKIES_FILE, 'utf8'));
    return Array.isArray(arr) ? arr.filter(c => c && c.name && c.value != null) : [];
  } catch (_) { return []; }
}
async function addCookiesSafe(context, cookies) {
  let n = 0;
  for (const c of cookies) {
    try {
      await context.addCookies([{
        name: c.name,
        value: c.value,
        domain: c.domain,
        path: c.path || '/',
        expires: c.expires || -1,
        httpOnly: !!c.httpOnly,
        secure: !!c.secure,
        sameSite: c.sameSite || 'Lax',
      }]);
      n++;
    } catch (_) { /* 个别 cookie 无效则跳过 */ }
  }
  return n;
}
async function backupCookies(context) {
  let cookies = [];
  try { cookies = await context.cookies([XINGTU_ORIGIN, 'https://xingtu.cn']); } catch (_) { cookies = []; }
  const keep = cookies.filter(c => c && c.domain && /xingtu\.cn/.test(c.domain));
  if (keep.length) {
    try { fs.writeFileSync(COOKIES_FILE, JSON.stringify(keep)); } catch (_) {}
  }
  return keep.length;
}

// ---- 登录二维码截图（仅无头模式需要）---------------------------------------
async function captureLoginQr(page) {
  const deadline = Date.now() + 20000;
  while (Date.now() < deadline) {
    try {
      const info = await page.evaluate(() => {
        const pick = (el) => el ? { x: el.getBoundingClientRect().x, y: el.getBoundingClientRect().y, w: el.clientWidth, h: el.clientHeight } : null;
        const byText = () => {
          const els = Array.from(document.querySelectorAll('div,span,button,a,p'));
          const hit = els.find(e => /二维码|扫码登录|安全验证|切换.*二维码/.test(e.textContent || '') && (e.clientWidth < 320) && (e.clientHeight < 220));
          return hit ? pick(hit) : null;
        };
        const qrEl = document.querySelector('img[src*="qrcode"], [class*="qrcode"], [class*="qr-code"], canvas');
        const box = pick(qrEl) || byText();
        return box ? { box, url: location.href, title: document.title } : null;
      });
      if (info && info.box && info.box.w > 60 && info.box.h > 60) {
        const b = info.box;
        const buf = await page.screenshot({
          clip: {
            x: Math.max(0, b.x - 14),
            y: Math.max(0, b.y - 14),
            width: Math.min(1300, b.w + 28),
            height: Math.min(1300, b.h + 28),
          },
        }).catch(() => null);
        if (buf) return 'data:image/png;base64,' + buf.toString('base64');
      }
    } catch (_) { /* evaluate 失败则重试 */ }
    await sleep(900);
  }
  return null;
}

// ---- 登录态判定（纯 URL 判定，兼容纯服务商账号落在 /provider/）---------------
function isLoggedInUrl(u) {
  if (!u) return false;
  return /xingtu\.cn/.test(u) && /^\/(?:sup|provider)(?:\/|$)/.test(new URL(u, XINGTU_ORIGIN).pathname);
}

// 兼容 /sup/ 首页对纯服务商账号的重定向：/sup/ 不行就转服务商广场 /provider/
async function convergeLoggedInPage(page) {
  const safeUrl = () => { try { return page.url() || ''; } catch (_) { return ''; } };
  try {
    if (!/xingtu\.cn/.test(safeUrl())) await page.goto(SUP_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
    else { try { await page.bringToFront(); } catch (_) {} }
  } catch (_) { /* 导航异常交给轮询处理 */ }
  await sleep(2200);
  if (isLoggedInUrl(safeUrl())) return;
  // /sup/ 未建立会话（被重定向到营销首页等），改走 /provider/ 服务商广场
  try {
    await page.goto(PROVIDER_MARKET_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
    await sleep(2000);
  } catch (_) { /* 忽略 */ }
}

async function checkLoggedIn(page, mode) {
  try {
    const u = page.url();
    if (isLoggedInUrl(u)) {
      if (mode === 'headless') await backupCookies(page.context());
      return true;
    }
    // /sup/ 对纯服务商账号重定向营销首页时，转 /provider/ 广场再判
    if (/xingtu\.cn/.test(u)) await convergeLoggedInPage(page).catch(() => {});
    return isLoggedInUrl(page.url());
  } catch (_) { return false; }
}

async function closeAll(sess) {
  try { if (sess) await sess.browser.close(); } catch (_) {}
  try { if (sess && sess.proc) sess.proc.kill('SIGKILL'); } catch (_) {}
  if (_session === sess) _session = null;
}

/**
 * 统一登录流程（与 xc-tagger 同构）：
 * - CDP 模式：系统浏览器窗口可见，服务商在窗口里登录；
 * - 无头模式：截图登录二维码写 _qr，网页上手机抖音扫码。
 * 返回 {ok, loggedIn, cancelled?, needInstall?, mode?, browserName?, message?}。
 */
async function runLogin(signal) {
  let sess = null;
  try {
    sess = await acquireSession();
  } catch (e) {
    return { ok: false, loggedIn: false, needInstall: true, message: friendlyErr(e) };
  }

  let page = null;
  try {
    page = await sessionWorkPage(sess);
  } catch (e) {
    return { ok: false, loggedIn: false, needInstall: true, message: friendlyErr(e) };
  }

  // ① 先复核当前是否已登录（任一标签页）
  try {
    if (await checkLoggedIn(page, sess.mode)) {
      await backupCookies(sess.context).catch(() => {});
      return { ok: true, loggedIn: true, mode: sess.mode, browserName: sess.browserName };
    }
  } catch (_) {}

  // ② 无头模式：尝试截图登录二维码；CDP 模式无需二维码（服务商在真实窗口操作）
  if (sess.mode === 'headless') {
    try {
      const qr = await captureLoginQr(page);
      if (qr) { _qr = qr; _qrAt = Date.now(); }
    } catch (_) {}
  }

  // ③ 轮询等待登录（每轮遍历所有标签页，任一进入 /sup//provider/ 控制台即成功）
  const deadline = Date.now() + 5 * 60 * 1000;
  while (Date.now() < deadline) {
    if (signal && signal.aborted) return { ok: false, cancelled: true, message: '登录已取消' };
    try {
      if (!(await sessionAlive(sess))) {
        return { ok: false, loggedIn: false, message: '浏览器已关闭或连接中断，请重新登录。' };
      }
      const pages = sess.context.pages();
      let okPage = null;
      for (const p of pages) {
        try {
          if (p.isClosed()) continue;
          if (isLoggedInUrl(p.url())) { okPage = p; break; }
        } catch (_) {}
      }
      // 未发现控制台页时，兜底复核工作页（并触发 /provider/ 收敛）
      if (!okPage && page && !page.isClosed()) {
        if (await checkLoggedIn(page, sess.mode)) okPage = page;
      }
      if (okPage) {
        sess._workPage = okPage;
        try { await okPage.bringToFront(); } catch (_) {}
        const n = await backupCookies(sess.context);
        try { await okPage.goto(LOGIN_URL, { waitUntil: 'domcontentloaded', timeout: 20000 }); } catch (_) {}
        console.log(`✅ 登录成功！模式=${sess.mode}，已备份星图 Cookie ${n} 条。`);
        if (sess.mode === 'cdp') {
          // 登录完成，自动关闭调试浏览器窗口；后续如需重新登录可再次点击
          setTimeout(() => closeAll(sess), 1500);
        }
        return { ok: true, loggedIn: true, mode: sess.mode, browserName: sess.browserName };
      }
    } catch (_) { /* 轮询异常不中断 */ }
    await sleep(1500);
  }
  return { ok: false, loggedIn: false, message: '登录超时（5 分钟未完成扫码），请重新点击登录。' };
}

let _qr = null;
let _qrAt = 0;
let _loginRunning = null;
let _loginSignal = { aborted: false };

// ============================================================================
// 星图接口层：页内同源 fetch（自带登录 Cookie），逐行检索，不做打分评级
// ============================================================================

/**
 * 在页面上下文内执行同源 fetch（自动携带星图登录 Cookie，不暴露 Cookie）。
 * 返回 { ok, status, json, text }；HTTP/业务码异常均返回结构化结果。
 */
async function xgFetch(page, api, body, method = 'POST') {
  return await page.evaluate(async ({ origin, api, body, method }) => {
    try {
      const resp = await fetch(origin + api, {
        method,
        credentials: 'include',
        headers: { 'Content-Type': 'application/json', Accept: 'application/json, text/plain, */F' },
        body: method === 'GET' ? undefined : JSON.stringify(body || {}),
      });
      const text = await resp.text();
      let json = null;
      try { json = JSON.parse(text); } catch (_) { json = null; }
      return { ok: resp.ok, status: resp.status, json, text: text.slice(0, 500) };
    } catch (e) {
      return { ok: false, status: 0, json: null, text: String(e).slice(0, 300) };
    }
  }, { origin: XINGTU_ORIGIN, api, body, method });
}

// basic_search_authors 的场景信封（与 xc-tagger v4.1.1 相同；广场搜索必需）
const GSEARCH_SCENE = {
  scene_id: 4,
  scene_name: 'service_market',
  scene_identity: 4,
  page_identity: 3,
};

/**
 * 把 basic_search_authors 单行结果映射为身份三元组（纯函数）。
 * 结构：{star_id, attribute_datas:{nick_name/author_nick_name, unique_id/author_unique_id, short_id}}
 */
function mapGsearchRow(row) {
  if (!row || typeof row !== 'object') return null;
  const d = (row.attribute_datas && typeof row.attribute_datas === 'object') ? row.attribute_datas : {};
  const id = normText(row.star_id ?? row.star_author_id ?? row.author_id ?? row.id);
  const name = normText(d.nick_name ?? d.author_nick_name ?? row.nick_name ?? row.nickname);
  let douyin = normText(d.unique_id ?? d.author_unique_id ?? row.unique_id);
  if (!douyin) douyin = normText(d.short_id ?? row.short_id);
  if (!id || (!name && !douyin)) return null;
  return { id, name, douyin };
}

/**
 * 服务商广场搜索 basic_search_authors（seach_type：101=抖音号 / 0=名称）。
 * 完整信封：scene_param + search_param + page_param + sort_param；逐条调用。
 */
async function basicSearchAuthors(page, keyword, seachType, size = 20) {
  const kw = cleanHandle(keyword);
  if (!kw) return [];
  const body = {
    scene_param: GSEARCH_SCENE,
    search_param: { seach_type: seachType, keyword: kw },
    page_param: { page: 1, limit: size },
    sort_param: { sort_type: 2, sort_field: { field_name: 'score' } },
  };
  const r = await xgFetch(page, '/gw/api/gsearch/basic_search_authors', body);
  dumpDebug(`gsearch_${seachType}_${kw.slice(0, 12)}`, r.json || { status: r.status, text: r.text });
  if (!r.json) return [];
  const j = r.json;
  const rows = Array.isArray(j.authors) ? j.authors
    : (j.data && Array.isArray(j.data.authors)) ? j.data.authors
    : Array.isArray(j.data) ? j.data : [];
  return rows.map(mapGsearchRow).filter(Boolean);
}

/**
 * 搜索结果选取（核心纯逻辑，单测重点覆盖）：
 * @param {Array} mapped basic_search_authors 映射后的结果（排序顺序）
 * @param {{wantDouyin?:string, nameHint?:string}} opt
 * @returns {{via:string, id:string, name:string, douyin:string}|null}
 * 规则：
 *  - 有 wantDouyin：先找 unique_id 完全一致的精确结果；多个精确结果时优先名称一致，否则取排序第一。
 *  - 无精确结果（字母/短抖音号走模糊召回）：取排序第一，但若提供了名称提示且名称不一致则不采用。
 */
function pickSearchResult(mapped, opt = {}) {
  if (!Array.isArray(mapped) || !mapped.length) return null;
  const wantDouyin = cleanHandle(opt.wantDouyin).toLowerCase();
  const nameHint = cleanHandle(opt.nameHint);
  if (wantDouyin) {
    const exact = mapped.filter(m => cleanHandle(m.douyin).toLowerCase() === wantDouyin);
    if (exact.length === 1) {
      return { via: 'douyin_exact', ...exact[0] };
    }
    if (exact.length > 1) {
      const named = nameHint ? exact.filter(m => nameEqual(m.name, nameHint)) : [];
      return { via: 'douyin_exact', ...(named[0] || exact[0]) };
    }
    const first = mapped[0];
    if (!nameHint || nameEqual(first.name, nameHint)) {
      return { via: 'douyin_fuzzy', ...first };
    }
    return null;
  }
  // 名称兜底（seach_type=0）：有名称时第一条须名称一致；无名称时信任排序取第一条
  const first = mapped[0];
  if (!nameHint || nameEqual(first.name, nameHint)) {
    return { via: 'name_search', ...first };
  }
  return null;
}

// ② 抖音号检索（seach_type=101）
async function searchByDouyin(page, douyin, nameHint) {
  const mapped = await basicSearchAuthors(page, douyin, 101);
  return pickSearchResult(mapped, { wantDouyin: douyin, nameHint });
}

// ③ 达人名称检索（seach_type=0，昵称模糊搜索，取第一条 + 名称一致性过滤）
async function searchByName(page, name) {
  const mapped = await basicSearchAuthors(page, name, 0);
  return pickSearchResult(mapped, { nameHint: name });
}

// ============================================================================
// 星图ID 直达达人主页 /xingtu/author/{id}/：取 unique_id（抖音号）与昵称
// ============================================================================

/**
 * 递归在任意 JSON 对象里寻找达人身份（纯函数）。
 * 命中条件：对象同时含 抖音号键(unique_id/author_unique_id/...)与 昵称键(nick_name/...)；
 * 若对象还带 ID 键(star_id/author_id/...)，则必须与 expectedId 一致，避免误取同页其他达人。
 * @returns {Array<{id:string, douyin:string, name:string}>}
 */
function findIdentityCandidates(node, expectedId, out = [], depth = 0, seen = new Set()) {
  if (!node || typeof node !== 'object' || depth > 9) return out;
  if (seen.has(node)) return out;
  seen.add(node);
  if (Array.isArray(node)) {
    for (const n of node) findIdentityCandidates(n, expectedId, out, depth + 1, seen);
    return out;
  }
  const keys = Object.keys(node);
  let name = '';
  let handle = '';
  let idVal = '';
  for (const k of keys) {
    const v = node[k];
    if (v == null || typeof v === 'object') continue;
    const sval = String(v).trim();
    if (!sval) continue;
    if (!name && /^(nick_?name|author_nick_?name|nickname)$/i.test(k)) name = sval;
    if (!handle && /^(unique_?id|author_unique_?id|douyin_?unique_?id|douyin_?account|account_?name)$/i.test(k)) handle = sval;
    if (!idVal && /^(star_?id|author_?id|user_?id|uid|id)$/i.test(k) && /^\d+$/.test(sval)) idVal = sval;
  }
  // unique_id 缺失时（纯数字账号），short_id 兜底
  if (!handle) {
    for (const k of keys) {
      const v = node[k];
      if (/^short_?id$/i.test(k) && v != null && typeof v !== 'object' && String(v).trim()) {
        handle = String(v).trim();
        break;
      }
    }
  }
  if (name && handle) {
    if (!idVal || !expectedId || String(idVal) === String(expectedId)) {
      out.push({ id: idVal || expectedId || '', douyin: handle, name });
    }
  }
  for (const k of keys) {
    const v = node[k];
    if (v && typeof v === 'object') findIdentityCandidates(v, expectedId, out, depth + 1, seen);
  }
  return out;
}

function pickIdentity(cands, expectedId) {
  if (!cands || !cands.length) return null;
  if (expectedId) {
    const exact = cands.find(c => String(c.id) === String(expectedId));
    if (exact) return exact;
  }
  return cands[0];
}

/**
 * 从文本里用括号配平的方式提取所有可解析 JSON 块（纯函数）。
 * 覆盖 RENDER_DATA(URI编码) / __NEXT_DATA__ / window.__INITIAL_STATE__= 等内嵌形态。
 * 返回解析成功、且内容疑似达人数据（含 unique_id/nick_name/star_id 键）的对象数组。
 */
function extractJsonBlobs(text) {
  const out = [];
  if (!text || typeof text !== 'string') return out;

  // 1) <script id="RENDER_DATA" ...>URI编码JSON</script>
  const renderRe = /<script[^>]*\bid=["']RENDER_DATA["'][^>]*>([\s\S]*?)<\/script>/gi;
  let m;
  while ((m = renderRe.exec(text)) !== null) {
    const raw = m[1].trim();
    for (const candidate of [raw, safeDecodeURI(raw)]) {
      const j = tryParse(candidate);
      if (j) out.push(j);
    }
  }

  // 2) 所有 inline script：先直接解析，再做括号配平扫描
  const scriptRe = /<script\b[^>]*>([\s\S]*?)<\/script>/gi;
  while ((m = scriptRe.exec(text)) !== null) {
    const body = m[1];
    if (!body || body.length > 400000 || !/[{}]/.test(body)) continue;
    const direct = tryParse(body.trim()) || tryParse(afterFirstBrace(body));
    if (direct) { out.push(direct); continue; }
    // 括号配平：从每个 '{' 出发，找平衡串尝试解析（步长跳跃降低开销）
    for (let i = 0; i < body.length; i++) {
      if (body[i] !== '{') continue;
      const obj = tryBalancedParse(body, i);
      if (obj) {
        out.push(obj);
        // 找到疑似达人对象即可，避免全文扫描
        if (looksLikeCreatorJson(obj)) break;
      }
      // 跳过明显不是 JSON 的区域
      if (body[i + 1] && !/["{\[]/.test(body[i + 1])) {
        const next = body.indexOf('{', i + 1);
        if (next < 0) break;
        i = next - 1;
      }
    }
  }
  return out;
}

function safeDecodeURI(s) {
  try { return decodeURIComponent(s); } catch (_) { return ''; }
}
function tryParse(s) {
  if (!s || typeof s !== 'string') return null;
  const t = s.trim();
  if (!t || (t[0] !== '{' && t[0] !== '[')) return null;
  try { return JSON.parse(t); } catch (_) { return null; }
}
function afterFirstBrace(s) {
  const i = s.indexOf('{');
  return i >= 0 ? s.slice(i) : '';
}
// 从 start（'{'）开始按字符串/转义规则配平，尝试 JSON.parse
function tryBalancedParse(text, start) {
  let depth = 0;
  let inStr = false;
  let esc = false;
  for (let i = start; i < text.length; i++) {
    const ch = text[i];
    if (inStr) {
      if (esc) esc = false;
      else if (ch === '\\') esc = true;
      else if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') inStr = true;
    else if (ch === '{') depth++;
    else if (ch === '}') {
      depth--;
      if (depth === 0) {
        const slice = text.slice(start, i + 1);
        if (slice.length < 20) return null;
        try { return JSON.parse(slice); } catch (_) { return null; }
      }
    }
  }
  return null;
}
function looksLikeCreatorJson(obj) {
  try {
    const s = JSON.stringify(obj).slice(0, 4000);
    return /unique_id|nick_name|star_id/.test(s);
  } catch (_) { return false; }
}

/**
 * ① 星图ID → 达人主页检索：
 * 新开标签页加载 /xingtu/author/{id}/，同时收集页面 XHR/Fetch JSON 与内嵌 JSON，
 * 递归提取 unique_id（抖音号）与昵称。
 * @returns {Promise<{id:string, douyin:string, name:string}|null>}
 */
async function fetchByStarId(page, starId) {
  const id = keepDigits(starId);
  if (!id) return null;
  const context = page.context();
  const np = await context.newPage();
  const xhrObjs = [];
  np.on('response', async (resp) => {
    try {
      const u = resp.url();
      if (!/xingtu\.cn/.test(u)) return;
      const rt = resp.request().resourceType();
      if (rt !== 'xhr' && rt !== 'fetch') return;
      const j = await resp.json().catch(() => null);
      if (j) xhrObjs.push(j);
    } catch (_) { /* 单个响应异常忽略 */ }
  });

  const target = AUTHOR_URL(id);
  try {
    await np.goto(target, { waitUntil: 'domcontentloaded', timeout: 45000 });
  } catch (e) {
    debugLog('达人主页导航异常：' + friendlyErr(e));
  }
  // 等待前端渲染/XHR 返回
  await np.waitForTimeout(3500).catch(() => {});

  let finalPath = '';
  let html = '';
  let bodyText = '';
  try { finalPath = new URL(np.url()).pathname; } catch (_) {}
  try { html = await np.content(); } catch (_) { html = ''; }
  try { bodyText = await np.evaluate(() => (document.body && document.body.innerText || '').slice(0, 3000)); } catch (_) {}

  dumpDebug(`author_home_${id}`, { finalPath, bodyText: bodyText.slice(0, 600), xhrCount: xhrObjs.length });

  // 明显的不存在/未入驻信号
  const notFound = /达人不存在|该达人不存在|未入驻星图|达人已注销|页面不存在|404/.test(bodyText);

  let identity = null;
  const embedded = extractJsonBlobs(html);
  for (const root of [...xhrObjs, ...embedded]) {
    const cand = pickIdentity(findIdentityCandidates(root, id), id);
    if (cand && (cand.douyin || cand.name)) { identity = cand; break; }
  }

  await np.close().catch(() => {});

  if (!identity || notFound) return null;
  // 落点已离开目标页且未提取到有效身份（可能被重定向到登录/首页）
  if (!finalPath.startsWith(`/xingtu/author/${id}`) && !(identity.douyin && identity.name)) return null;
  return { id, douyin: identity.douyin || '', name: identity.name || '' };
}

// ============================================================================
// 名单解析 / 导出表构造 / 逐行三路互查
// ============================================================================

/**
 * 解析上传的 Excel/CSV 名单（纯逻辑，单测覆盖）。
 * 列可为【抖音号 / 星图ID / 达人名称】任意一列或多列：
 *  - 抖音号列：抖音号/抖音ID/抖音账号/douyin/unique_id 等；
 *  - 星图ID列：星图ID/达人ID/星图达人ID/star_id/author_id/id/uid 等；
 *  - 名称列：达人名称/达人名/昵称/名字/主播/name 等。
 * 返回 { items:[{row,id,douyin,name,orig}], headers, detected:{douyinKey,idKey,nameKey} }
 */
function parseListBuffer(buf, filename) {
  const wb = XLSX.read(buf, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: false });
  if (!rows.length) return { items: [], headers: [], detected: {} };
  const headers = Object.keys(rows[0]);

  // ① 抖音号列优先识别（避免被通用 ID 正则抢走，如表头「抖音ID」）
  const douyinKey = headers.find(h => /抖音号|抖音id|抖音账号|抖音|douyin|unique_?id/i.test(h));
  // ② 星图ID列
  const idKey = headers.find(h => h !== douyinKey && /星图id|星图达人id|达人id|star_?id|author_?id|\bid\b|uid|(?:^|[\s_])编号$/i.test(h));
  // ③ 名称列（跳过前两类；「星图达人ID」也含达人，不能误选）
  const nameKey = headers.find(h => h !== idKey && h !== douyinKey && /达人名称|达人名|昵称|名字|主播名|主播|账号名|name/i)
    || headers.find(h => h !== idKey && h !== douyinKey);

  const items = [];
  rows.forEach((r, i) => {
    let id = idKey ? normText(r[idKey]) : '';
    let douyin = douyinKey ? normText(r[douyinKey]) : '';
    let name = nameKey ? normText(r[nameKey]) : '';

    // 单列模板：该列按内容判别归属
    if (!douyinKey && !idKey && nameKey) {
      const v = normText(r[nameKey]);
      if (/^\d{12,}$/.test(cleanHandle(v))) id = keepDigits(v);
      else douyin = v;
      name = '';
    } else {
      // ID 单元格：保留其中数字（星图ID 为纯数字）
      if (id) id = keepDigits(id) || cleanHandle(id);
      // 无 ID 列时，抖音号列若填入了 12 位以上纯数字，按星图ID 处理
      if (!idKey && douyin && /^\d{12,}$/.test(cleanHandle(douyin))) {
        id = keepDigits(douyin);
        douyin = '';
      } else if (douyin) {
        douyin = cleanHandle(douyin);
      }
      if (name) name = cleanHandle(name);
    }

    if (!id && !douyin && !name) return;
    items.push({ row: i + 1, id, douyin, name, orig: r });
  });

  return { items, headers, count: items.length, detected: { douyinKey: douyinKey || '', idKey: idKey || '', nameKey: nameKey || '' } };
}

const VIA_LABEL = {
  id_home: '星图ID直达主页',
  douyin_exact: '抖音号精确匹配',
  douyin_fuzzy: '抖音号模糊匹配(排序第一)',
  name_search: '名称模糊搜索(排序第一)',
};

// 表头角色识别（与 parse 列检测同口径）
function classifyHeaders(headers) {
  const douyinH = headers.find(h => /抖音号|抖音id|抖音账号|抖音|douyin|unique_?id/i.test(h));
  const idH = headers.find(h => h !== douyinH && /星图id|星图达人id|达人id|star_?id|author_?id|\bid\b|uid|(?:^|[\s_])编号$/i.test(h));
  const nameH = headers.find(h => h !== idH && h !== douyinH && /达人名称|达人名|昵称|名字|主播名|主播|账号名|name/i.test(h));
  return { douyinH: douyinH || '', idH: idH || '', nameH: nameH || '' };
}

/**
 * 根据互查结果构造导出表（纯函数，单测覆盖）：
 *  - 原始列全部保留；原角色列为空时用补全值回填；
 *  - 原本不存在的列追加规范列（抖音号/星图ID/达人名称）；
 *  - 末两列固定为 状态 / 匹配方式。
 */
function buildExportRows(headers, results) {
  const hs = headers || [];
  const cls = classifyHeaders(hs);
  const finalHeaders = [...hs];
  if (!cls.douyinH) finalHeaders.push('抖音号');
  if (!cls.idH) finalHeaders.push('星图ID');
  if (!cls.nameH) finalHeaders.push('达人名称');
  finalHeaders.push('状态', '匹配方式');

  const isEmpty = (v) => v == null || String(v).trim() === '';

  const rows = (results || []).map((r) => {
    const o = {};
    for (const h of hs) o[h] = (r.orig && h in r.orig) ? r.orig[h] : '';
    const fill = (existingH, canonical, v) => {
      if (!v) { if (!existingH) o[canonical] = ''; return; }
      if (existingH) { if (isEmpty(o[existingH])) o[existingH] = v; }
      else o[canonical] = v;
    };
    fill(cls.douyinH, '抖音号', r.douyin);
    fill(cls.idH, '星图ID', r.id);
    fill(cls.nameH, '达人名称', r.name);
    o['状态'] = r.found ? '成功' : '未检索到';
    o['匹配方式'] = (r.via && VIA_LABEL[r.via]) || '';
    return o;
  });

  return { headers: finalHeaders, rows };
}

function resultOk(it, r, via) {
  return {
    row: it.row, orig: it.orig,
    input: { id: it.id || '', douyin: it.douyin || '', name: it.name || '' },
    id: r.id || '', douyin: r.douyin || '', name: r.name || '',
    found: true, via, note: '',
  };
}
function resultFail(it, inId, inDy, inNm, note) {
  return {
    row: it.row, orig: it.orig,
    input: { id: inId, douyin: inDy, name: inNm },
    id: inId, douyin: inDy, name: inNm,
    found: false, via: '', note,
  };
}

// 默认检索器（单测可注入桩实现；byId 参数顺序为 (id,page) 与 resolveItem 调用一致）
const DEFAULT_FETCHERS = {
  byId: (id, page) => fetchByStarId(page, id),
  byDouyin: (page, handle, hint) => searchByDouyin(page, handle, hint),
  byName: (page, name) => searchByName(page, name),
};

/**
 * 单行三路互查（星图ID > 抖音号 > 名称），带缓存避免重复输入重复请求。
 * @param {Map} caches {id:Map, dy:Map, nm:Map}
 */
async function resolveItem(page, it, caches, deps) {
  const F = deps || DEFAULT_FETCHERS;
  const inId = it.id ? (keepDigits(it.id) || cleanHandle(it.id)) : '';
  const inDy = it.douyin ? cleanHandle(it.douyin) : '';
  const inNm = it.name ? cleanHandle(it.name) : '';

  if (!inId && !inDy && !inNm) return resultFail(it, '', '', '', '三列均为空，无法检索');

  // ① 星图ID → 达人主页
  if (inId) {
    if (!caches.id.has(inId)) {
      let r = null;
      try { r = await F.byId(inId, page); } catch (_) { r = null; }
      caches.id.set(inId, r);
      await sleep(650 + Math.floor(Math.random() * 350));
    }
    const r = caches.id.get(inId);
    if (r && (r.douyin || r.name)) {
      return resultOk(it, { id: r.id || inId, douyin: r.douyin, name: r.name }, 'id_home');
    }
  }

  // ② 抖音号精确检索（seach_type=101）
  if (inDy) {
    const key = inDy.toLowerCase() + '|' + (inNm ? inNm.toLowerCase() : '');
    if (!caches.dy.has(key)) {
      let r = null;
      try { r = await F.byDouyin(page, inDy, inNm); } catch (_) { r = null; }
      caches.dy.set(key, r);
      await sleep(550 + Math.floor(Math.random() * 350));
    }
    const r = caches.dy.get(key);
    if (r && r.id) {
      return resultOk(it, { id: r.id, douyin: r.douyin || inDy, name: r.name }, r.via || 'douyin_exact');
    }
  }

  // ③ 名称兜底检索（seach_type=0）
  if (inNm) {
    const key = inNm.toLowerCase();
    if (!caches.nm.has(key)) {
      let r = null;
      try { r = await F.byName(page, inNm); } catch (_) { r = null; }
      caches.nm.set(key, r);
      await sleep(550 + Math.floor(Math.random() * 350));
    }
    const r = caches.nm.get(key);
    if (r && r.id) {
      return resultOk(it, { id: r.id, douyin: r.douyin, name: r.name || inNm }, 'name_search');
    }
  }

  return resultFail(it, inId, inDy, inNm, '三路检索均未命中（可能未入驻星图、抖音号/名称有误或检索结果名称不一致）');
}

// ============================================================================
// HTTP 服务
// ============================================================================

const app = express();
app.use(cors({
  origin(origin, cb) {
    // 同源（页面由本服务提供，浏览器不带 Origin）直接放行
    if (!origin) return cb(null, true);
    if (ALLOWED_ORIGINS.includes(origin)) return cb(null, true);
    try {
      const u = new URL(origin);
      if ((u.hostname === 'localhost' || u.hostname === '127.0.0.1') && /^\d+$/.test(u.port)) return cb(null, true);
      if (u.hostname.endsWith('github.io') && u.hostname.startsWith('didimarco26')) return cb(null, true);
    } catch (_) { /* origin 非 URL，拒绝 */ }
    return cb(new Error('Origin not allowed by CORS: ' + origin));
  },
}));
app.use(express.json({ limit: '15mb' }));

// 静态页面（工具自带网页界面）
app.use(express.static(PUBLIC_DIR));
// 前端导出 Excel 用的 SheetJS
app.get('/vendor/xlsx.full.min.js', (req, res) => {
  res.sendFile(path.join(__dirname, 'node_modules', 'xlsx', 'dist', 'xlsx.full.min.js'));
});

// 互查任务状态
let lookupProgress = { running: false, total: 0, done: 0, current: null, phase: '' };
let lookupResults = [];

// GET /health — 登录态 + 任务状态
app.get('/health', async (req, res) => {
  let loggedIn = false;
  let browser = null;
  try {
    if (_session && await sessionAlive(_session)) {
      browser = { name: _session.browserName, mode: _session.mode };
      const p = await sessionWorkPage(_session);
      loggedIn = await checkLoggedIn(p, _session.mode);
    } else if (readCookieFile().length > 0) {
      browser = { name: 'Cookie 备份（未启动浏览器，将在登录/互查时自动恢复）', mode: 'cookie' };
    }
  } catch (_) { /* health 永不抛错 */ }
  res.json({
    ok: true,
    version: VERSION,
    loggedIn,
    loginPending: !!(_loginRunning),
    browser,
    cookieCount: readCookieFile().length,
    lookup: { ...lookupProgress },
    loginUrl: LOGIN_URL,
  });
});

// POST /login — 启动统一登录流程（5 分钟超时）；同次只跑一个
app.post('/login', async (req, res) => {
  if (_loginRunning) {
    try { const r = await _loginRunning; return res.json(r); }
    catch (e) { return res.json({ ok: false, message: friendlyErr(e) }); }
  }
  _loginSignal = { aborted: false };
  _loginRunning = runLogin(_loginSignal);
  _loginRunning.finally(() => { _loginRunning = null; });
  try {
    const result = await _loginRunning;
    res.json(result);
  } catch (e) {
    res.json({ ok: false, loggedIn: false, message: friendlyErr(e) });
  }
});

// GET /login/qr — 无头模式登录二维码（10 分钟有效）
app.get('/login/qr', (req, res) => {
  if (_qr && Date.now() - _qrAt < 10 * 60 * 1000) {
    res.setHeader('Cache-Control', 'no-store').json({ qr: _qr });
  } else {
    res.json({ qr: null, message: '请先点击「登录星图」生成二维码。' });
  }
});

// raw body（文件上传）必须在 /parse 路由之前挂载
app.use('/parse', (req, res, next) => { express.raw({ type: '*/*', limit: '15mb' })(req, res, next); });

// POST /parse — 解析名单（multipart 原始 body，文件名走 query）
app.post('/parse', (req, res) => {
  try {
    if (!Buffer.isBuffer(req.body) || req.body.length === 0) {
      return res.status(400).json({ ok: false, error: '未收到文件内容。' });
    }
    const fn = req.query.fn ? String(req.query.fn) : 'list.xlsx';
    const lower = fn.toLowerCase();
    if (!/\.(xlsx|xls|csv)$/.test(lower)) {
      return res.status(400).json({ ok: false, error: '仅支持 .xlsx / .xls / .csv 文件。' });
    }
    const { items, headers, detected } = parseListBuffer(req.body, fn);
    if (!items.length) {
      return res.status(400).json({ ok: false, error: '未从文件中解析到有效行（请确认首个工作表含达人信息）。' });
    }
    res.json({ ok: true, count: items.length, headers, detected, preview: items.slice(0, 5) });
  } catch (e) {
    res.status(500).json({ ok: false, error: '解析失败：' + friendlyErr(e) });
  }
});

// POST /lookup — 逐行三路互查（异步执行，前端轮询 /progress）
app.post('/lookup', async (req, res) => {
  if (lookupProgress.running) {
    return res.status(409).json({ ok: false, error: '已有互查任务在执行中。' });
  }
  const items = Array.isArray(req.body && req.body.items) ? req.body.items : [];
  if (!items.length) return res.status(400).json({ ok: false, error: '名单为空。' });

  let sess = null;
  try { sess = await acquireSession(); }
  catch (e) { return res.status(500).json({ ok: false, error: friendlyErr(e), needInstall: true }); }

  let page = null;
  try { page = await sessionWorkPage(sess); }
  catch (e) { return res.status(500).json({ ok: false, error: friendlyErr(e) }); }

  if (!(await checkLoggedIn(page, sess.mode))) {
    return res.status(401).json({ ok: false, error: '星图登录态已失效，请重新点击「登录星图」后再开始。' });
  }

  lookupResults = [];
  lookupProgress = { running: true, total: items.length, done: 0, current: null, phase: '' };
  res.json({ ok: true, total: items.length });

  const caches = { id: new Map(), dy: new Map(), nm: new Map() };
  (async () => {
    const phaseOf = (it) => {
      if (it.id) return '星图ID主页检索';
      if (it.douyin) return '抖音号精确检索';
      return '名称模糊检索';
    };
    for (const it of items) {
      lookupProgress.current = it.row;
      lookupProgress.phase = phaseOf(it);
      try {
        const r = await resolveItem(page, it, caches);
        lookupResults.push(r);
      } catch (e) {
        lookupResults.push(resultFail(it, it.id || '', it.douyin || '', it.name || '', '检索异常：' + friendlyErr(e)));
      }
      lookupProgress.done++;
    }
    lookupProgress.running = false;
    lookupProgress.current = null;
    lookupProgress.phase = '完成';
    const okN = lookupResults.filter(r => r.found).length;
    console.log(`✅ 互查完成：${lookupResults.length} 行，成功 ${okN}，未检索到 ${lookupResults.length - okN}。`);
  })();
});

// GET /progress — 进度 + since 游标之后的增量结果
app.get('/progress', (req, res) => {
  const since = parseInt(req.query.since, 10);
  const payload = { ...lookupProgress };
  if (!Number.isFinite(since) || since < 0) {
    res.json({ ...payload, results: lookupResults });
  } else {
    res.json({ ...payload, results: lookupResults.slice(since) });
  }
});

// ---- 启动 ----------------------------------------------------------------
function openBrowser(url) {
  try {
    const { spawn } = require('child_process');
    if (process.platform === 'darwin') {
      spawn('open', [url], { stdio: 'ignore' }).unref();
    } else if (process.platform === 'win32') {
      spawn('cmd', ['/c', 'start', '', url], { stdio: 'ignore' }).unref();
    } else {
      spawn('xdg-open', [url], { stdio: 'ignore' }).unref();
    }
  } catch (_) { /* 自动打开失败时，用户可手动访问地址 */ }
}

// 仅在直接运行本文件时启动服务（被 test-logic require 时不启动）
let server = null;
if (require.main === module) {
  server = app.listen(PORT, HOST, () => {
    const url = `http://${HOST}:${PORT}`;
    console.log('');
    console.log('============================================================');
    console.log(`  xc-id-lookup 达人信息互查 v${VERSION} 已启动`);
    console.log(`  本地服务地址：${url}（仅本机可访问）`);
    console.log('  保持本窗口开启；关闭窗口（或 Ctrl+C）即停止工具。');
    console.log('============================================================');
    console.log('');
    // 自动打开工具网页
    openBrowser(url);
  });
}

process.on('SIGINT', async () => {
  console.log('\n正在停止…');
  try { if (_session) await closeAll(_session); } catch (_) {}
  try { if (server) server.close(); } catch (_) {}
  process.exit(0);
});
process.on('SIGTERM', () => process.exit(0));

// 导出纯函数供单测引用
module.exports = {
  VERSION,
  browserCandidates,
  connectExistingCdp,
  cleanHandle,
  nameEqual,
  keepDigits,
  mapGsearchRow,
  pickSearchResult,
  findIdentityCandidates,
  pickIdentity,
  extractJsonBlobs,
  parseListBuffer,
  classifyHeaders,
  buildExportRows,
  resolveItem,
};
