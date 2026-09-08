#!/usr/bin/env node
/* eslint-disable */
/**
 * 星川服务商达人自助打标 · 本地一键工具 (xc-tagger) v3
 * ------------------------------------------------------------------
 * 用途：服务商在本机运行本工具，它会：
 *   1) 在 127.0.0.1:7842 起一个本地 HTTP 服务（只监听本机，不对外）；
 *   2) 登录（仅需一次）：网页点「🚀 登录星图」→ POST /login 启动一个【有界面
 *      headed Chromium 窗口】打开巨量星图达人广场，服务商在窗口里扫码/验证码
 *      登录；工具每 2 秒检测一次，检测到登录成功后把 Cookie 写入本机
 *      .xc-cookies.json 并自动关闭窗口（超时 180 秒；窗口被手动关闭/网页点
 *      「取消」都会中止本次登录）。
 *   3) 打标：POST /tag 仍以【无头 headless 模式】后台运行 Chromium（不弹窗），
 *      启动时自动加载 .xc-cookies.json；检测到有效会话时自动回写刷新 Cookie。
 *      headed 与 headless 是两个独立浏览器实例，登录态靠 .xc-cookies.json 传递。
 *   4) 网页（星川决策工作台·服务商 GitHub 版）调用本地接口：
 *        GET  /health   → 探测服务与登录状态（含 loginPending 扫码等待标志）
 *        POST /login    → 弹出 Chrome 窗口扫码登录，成功后自动关窗
 *        POST /parse    → 解析上传的 Excel/CSV 达人名单，提取 达人ID / 昵称
 *        POST /tag      → 用星图登录态抓取达人信息并按存量逻辑分层打标
 *   Cookie 只留在服务商本机（.xc-cookies.json），不上传、不落库。
 *
 * 运行：
 *   1. npm install        （首次，会自动安装 Playwright Chromium）
 *   2. node xc-tagger.js  （后台无头运行；在网页里点「登录星图」弹窗扫码即可）
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

// ---- 浏览器（Playwright 持久化登录态）-------------------------------------
let _browser = null;
let _page = null;
let _launching = null;

/**
 * 获取 Playwright 自己管理的 Chromium 可执行路径。
 * playwright.chromium.executablePath() 在部分 Windows 机器上会返回系统 Edge 的路径
 * （Edge 被自动化控制后会主动退出，报 "Target page, context or browser has been closed"）。
 * 此函数尝试多种方式拿到真正的 Playwright Chromium，并做路径合法性校验。
 */
function getPlaywrightChromiumPath() {
  // 方式1：通过 Playwright 内部 registry 查询真实下载路径（1.x 可用）
  try {
    const { registry } = require('playwright/lib/server/registry');
    const executable = registry.findExecutable('chromium');
    if (executable) {
      const p = executable.executablePath('linux') || executable.executablePath();
      if (p && fs.existsSync(p) && !/edge|microsoft/i.test(p)) return p;
    }
  } catch (_) { /* 内部 API 不可用时忽略 */ }

  // 方式2：通过环境变量 PLAYWRIGHT_BROWSERS_PATH 或默认安装位置手动查找
  try {
    const os = require('os');
    const candidateBases = [
      process.env.PLAYWRIGHT_BROWSERS_PATH,
      path.join(os.homedir(), '.cache', 'ms-playwright'),
      path.join(os.homedir(), 'AppData', 'Local', 'ms-playwright'),
      path.join(os.homedir(), 'Library', 'Caches', 'ms-playwright'),
      path.join(os.homedir(), '.cache', 'playwright'),
    ].filter(Boolean);

    for (const base of candidateBases) {
      if (!fs.existsSync(base)) continue;
      const entries = fs.readdirSync(base).filter(d => d.startsWith('chromium-'));
      for (const entry of entries.sort().reverse()) { // 优先最新版本
        const candidates = [
          path.join(base, entry, 'chrome-win', 'chrome.exe'),
          path.join(base, entry, 'chrome-linux', 'chrome'),
          path.join(base, entry, 'chrome-mac', 'Chromium.app', 'Contents', 'MacOS', 'Chromium'),
        ];
        for (const p of candidates) {
          if (fs.existsSync(p) && !/edge|microsoft/i.test(p)) return p;
        }
      }
    }
  } catch (_) { /* 查找失败忽略 */ }

  // 方式3：兜底用 executablePath()，但检查是否是 Edge（含 edge/microsoft 关键字则报错）
  const playwright = require('playwright');
  const exePath = playwright.chromium.executablePath();
  if (/edge|microsoft/i.test(exePath)) {
    throw new Error(
      'Playwright 指向了系统 Edge 而非 Playwright 自带 Chromium。\n' +
      '路径：' + exePath + '\n' +
      '请在工具目录的命令行执行以下命令后重启工具：\n' +
      '  npx playwright install chromium\n' +
      '若问题仍存在，请设置环境变量后重试：\n' +
      '  set PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD=0\n' +
      '  npx playwright install chromium'
    );
  }
  if (!fs.existsSync(exePath)) {
    throw new Error('Playwright 自带 Chromium 未安装（' + exePath + ' 不存在）。\n' +
      '请在工具目录运行：npx playwright install chromium');
  }
  return exePath;
}

// ---- Cookie 登录态（无头打标 + 有界面登录）----------------------------------
function readCookieFile() {
  try {
    if (!fs.existsSync(COOKIES_FILE)) return [];
    const arr = JSON.parse(fs.readFileSync(COOKIES_FILE, 'utf8'));
    return Array.isArray(arr) ? arr : [];
  } catch (_) { return []; }
}

// 逐条注入 Cookie，单条失败不影响其余，返回成功条数
async function addCookiesSafe(ctx, cookies) {
  let ok = 0;
  for (const ck of cookies) {
    try { await ctx.addCookies([ck]); ok++; } catch (_) { /* 跳过非法 / 过期项 */ }
  }
  return ok;
}

async function ensureBrowser() {
  if (_browser && _browser.isConnected && _browser.isConnected() && _page && !_page.isClosed()) return _page;
  if (_launching) return _launching;
  _launching = (async () => {
    const exePath = getPlaywrightChromiumPath();
    console.log('✅ 使用 Chromium（无头后台运行，不弹窗）：' + exePath);
    const playwright = require('playwright');
    // 用普通 launch（非持久化）+ headless:true：Windows 安全策略下 headless:false 弹窗的
    // Chromium 会被直接关闭（exitCode=0，报 "Target page, context or browser has been closed"）；
    // 无头模式不弹窗、稳定运行，登录态靠 Cookie 文件注入。
    _browser = await playwright.chromium.launch({
      executablePath: exePath, // 锁定 Playwright 自带 Chromium，不用系统 Edge/Chrome
      headless: true,
      args: ['--disable-blink-features=AutomationControlled', '--no-sandbox'],
    });
    const ctx = await _browser.newContext({
      userAgent: UA,
      viewport: { width: 1280, height: 860 },
    });
    // 注入已保存的星图 Cookie
    const saved = readCookieFile();
    if (saved.length) {
      const n = await addCookiesSafe(ctx, saved);
      if (n > 0) console.log('✅ 已加载星图登录 Cookie（' + n + ' 条）');
    }
    _page = await ctx.newPage();
    try {
      await _page.goto(SQUARE_URL, { waitUntil: 'domcontentloaded', timeout: 30000 });
      await _page.waitForTimeout(2500); // 等页面 JS 渲染，避免误判登录态
    } catch (_) { /* 网络波动忽略 */ }
    _browser.on('disconnected', () => { _browser = null; _page = null; _launching = null; });
    return _page;
  })();
  try { return await _launching; } finally { _launching = null; }
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

// ---- 有界面登录（headed Chromium 弹窗扫码）---------------------------------
// /login 正在等待扫码时为 true，/health 据此让前端显示「等待扫码中…」。
let _loginPending = false;
const LOGIN_WAIT_MS = 180000; // 扫码最长等待 3 分钟

/**
 * 启动一个【有界面】Chromium 窗口打开星图达人广场，轮询检测登录态。
 * 成功：Cookie 落盘 .xc-cookies.json，关窗返回 {ok:true}；
 * 失败：超时 / 窗口被关 / 前端取消，关窗返回 {ok:false,message}。
 * @param {{aborted:boolean}} signal 前端取消标志（请求中断时置 true）
 */
async function headedLogin(signal) {
  const playwright = require('playwright');
  const exePath = getPlaywrightChromiumPath();
  console.log('🟢 启动有界面 Chromium 供扫码登录：' + exePath);
  const browser = await playwright.chromium.launch({
    executablePath: exePath,
    headless: false, // 有界面，供服务商扫码
    args: ['--disable-blink-features=AutomationControlled'],
  });
  let disconnected = false;
  browser.on('disconnected', () => { disconnected = true; });
  try {
    const ctx = await browser.newContext({
      userAgent: UA,
      viewport: { width: 1280, height: 800 },
    });
    const page = await ctx.newPage();
    try {
      await page.goto(SQUARE_URL, { waitUntil: 'domcontentloaded', timeout: 45000 });
    } catch (e) {
      console.log('⚠️ 打开星图页面异常（仍可在窗口内手动刷新）：' + friendlyErr(e));
    }
    const deadline = Date.now() + LOGIN_WAIT_MS;
    while (Date.now() < deadline) {
      if (disconnected) return { ok: false, message: '登录窗口被关闭，请重新点「登录星图」再试' };
      if (signal && signal.aborted) return { ok: false, message: '已取消登录' };
      let logged = false;
      try { logged = await checkLoggedIn(page); } catch (_) { /* 页面跳转中，下轮再测 */ }
      if (logged) {
        // checkLoggedIn 成功时已回写一次 Cookie，这里再显式落盘兜底
        try {
          const cookies = await ctx.cookies();
          if (Array.isArray(cookies) && cookies.length) {
            fs.writeFileSync(COOKIES_FILE, JSON.stringify(cookies, null, 2), 'utf8');
          }
        } catch (_) { /* 忽略 */ }
        console.log('✅ 扫码登录成功，Cookie 已保存到 .xc-cookies.json');
        return { ok: true, loggedIn: true, message: '登录成功，登录窗口即将自动关闭' };
      }
      await sleep(2000);
    }
    return { ok: false, loggedIn: false, message: '登录超时（3 分钟内未检测到登录成功），请重试' };
  } finally {
    try { await browser.close(); } catch (_) { /* 已关闭则忽略 */ }
  }
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

  // 健康检查 / 登录状态
  app.get('/health', async (_req, res) => {
    let loggedIn = false;
    try {
      if (_page) loggedIn = await checkLoggedIn(_page);
      else if (!_loginPending) ensureBrowser().catch(() => {}); // 尚未启动则后台预热（扫码等待中不抢占）
    } catch (_) {}
    const cookieCount = readCookieFile().filter(c => /xingtu/i.test(c.domain || '')).length;
    res.json({
      ok: true, loggedIn, version: '3.0.0', port: PORT,
      loginUrl: SQUARE_URL,
      cookiesFile: '.xc-cookies.json',
      cookieCount,
      loginPending: _loginPending,
      loginNote: loggedIn ? '' :
        '请点工作台「🚀 登录星图」按钮，本机会弹出 Chrome 窗口，在窗口里完成星图扫码登录，' +
        '窗口会自动关闭（登录态保存在本机，仅需一次）。',
    });
  });

  // 弹窗扫码登录：POST /login
  // - 已有有效登录态 → 直接返回；
  // - 否则启动【有界面】Chromium 窗口，每 2 秒检测一次，最多等 180 秒；
  // - 前端取消（请求中断）/ 窗口被手动关闭 / 超时，都会中止并返回提示。
  app.post('/login', async (req, res) => {
    if (_loginPending) {
      return res.status(409).json({ ok: false, pending: true, error: '已有登录窗口在等待扫码，请在弹出的 Chrome 窗口里完成登录' });
    }
    let aborted = false;
    req.on('close', () => { aborted = true; }); // 前端点「取消」或离开页面
    try {
      // 1) 先用现有 Cookie 快速验证一次，已登录就不弹窗
      try {
        const page = await ensureBrowser();
        if (await checkLoggedIn(page)) {
          return res.json({ ok: true, loggedIn: true, message: '已登录，无需重复操作' });
        }
      } catch (_) { /* 无头浏览器异常也继续走弹窗登录 */ }
      // 2) 弹出有界面窗口等扫码
      _loginPending = true;
      const r = await headedLogin({ get aborted() { return aborted; } });
      if (r.ok) {
        // 登录成功：重置无头浏览器，下次打标用新 Cookie 启动
        if (_browser) { try { await _browser.close(); } catch (_) {} }
        _browser = null; _page = null; _launching = null;
      }
      res.json({
        ok: !!r.ok, loggedIn: !!r.loggedIn, pending: false,
        message: r.message || (r.ok ? '登录成功' : '登录失败，请重试'),
      });
    } catch (e) {
      const raw = friendlyErr(e).split('Browser logs')[0].trim();
      res.status(500).json({
        ok: false,
        error: '登录窗口启动失败：' + raw +
          '。请确认电脑有图形桌面环境（不能在远程服务器/无桌面环境运行）；' +
          '若浏览器被安全软件拦截，请放行后重试。',
      });
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
      const page = await ensureBrowser();
      const loggedIn = await checkLoggedIn(page);
      if (!loggedIn) {
        return res.status(401).json({ ok: false, loggedIn: false, error: '星图未登录：请点工作台「🚀 登录星图」按钮，在弹出的 Chrome 窗口里扫码登录后再开始打标' });
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
    console.log('  星川服务商达人自助打标 · 本地工具已启动 v3.0.0');
    console.log(`  本地服务：http://${HOST}:${PORT}`);
    console.log('  工作台网页：https://didimarco26.github.io/xingchuan-workbench/');
    console.log('--------------------------------------------------');
    console.log('  使用 3 步：');
    console.log('   1) 保持本窗口打开（打标时浏览器在后台无头运行，不弹窗）；');
    console.log('   2) 在工作台网页点「🚀 登录星图」，本机弹出 Chrome 窗口，');
    console.log('      扫码登录后窗口自动关闭（登录态保存在本机，仅需一次）；');
    console.log('   3) 上传达人名单 Excel，点「开始打标」等待结果。');
    console.log('==================================================\n');
    ensureBrowser().then(async (page) => {
      const in_ = await checkLoggedIn(page);
      console.log(in_ ? '✅ 检测到星图登录态有效（.xc-cookies.json），可直接使用。\n'
        : '⚠️  尚未登录：请到工作台网页点「🚀 登录星图」扫码（弹窗约 3 分钟内有效）。\n');
    }).catch(e => console.log('⚠️  后台浏览器启动失败：' + friendlyErr(e) + '\n   登录时会自动重试；若提示缺少浏览器，请运行：npx playwright install chromium\n'));
  });
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
function friendlyErr(e) { return (e && (e.message || String(e))) || '未知错误'; }

startServer();
