#!/usr/bin/env node
/* eslint-disable */
/**
 * 星川服务商达人自助打标 · 本地一键工具 (xc-tagger)
 * ------------------------------------------------------------------
 * 用途：服务商在本机运行本工具，它会：
 *   1) 在 127.0.0.1:7842 起一个本地 HTTP 服务（只监听本机，不对外）；
 *   2) 用 Playwright 打开持久化 Chrome（登录态保存在 .xc-chrome-profile/），
 *      首次运行自动打开「巨量星图 · 达人广场」，扫码登录一次即可长期复用；
 *   3) 网页（星川决策工作台·服务商 GitHub 版）调用本地接口：
 *        GET  /health  → 探测服务与登录状态
 *        POST /parse   → 解析上传的 Excel/CSV 达人名单，提取 达人ID / 昵称
 *        POST /tag     → 用星图登录态抓取达人信息并按存量逻辑分层打标
 *   Cookie / 登录态只留在服务商本机（.xc-chrome-profile/），不上传、不落库。
 *
 * 运行：
 *   1. npm install        （首次，会自动安装 Playwright Chromium）
 *   2. node xc-tagger.js  （启动后按提示在弹出的 Chrome 里扫码登录星图）
 *
 * 安全：CORS 仅放行 didimarco26.github.io 与本机页面；服务只绑定 127.0.0.1。
 */

'use strict';

const path = require('path');

// 启动防御（必须早于任何外部依赖 require）：工具所在路径含空格/括号时，
// Chromium --user-data-dir 会解析失败、浏览器启动即退出（has been closed），登录态无法保存。
// 正常双击启动脚本会先自动迁移到干净路径；走到这里说明是在坏路径里直接命令行运行的。
if (/[ ()]/.test(__dirname)) {
  console.error('\n==================================================');
  console.error('[ERROR] 工具所在文件夹路径含空格或括号：');
  console.error('  ' + __dirname);
  console.error('该路径下 Chromium 无法启动，星图登录态也无法保存。');
  console.error('请双击 start-xc-tagger.command（Mac）/ start-xc-tagger.bat（Windows）启动，');
  console.error('启动脚本会自动把工具复制到干净路径（如 ~/xc-tagger 或 C:\\xc-tagger）；');
  console.error('或手动把整个 xc-tagger 文件夹移动到不含空格/括号的路径后重试。');
  console.error('==================================================\n');
  process.exit(2);
}

const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const fs = require('fs');

// ---- 配置 ----------------------------------------------------------------
const PORT = 7842;
const HOST = '127.0.0.1';
// Profile 目录：必须是绝对路径。含空格/括号的路径会让 Chromium --user-data-dir 解析失败、
// 浏览器启动即退出（launchPersistentContext: Target page, context or browser has been closed）。
const PROFILE_DIR = path.resolve(__dirname, '.xc-chrome-profile');
// 巨量星图 · 达人广场（广告主侧）。抓取接口与该页同源(www.xingtu.cn)，天然带 Cookie。
const XINGTU_ORIGIN = 'https://www.xingtu.cn';
const SQUARE_URL = 'https://www.xingtu.cn/ad/creator/square';
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

async function ensureBrowser() {
  if (_browser && _page && _browser.isConnected && _browser.isConnected()) return _page;
  if (_launching) return _launching;
  _launching = (async () => {
    const playwright = require('playwright');
    // 强制使用 Playwright 自己下载的 Chromium：显式传 executablePath，
    // 杜绝任何情况下回退到系统 Edge/Chrome（系统 Edge 被自动化控制时会主动退出，
    // 报 "Target page, context or browser has been closed"）。
    const exePath = playwright.chromium.executablePath();
    if (!fs.existsSync(exePath)) {
      throw new Error('Playwright 自带 Chromium 未安装（' + exePath + ' 不存在）。\n' +
        '   请双击启动脚本（会自动安装），或在工具目录执行：npx playwright install chromium');
    }
    fs.mkdirSync(PROFILE_DIR, { recursive: true });
    _browser = await playwright.chromium.launchPersistentContext(PROFILE_DIR, {
      executablePath: exePath, // 锁定 Playwright Chromium，不用系统浏览器
      headless: false, // 需要可见窗口以便扫码登录
      viewport: { width: 1280, height: 860 },
      args: ['--disable-blink-features=AutomationControlled', '--start-maximized'],
    });
    let page = _browser.pages && _browser.pages()[0];
    if (!page) page = await _browser.newPage();
    _page = page;
    try { await page.goto(SQUARE_URL, { waitUntil: 'domcontentloaded', timeout: 60000 }); } catch (_) { /* 登录跳转/网络波动忽略 */ }
    _browser.on('close', () => { _browser = null; _page = null; _launching = null; });
    return _page;
  })();
  try { return await _launching; } finally { _launching = null; }
}

// 是否已登录：URL 含 redirect_uri（被踢去登录）或页面出现登录按钮，则视为未登录
async function checkLoggedIn(page) {
  try {
    const url = page.url() || '';
    if (/redirect_uri|login|passport/.test(url)) return false;
    return await page.evaluate(() => {
      const t = document.body ? document.body.innerText : '';
      // 出现明显「扫码登录 / 登录」入口且无达人广场特征，判为未登录
      const hasLoginBtn = /扫码登录|登录巨量星图|手机号登录|验证码登录/.test(t) && !/达人广场|找达人|达人榜单/.test(t);
      return !hasLoginBtn;
    });
  } catch { return false; }
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
  // 达人人设标签（如：测评/剧情/种草/知识…）
  const persona = normText(pick(raw, [/persona|人设|标签|tag|label/i])).split(/[、,，/|]/).map(s => s.trim()).filter(Boolean).slice(0, 5);
  // 内容形式标签（短视频/直播/图文…）
  const forms = normText(pick(raw, [/content_?type|内容_?形式|形式|video_?type|material/i])).split(/[、,，/|]/).map(s => s.trim()).filter(Boolean).slice(0, 5);
  // 主要带货类目 / 行业
  const category = normText(pick(raw, [/category|cate|类目|行业|industry|vertical/i])).slice(0, 60);
  return { id, name, fans, fansTier, sLevel, lLevel, deliveries, consumption, persona, forms, category, raw };
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
    try { if (_page) loggedIn = await checkLoggedIn(_page); } catch (_) {}
    res.json({ ok: true, loggedIn, version: '1.0.0', port: PORT });
  });

  // 主动唤起浏览器登录
  app.post('/login', async (_req, res) => {
    try {
      const page = await ensureBrowser();
      await page.goto(SQUARE_URL, { waitUntil: 'domcontentloaded', timeout: 60000 }).catch(() => {});
      res.json({ ok: true, loggedIn: await checkLoggedIn(page), message: '已打开星图达人广场，请在弹出的 Chrome 窗口中扫码/登录' });
    } catch (e) {
      res.status(500).json({ ok: false, error: friendlyErr(e) });
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
        return res.status(401).json({ ok: false, loggedIn: false, error: '星图未登录，请在弹出的 Chrome 窗口完成扫码登录后重试' });
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
          persona: auth.persona || [], forms: auth.forms || [], category: auth.category || '',
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
    console.log('  星川服务商达人自助打标 · 本地工具已启动');
    console.log(`  本地服务：http://${HOST}:${PORT}`);
    console.log('  网页检测到连接后即可上传名单、一键打标。');
    console.log('--------------------------------------------------');
    console.log('  正在打开浏览器准备星图登录…（首次请扫码登录）');
    console.log('==================================================\n');
    ensureBrowser().then(async (page) => {
      const in_ = await checkLoggedIn(page);
      console.log(in_ ? '✅ 检测到星图已登录，可直接使用。\n' : '⚠️  未检测到登录，请在弹出的 Chrome 窗口扫码登录星图。\n');
    }).catch(e => console.log('⚠️  浏览器启动失败：' + friendlyErr(e) + '\n   若提示缺少浏览器，请运行：npx playwright install chromium\n   若提示 "has been closed"，请确认 xc-tagger 文件夹路径不含空格/括号（建议 ~/xc-tagger 或 C:\\xc-tagger）。\n'));
  });
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
function friendlyErr(e) { return (e && (e.message || String(e))) || '未知错误'; }

startServer();
