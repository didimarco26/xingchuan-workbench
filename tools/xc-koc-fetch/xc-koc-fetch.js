#!/usr/bin/env node
/* eslint-disable */
/**
 * xc-koc-fetch · 星川 KOC 契合度论证 · 抖音主页作品本地抓取助手 v1.0.0
 * ------------------------------------------------------------------
 * 用途：KOC 面板（对内工作台 / 公开版 gh-pages）在浏览器里无法直接抓取
 *   抖音主页近期作品（数据中心 IP 与跨域均被风控）。本助手在使用者本机
 *   运行，复用本机真实 Chrome（住宅 IP + 真实浏览器签名），以移动端
 *   share 页打开达人主页，截获页面自身发出的作品列表 XHR：
 *     https://www.iesdouyin.com/web/api/v2/aweme/post/ （reflow 通道）
 *   拿到近期 15~30 条作品（标题文案/互动/标签/时长/封面），供 KOC
 *   五轴评分做内容样本。只采集公开可见内容，不登录、不绕过访问控制。
 *
 * 服务：仅监听 127.0.0.1:7844
 *   GET  /health -> {ok,version}
 *   POST /fetch  {links:[...]} -> {ok,creators:[{secUid,nickname,videos}]}
 * 安全：只绑定本机；CORS 白名单（github.io / bytedance 内网域名 / localhost）；
 *   响应 Private Network Access 预检头（公网 https 页面访问 localhost 必需）。
 * 浏览器：Playwright 调系统 Chrome（channel=chrome，失败退 msedge），
 *   默认无头静默运行；XC_KOC_HEADFUL=1 可有头调试。
 */
const express = require('express');
const cors = require('cors');
const path = require('path');
const os = require('os');
const fs = require('fs');
const http = require('http');
const { chromium } = require('playwright');

const VERSION = '1.0.0';
const HOST = '127.0.0.1';
const PORT = Number(process.env.XC_KOC_PORT || 7844);
const HEADFUL = process.env.XC_KOC_HEADFUL === '1';
const UA_MOBILE = 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1';
const PER_CREATOR_TIMEOUT = 55000;
const MAX_PAGES = 25; // 每个达人最多收集作品数（含置顶）
const PROFILE_DIR = process.env.XC_KOC_PROFILE || path.join(os.homedir(), '.xc-koc-fetch-profile');
const PUBLIC_DIR = path.join(__dirname, 'public');

const app = express();

// ---------- CORS（含 PNA 预检） ----------
function originAllowed(origin) {
  if (!origin) return true;
  try {
    const u = new URL(origin);
    if (u.hostname === 'localhost' || u.hostname === '127.0.0.1') return true;
    if (u.hostname === 'didimarco26.github.io') return true;
    if (u.hostname.endsWith('.github.io') && u.hostname.startsWith('didimarco26')) return true;
    if (u.hostname.endsWith('.bytedance.net') || u.hostname.endsWith('.bytedance.com')) return true;
  } catch (_) {}
  return false;
}
app.use((req, res, next) => {
  const origin = req.headers.origin;
  if (originAllowed(origin)) {
    if (origin) res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Vary', 'Origin');
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    res.setHeader('Access-Control-Allow-Private-Network', 'true'); // Chrome PNA 预检
  }
  if (req.method === 'OPTIONS') return res.status(204).end();
  next();
});
app.use(cors({ origin: (o, cb) => cb(null, originAllowed(o)) }));
app.use(express.json({ limit: '2mb' }));
app.use(express.static(PUBLIC_DIR));

// ---------- 浏览器单例（持久 profile，移动端上下文） ----------
let _ctx = null;
let _ctxBusy = false;
async function getContext() {
  if (_ctx) return _ctx;
  fs.mkdirSync(PROFILE_DIR, { recursive: true });
  const launchOpts = {
    headless: !HEADFUL,
    locale: 'zh-CN',
    viewport: { width: 414, height: 800 },
    userAgent: UA_MOBILE,
    args: ['--no-sandbox', '--disable-blink-features=AutomationControlled', '--lang=zh-CN', '--mute-audio'],
  };
  // channel=chrome 失败则退系统自带 Chromium/Edge
  try {
    _ctx = await chromium.launchPersistentContext(PROFILE_DIR, { ...launchOpts, channel: 'chrome' });
  } catch (e) {
    try { _ctx = await chromium.launchPersistentContext(PROFILE_DIR, { ...launchOpts, channel: 'msedge' }); }
    catch (e2) { _ctx = await chromium.launchPersistentContext(PROFILE_DIR, launchOpts); }
  }
  await _ctx.addInitScript(() => { try { Object.defineProperty(navigator, 'webdriver', { get: () => false }); } catch (_) {} });
  _ctx.on('close', () => { _ctx = null; });
  return _ctx;
}

// ---------- 链接归一化：提取 sec_uid（必要时跟随短链） ----------
const SEC_RE = /(MS4wLjABAAAA[a-zA-Z0-9_\-]+)/;
async function resolveSecUid(rawUrl) {
  let u = String(rawUrl || '').trim();
  if (!u) return null;
  // 纯文本里夹链接
  const m0 = u.match(/https?:\/\/[^\s，,]+/);
  if (m0) u = m0[0];
  const m1 = u.match(SEC_RE);
  if (m1) return { secUid: m1[1], kind: 'user' };
  const um = u.match(/douyin\.com\/user\/([A-Za-z0-9_\-]+)/);
  if (um) return { secUid: um[1], kind: 'user' };
  const vm = u.match(/(?:douyin\.com|iesdouyin\.com)\/(?:share\/)?video\/(\d+)/);
  const videoId = vm ? vm[1] : null;
  // 短链（v.douyin.com）跟随 302
  if (/v\.douyin\.com|v\.iesdouyin\.com/.test(u) || videoId) {
    const followed = await followRedirects(videoId ? `https://www.iesdouyin.com/share/video/${videoId}` : u);
    const m2 = (followed.finalUrl || '').match(SEC_RE) || (followed.body || '').match(SEC_RE);
    if (m2) return { secUid: m2[1], kind: 'user', fromVideo: !!videoId };
    return videoId ? { secUid: null, kind: 'video', videoId, finalUrl: followed.finalUrl } : null;
  }
  return null;
}
async function followRedirects(startUrl) {
  let url = startUrl;
  let body = '';
  for (let i = 0; i < 6; i++) {
    try {
      const r = await fetch(url, { redirect: 'manual', headers: { 'User-Agent': UA_MOBILE } });
      if (r.status >= 300 && r.status < 400 && r.headers.get('location')) {
        const loc = r.headers.get('location');
        url = /^https?:\/\//.test(loc) ? loc : new URL(loc, url).href;
        continue;
      }
      if (r.status === 200) body = (await r.text()).slice(0, 200000);
      return { finalUrl: url, body, status: r.status };
    } catch (e) { return { finalUrl: url, body, err: String(e).slice(0, 120) }; }
  }
  return { finalUrl: url, body };
}

// ---------- 核心：抓取单个达人近期作品 ----------
async function scrapeCreator(secUid) {
  const ctx = await getContext();
  const page = await ctx.newPage();
  const batches = [];
  let profile = {};
  const onResp = async (r) => {
    const u = r.url();
    if (/web\/api\/v2\/aweme\/post/.test(u) && r.status() === 200) {
      try {
        const j = JSON.parse((await r.body()).toString('utf8'));
        if (j && j.status_code === 0 && Array.isArray(j.aweme_list) && j.aweme_list.length) {
          batches.push(j);
          const a = j.aweme_list.find(x => x && x.author) || {};
          if (a.author && (!profile.nickname || !profile.followers)) {
            profile = {
              nickname: a.author.nickname || profile.nickname || '',
              douyinId: a.author.unique_id || profile.douyinId || '',
              followers: a.author.follower_count || profile.followers || 0,
            };
          }
        }
      } catch (_) {}
    }
  };
  page.on('response', onResp);
  const timer = setTimeout(() => {}, 0); // 占位
  clearTimeout(timer);
  try {
    await page.goto('https://www.iesdouyin.com/share/user/' + secUid + '?from_ssr=1', { waitUntil: 'domcontentloaded', timeout: 40000 });
    const deadline = Date.now() + PER_CREATOR_TIMEOUT;
    let lastCount = 0, stall = 0;
    while (Date.now() < deadline) {
      await page.waitForTimeout(1500);
      const n = batches.reduce((s, b) => s + b.aweme_list.length, 0);
      if (n >= MAX_PAGES) break;
      // 滚动触发翻页 XHR
      await page.mouse.wheel(0, 700).catch(() => {});
      if (n === lastCount) { stall++; if (stall >= 8) break; } else { stall = 0; lastCount = n; }
    }
  } finally {
    page.removeListener('response', onResp);
    await page.close().catch(() => {});
  }
  // 汇总去重
  const seen = new Set();
  const videos = [];
  for (const b of batches) {
    for (const it of b.aweme_list) {
      if (!it || !it.aweme_id || seen.has(it.aweme_id)) continue;
      seen.add(it.aweme_id);
      const st = it.statistics || {};
      const cover = (it.video && it.video.cover && it.video.cover.url_list && it.video.cover.url_list[0])
        || (it.video && it.video.origin_cover && it.video.origin_cover.url_list && it.video.origin_cover.url_list[0]) || '';
      videos.push({
        awemeId: String(it.aweme_id),
        title: String(it.desc || '').replace(/\s+/g, ' ').trim().slice(0, 200),
        createTime: it.create_time || 0,
        isTop: it.is_top ? 1 : 0,
        duration: it.video && it.video.duration ? it.video.duration : 0,
        digg: st.digg_count || 0, comment: st.comment_count || 0, share: st.share_count || 0,
        collect: st.collect_count || 0, play: st.play_count || 0,
        tags: (it.text_extra || []).map(t => t.hashtag_name).filter(Boolean).slice(0, 8),
        cover,
        url: 'https://www.douyin.com/video/' + it.aweme_id,
      });
      if (videos.length >= MAX_PAGES) break;
    }
    if (videos.length >= MAX_PAGES) break;
  }
  return { profile, videos };
}

// 串行队列（同一浏览器上下文，避免相互干扰）
let _chain = Promise.resolve();
function queue(task) {
  const run = _chain.then(task, task);
  _chain = run.catch(() => {});
  return run;
}

// ---------- HTTP 接口 ----------
app.get('/health', (req, res) => {
  res.json({ ok: true, service: 'xc-koc-fetch', version: VERSION, port: PORT, headful: HEADFUL });
});

app.post('/fetch', async (req, res) => {
  const links = Array.isArray(req.body && req.body.links) ? req.body.links.slice(0, 30) : [];
  if (!links.length) return res.status(400).json({ ok: false, error: 'links 不能为空' });
  if (_ctxBusy) return res.status(429).json({ ok: false, error: '上一批抓取进行中，请稍候' });
  _ctxBusy = true;
  try {
    // 归一化去重
    const targets = [];
    const secSeen = new Set();
    const unresolved = [];
    for (const raw of links) {
      const r = await resolveSecUid(raw);
      if (r && r.secUid && !secSeen.has(r.secUid)) { secSeen.add(r.secUid); targets.push({ secUid: r.secUid, source: String(raw).slice(0, 300) }); }
      else if (!r || !r.secUid) unresolved.push({ link: String(raw).slice(0, 300), reason: r ? '仅视频链接，未能反查达人主页' : '无法识别链接' });
    }
    const creators = [];
    for (const t of targets) {
      const one = await queue(() => scrapeCreator(t.secUid));
      creators.push({
        secUid: t.secUid,
        source: t.source,
        nickname: one.profile.nickname || '',
        douyinId: one.profile.douyinId || '',
        followers: one.profile.followers || 0,
        videoCount: one.videos.length,
        videos: one.videos,
        status: one.videos.length ? 'ok' : 'no_videos',
      });
    }
    res.json({ ok: true, version: VERSION, creators, unresolved });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e && e.message || e).slice(0, 200) });
  } finally {
    _ctxBusy = false;
  }
});

// 端口占用时自动顺延
const server = http.createServer(app);
server.on('error', (e) => {
  if (e.code === 'EADDRINUSE') {
    console.error(`[xc-koc-fetch] 端口 ${PORT} 已被占用（助手可能已在运行），无需重复启动。`);
    process.exit(0);
  }
  throw e;
});
server.listen(PORT, HOST, () => {
  console.log('========================================');
  console.log(`  xc-koc-fetch v${VERSION} 已启动`);
  console.log(`  本地服务：http://${HOST}:${PORT}`);
  console.log(`  浏览器：系统 Chrome · ${HEADFUL ? '有头模式' : '无头静默'} · profile ${PROFILE_DIR}`);
  console.log('  请回到星川工作台页面使用（此窗口保持打开）');
  console.log('========================================');
  if (process.env.XC_KOC_OPEN !== '0') {
    const open = process.platform === 'darwin' ? 'open' : process.platform === 'win32' ? 'start' : 'xdg-open';
    require('child_process').exec(`${open} http://${HOST}:${PORT}`).unref();
  }
});

process.on('SIGINT', async () => { try { _ctx && await _ctx.close(); } catch (_) {} process.exit(0); });
process.on('SIGTERM', async () => { try { _ctx && await _ctx.close(); } catch (_) {} process.exit(0); });
