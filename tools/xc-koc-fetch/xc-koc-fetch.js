#!/usr/bin/env node
/* eslint-disable */
/**
 * xc-koc-fetch · 星川 KOC 契合度论证 · 抖音主页作品本地抓取助手 v1.1.0
 * ------------------------------------------------------------------
 * v1.1.0 关键修复：
 *   未登录态下抖音会把主页作品列表截断为「看更多最新作品 · 登录」，
 *   v1.0.0 用独立 profile 自启 Chrome（未登录）导致 0 视频、KOC 卡片
 *   误判「数据不足」。现改为：
 *   ① 主通道：connectOverCDP 127.0.0.1:9222，复用使用者本机已登录的
 *      Chrome（住宅 IP + 登录态），打开桌面版 www.douyin.com 主页，
 *      截获页面自身发出的作品列表 XHR（aweme/v1/web/aweme/post）。
 *   ② 兜底：9222 不可用时退回独立移动端 share 页；若被登录墙截断，
 *      明确返回 status:'login_required' 与原因，绝不静默返回空列表。
 *
 * 服务：仅监听 127.0.0.1:7844
 *   GET  /health -> {ok,version,mode:'cdp'|'standalone'}
 *   POST /fetch  {links:[...]} -> {ok,mode,creators:[...]}
 * 安全：只绑定本机；CORS 白名单（github.io / bytedance 内网域名 / localhost）；
 *   响应 Private Network Access 预检头（公网 https 页面访问 localhost 必需）。
 * 只采集登录用户本就能看到的公开内容，不触碰账号写操作。
 */
const express = require('express');
const cors = require('cors');
const os = require('os');
const path = require('path');
const { chromium } = require('playwright');

const VERSION = '1.1.0';
const PORT = 7844;
const CDP_URL = 'http://127.0.0.1:9222';
const UA_MOBILE = 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1';
const MAX_VIDEOS = 30;
const NAV_TIMEOUT = 35000;
const FIRST_API_TIMEOUT = 18000;

const ALLOWED = [
  /^https:\/\/([\w-]+\.)?github\.io$/,
  /^https:\/\/([\w-]+\.)?aime-app\.bytedance\.net$/,
  /^https:\/\/([\w-]+\.)?aime-site\.bytedance\.net$/,
  /^https?:\/\/localhost(:\d+)?$/,
  /^https?:\/\/127\.0\.0\.1(:\d+)?$/,
];

const app = express();
app.use(express.json({ limit: '2mb' }));
app.options('*', cors({
  origin(origin, cb) {
    if (!origin || ALLOWED.some(r => r.test(origin))) return cb(null, true);
    return cb(null, false);
  },
  allowedHeaders: ['Content-Type'],
  exposedHeaders: ['Access-Control-Allow-Private-Network'],
}));
app.use((req, res, next) => {
  const origin = req.headers.origin;
  if (origin && ALLOWED.some(r => r.test(origin))) {
    res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Access-Control-Allow-Private-Network', 'true');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
  next();
});

// ---------------- 工具 ----------------

const sleep = ms => new Promise(r => setTimeout(r, ms));

function extractSecUid(text) {
  const s = String(text || '');
  let m = s.match(/douyin\.com\/user\/([\w-]+)/);
  if (m) return m[1];
  m = s.match(/iesdouyin\.com\/share\/user\/([\w-]+)/);
  if (m) return m[1];
  m = s.match(/[?&]sec_uid=([\w-]+)/);
  if (m) return decodeURIComponent(m[1]);
  return '';
}

const statNum = v => {
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
};

function mapDesktopAweme(a) {
  try {
    const stats = a.statistics || {};
    const tags = Array.isArray(a.text_extra)
      ? a.text_extra.map(t => (t.hashtag_name ? `#${t.hashtag_name}` : '').trim()).filter(Boolean).slice(0, 6)
      : [];
    const cover = (a.video && (a.video.cover || a.video.origin_cover || {}).url_list || [])[0] || '';
    const url = `https://www.douyin.com/video/${a.aweme_id}`;
    return {
      id: String(a.aweme_id || ''),
      title: String(a.desc || '').slice(0, 300),
      url,
      stats: `赞${(stats.digg_count ?? 0)} 评${(stats.comment_count ?? 0)} 藏${(stats.collect_count ?? 0)} 转${(stats.share_count ?? 0)}`,
      desc: String(a.desc || '').slice(0, 500),
      durationSec: Math.round(Number(a.duration || 0) / 1000),
      createTime: a.create_time || 0,
      tags,
      cover,
    };
  } catch { return null; }
}

function mapMobileAweme(a) {
  try {
    const stats = a.statistics || {};
    const tags = Array.isArray(a.text_extra)
      ? a.text_extra.map(t => (t.hashtag_name ? `#${t.hashtag_name}` : '').trim()).filter(Boolean).slice(0, 6)
      : [];
    const cover = (a.video && (a.video.cover && a.video.cover.url_list || [])[0]) || '';
    return {
      id: String(a.aweme_id || ''),
      title: String(a.desc || '').slice(0, 300),
      url: `https://www.iesdouyin.com/share/video/${a.aweme_id}`,
      stats: `赞${statNum(stats.digg_count)} 评${statNum(stats.comment_count)} 藏${statNum(stats.collect_count)} 转${statNum(stats.share_count)}`,
      desc: String(a.desc || '').slice(0, 500),
      durationSec: Math.round(Number(a.duration || 0) / 1000),
      createTime: a.create_time || 0,
      tags,
      cover,
    };
  } catch { return null; }
}

// ---------------- CDP 主通道（复用本机已登录 Chrome） ----------------

let cdpBrowserPromise = null;

async function getCdpBrowser() {
  if (!cdpBrowserPromise) {
    cdpBrowserPromise = chromium.connectOverCDP(CDP_URL).catch(e => {
      cdpBrowserPromise = null;
      throw e;
    });
  }
  return cdpBrowserPromise;
}

async function cdpAvailable() {
  try {
    await getCdpBrowser();
    return true;
  } catch { return false; }
}

/**
 * 在已登录 Chrome 中打开新标签页抓桌面版主页作品。
 * 连接在服务进程内常驻复用；只关闭自己打开的页面，绝不关闭使用者的浏览器。
 */
async function fetchCreatorCdp(link) {
  const result = {
    source: 'cdp',
    input: link,
    secUid: '',
    nickname: '',
    douyinId: '',
    followers: null,
    worksCount: null,
    videos: [],
    status: 'failed',
    reason: '',
  };

  const browser = await getCdpBrowser();
  const contexts = browser.contexts();
  const ctx = contexts[0];
  if (!ctx) throw new Error('CDP 连接成功但拿不到浏览器上下文');
  const page = await ctx.newPage();
  // 标签页标题尽量低调，结束后立即关闭
  try { await page.setViewportSize({ width: 1280, height: 900 }); } catch {}

  const videoMap = new Map();
  let profile = null;
  let firstApiAt = 0;

  const onResponse = async resp => {
    try {
      const u = resp.url();
      if (u.includes('/aweme/v1/web/aweme/post')) {
        let json = null;
        try { json = await resp.json(); } catch { return; }
        const list = Array.isArray(json?.aweme_list) ? json.aweme_list : [];
        if (list.length && !firstApiAt) firstApiAt = Date.now();
        for (const a of list) {
          const v = mapDesktopAweme(a);
          if (v && v.id) videoMap.set(v.id, v);
        }
      } else if (u.includes('/aweme/v1/web/user/profile/other')) {
        try {
          const json = await resp.json();
          if (json?.user) profile = json.user;
        } catch {}
      }
    } catch {}
  };
  page.on('response', onResponse);

  try {
    const target = extractSecUid(link)
      ? (link.includes('douyin.com') ? link : `https://www.douyin.com/user/${extractSecUid(link)}`)
      : link;
    try {
      await page.goto(target, { waitUntil: 'domcontentloaded', timeout: NAV_TIMEOUT });
    } catch (e) {
      result.reason = `主页打开超时或失败：${String(e.message || e).slice(0, 120)}`;
      return result;
    }
    await sleep(2500);

    // 从落地 URL 补抓 secUid（短链跳转后）
    if (!result.secUid) {
      try { result.secUid = extractSecUid(page.url()) || extractSecUid(link); } catch {}
    }

    // 等首屏作品 XHR，最多 18s；期间持续累积
    const t0 = Date.now();
    while (Date.now() - t0 < FIRST_API_TIMEOUT && videoMap.size === 0) {
      await sleep(700);
    }

    // 滚动触发翻页 XHR（3 屏），拿满 MAX_VIDEOS 即停
    for (let i = 0; i < 5 && videoMap.size < MAX_VIDEOS; i++) {
      try { await page.mouse.wheel(0, 2400); } catch {}
      await sleep(1600);
    }
    // 给最后一个 XHR 一点解析时间
    await sleep(800);

    // —— 基础资料：优先 profile XHR，其次首条视频 author，再次页面文本兜底 ——
    if (profile) {
      result.nickname = String(profile.nickname || '');
      result.douyinId = String(profile.unique_id || profile.short_id || '');
      result.followers = statNum(profile.follower_count) || null;
      result.worksCount = statNum(profile.aweme_count) || null;
    }
    if (!result.nickname && videoMap.size) {
      try {
        const firstAuthor = JSON.stringify([...videoMap.values()]); // author 已在 map 时丢弃，改用 DOM
        void firstAuthor;
      } catch {}
    }
    if (!result.nickname) {
      try {
        result.nickname = (await page.locator('h1').first().textContent({ timeout: 2000 }).catch(() => '') || '').trim();
      } catch {}
    }

    // —— 登录墙识别 ——
    let wallText = '';
    try {
      wallText = (await page.locator('body').innerText({ timeout: 2000 }).catch(() => '') || '');
    } catch {}
    const loginWall = /看更多最新作品|登录后即可看更多|登录之后看更多|扫码登录|登录查看更多/.test(wallText);

    result.videos = [...videoMap.values()]
      .sort((a, b) => (b.createTime || 0) - (a.createTime || 0))
      .slice(0, MAX_VIDEOS);

    if (result.videos.length) {
      result.status = 'ok';
      result.reason = '';
    } else if (loginWall) {
      result.status = 'login_required';
      result.reason = '已连接 9222 Chrome，但该 Chrome 未登录抖音（主页出现登录墙，作品列表被截断）。请在这个 Chrome 里登录 www.douyin.com 后重试。';
    } else if (firstApiAt) {
      result.status = 'empty';
      result.reason = '主页作品接口返回为空（可能账号无公开作品或被风控临时限制）';
    } else {
      result.status = 'failed';
      result.reason = '未截获到主页作品接口响应（页面加载不完整或被风控验证页拦截）';
    }
    return result;
  } finally {
    try { page.removeListener('response', onResponse); } catch {}
    try { await page.close(); } catch {}
  }
}

// ---------------- 独立模式兜底（移动 share 页 + 登录墙检测） ----------------

async function fetchCreatorStandalone(link, browser) {
  const result = {
    source: 'standalone',
    input: link,
    secUid: extractSecUid(link),
    nickname: '',
    douyinId: '',
    followers: null,
    worksCount: null,
    videos: [],
    status: 'failed',
    reason: '',
  };
  const page = await browser.newPage();
  try {
    await page.setExtraHTTPHeaders({ Referer: 'https://www.iesdouyin.com/' });
    await page.setViewportSize({ width: 390, height: 844 });

    const videos = [];
    const onResp = async resp => {
      try {
        if (!resp.url().includes('/web/api/v2/aweme/post')) return;
        const j = await resp.json();
        for (const a of (Array.isArray(j.aweme_list) ? j.aweme_list : [])) {
          const v = mapMobileAweme(a);
          if (v && v.id && !videos.some(x => x.id === v.id)) videos.push(v);
        }
      } catch {}
    };
    page.on('response', onResp);

    const url = result.secUid
      ? `https://www.iesdouyin.com/share/user/${result.secUid}`
      : link;
    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: NAV_TIMEOUT });
    } catch (e) {
      result.reason = `移动主页打开失败：${String(e.message || e).slice(0, 100)}`;
      return result;
    }
    await sleep(4000);
    try { result.nickname = (await page.locator('.jsx-791823409235732f05b,.nickname,h1').first().textContent({ timeout: 1500 }).catch(() => '') || '').replace(/\s+/g, ' ').trim(); } catch {}
    for (let i = 0; i < 4 && videos.length < MAX_VIDEOS; i++) {
      try { await page.mouse.wheel(0, 1600); } catch {}
      await sleep(1500);
    }
    await sleep(800);

    let bodyText = '';
    try { bodyText = await page.locator('body').innerText({ timeout: 1500 }).catch(() => ''); } catch {}
    const loginWall = /看更多最新作品|登录后看更多|登录查看更多|登录抖音|登录后即可/.test(bodyText);

    result.videos = videos.sort((a, b) => (b.createTime || 0) - (a.createTime || 0)).slice(0, MAX_VIDEOS);
    if (result.videos.length) {
      result.status = 'ok';
    } else if (loginWall) {
      result.status = 'login_required';
      result.reason = '未登录状态下作品列表被抖音截断。请改用「调试端口方式」：完全退出 Chrome 后执行  open -a "Google Chrome" --args --remote-debugging-port=9222，并在该 Chrome 登录抖音后重试。';
    } else {
      result.status = 'empty';
      result.reason = '未抓到公开作品（无作品或被风控）';
    }
    return result;
  } finally {
    try { await page.close(); } catch {}
  }
}

// ---------------- 服务 ----------------

app.get('/health', (req, res) => {
  // health 不做阻塞探测，返回当前已知状态；mode 由 /fetch 实际使用通道决定
  res.json({ ok: true, service: 'xc-koc-fetch', version: VERSION, port: PORT });
});

app.post('/fetch', async (req, res) => {
  const links = Array.isArray(req.body?.links)
    ? req.body.links.map(x => String(x || '').trim()).filter(Boolean)
    : [];
  if (!links.length) return res.status(400).json({ ok: false, error: '缺少 links' });

  const useCdp = await cdpAvailable();
  const creators = [];
  const warnings = [];

  if (useCdp) {
    // 主通道：逐个用已登录 Chrome 抓
    for (const link of links) {
      try {
        const r = await fetchCreatorCdp(link);
        creators.push(r);
      } catch (e) {
        creators.push({ input: link, secUid: extractSecUid(link), videos: [], status: 'failed', reason: String(e.message || e).slice(0, 160) });
      }
    }
  } else {
    warnings.push('未检测到 9222 调试端口 Chrome，已退回独立模式（未登录态作品可能被截断）。建议：完全退出 Chrome 后执行 open -a "Google Chrome" --args --remote-debugging-port=9222 并登录抖音。');
    let browser = null;
    try {
      // 独立模式：持久化 profile（首次 headful 登录后，后续可复用登录态）
      const profileDir = path.join(os.homedir(), '.xc-koc-fetch-profile');
      browser = await chromium.launchPersistentContext(profileDir, {
        headless: false,
        channel: 'chrome',
        isMobile: true,
        userAgent: UA_MOBILE,
        viewport: { width: 390, height: 844 },
        args: ['--disable-blink-features=AutomationControlled'],
      });
    } catch (e) {
      warnings.push(`独立模式 headful 启动失败（${String(e.message || e).slice(0, 100)}），尝试无头兜底`);
      try {
        browser = await chromium.launch({
          headless: true,
          channel: 'chrome',
          args: ['--disable-blink-features=AutomationControlled'],
        });
      } catch (e2) {
        return res.status(500).json({ ok: false, error: '本机 Chrome 启动失败：' + String(e2.message || e2).slice(0, 160) });
      }
    }
    try {
      for (const link of links) {
        try {
          const r = await fetchCreatorStandalone(link, browser);
          creators.push(r);
        } catch (e) {
          creators.push({ input: link, secUid: extractSecUid(link), videos: [], status: 'failed', reason: String(e.message || e).slice(0, 160) });
        }
      }
    } finally {
      try { await browser?.close(); } catch {}
    }
  }

  res.json({
    ok: true,
    service: 'xc-koc-fetch',
    version: VERSION,
    mode: useCdp ? 'cdp' : 'standalone',
    warnings,
    creators,
  });
});

app.listen(PORT, '127.0.0.1', () => {
  console.log(`[xc-koc-fetch v${VERSION}] 本机抓取服务已启动 http://127.0.0.1:${PORT}`);
  console.log('优先连接本机已登录 Chrome（CDP 127.0.0.1:9222）；不可用时退回独立模式。');
});
