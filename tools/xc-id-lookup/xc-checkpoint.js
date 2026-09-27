/* xc-checkpoint.js — 本机批量任务「断点续跑」通用模块
 * 适用于 xc-tagger（达人打标）与 xc-id-lookup（星图达人匹配）。
 *
 * 设计：
 *  - 任务身份 jobId = sha256(工具 | 来源 | 行数 | 每行输入指纹)，
 *    同一份名单（顺序一致）每次计算结果相同；
 *  - 每完成一条立即把结果写入 checkpoint（原子写：tmp+rename），
 *    中途关闭窗口 / 断网 / 崩溃，已完成的行都不丢；
 *  - 续跑时按行号取回已完成结果，仅处理剩余行；全部完成则直接返回；
 *  - 默认仅保留最近 20 个断点，自动清理更旧的。
 */
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');

// 单行输入指纹（任一输入列 + 记录ID），用于任务身份与完整性校验
function rowKey(it) {
  return [it.recordId, it.id, it.douyin, it.name]
    .map((s) => String(s == null ? '' : s).trim())
    .join('#');
}

function computeJobId(tool, source, items) {
  const h = crypto.createHash('sha256');
  h.update(String(tool) + '|' + String(source) + '|' + items.length);
  for (const it of items || []) h.update('\n' + rowKey(it));
  return h.digest('hex').slice(0, 20);
}

function fileOf(dir, jobId) {
  return path.join(dir, jobId + '.json');
}

function load(dir, jobId) {
  try {
    return JSON.parse(fs.readFileSync(fileOf(dir, jobId), 'utf8'));
  } catch (_) {
    return null;
  }
}

// 轻量摘要（不含完整结果，供前端提示）
function peek(dir, jobId) {
  const cp = load(dir, jobId);
  if (!cp) return null;
  const results = cp.results || {};
  return {
    found: true,
    jobId,
    source: cp.source,
    total: cp.total,
    done: Object.keys(results).length,
    finished: !!cp.finished,
    createdAt: cp.createdAt,
    updatedAt: cp.updatedAt,
  };
}

function createCp(info) {
  const now = Date.now();
  return {
    v: 1,
    tool: info.tool,
    source: info.source,
    total: info.total,
    meta: info.meta || {},
    createdAt: now,
    updatedAt: now,
    finished: false,
    results: {}, // { [行号(row, 字符串)]: 该行结果 }
  };
}

// 原子保存（先写临时文件再 rename），并顺手清理旧断点
function save(dir, jobId, cp) {
  fs.mkdirSync(dir, { recursive: true });
  cp.updatedAt = Date.now();
  const tmp = fileOf(dir, jobId) + '.tmp-' + process.pid;
  fs.writeFileSync(tmp, JSON.stringify(cp));
  fs.renameSync(tmp, fileOf(dir, jobId));
  prune(dir);
}

function discard(dir, jobId) {
  try {
    fs.unlinkSync(fileOf(dir, jobId));
  } catch (_) { /* 不存在视为已清理 */ }
  return true;
}

function prune(dir, keep) {
  keep = keep || 20;
  let files = [];
  try {
    files = fs.readdirSync(dir)
      .filter((f) => f.endsWith('.json'))
      .map((f) => ({ f, mtime: fs.statSync(path.join(dir, f)).mtimeMs }));
  } catch (_) {
    return;
  }
  files.sort((a, b) => b.mtime - a.mtime);
  for (const it of files.slice(keep)) {
    try { fs.unlinkSync(path.join(dir, it.f)); } catch (_) { /* 忽略 */ }
  }
}

module.exports = {
  rowKey,
  computeJobId,
  load,
  peek,
  createCp,
  save,
  discard,
  prune,
};
