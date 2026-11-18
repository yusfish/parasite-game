/**
 * 寄生蟲卵辨識挑戰 — 全科排行榜後端（Google Apps Script）
 * 部署：擴充功能 › Apps Script › 貼上本檔 › 部署 › 新增部署作業 › 網頁應用程式
 *       執行身分：我　／　誰可以存取：所有人
 * 資料存在綁定的試算表「scores」工作表，每位挑戰者（裝置 ID）一列，各難度只留最高分。
 */
const SHEET = 'scores';
const HEAD = ['id', 'name', 'b1', 't1', 'b2', 't2', 'updated'];
const MAX_SCORE = 2400;          // 12 題 × 每題最高 200 分
const MIN_GAP_SEC = 20;          // 同一裝置兩次送分的最短間隔
const CACHE_SEC = 30;            // 排行榜讀取快取，減輕試算表負擔

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) { sh = ss.insertSheet(SHEET); sh.appendRow(HEAD); sh.setFrozenRows(1); }
  return sh;
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function rows_() {
  const v = sheet_().getDataRange().getValues();
  return v.slice(1).filter(r => r[0]).map(r => ({
    id: String(r[0]), name: String(r[1]),
    b1: r[2] === '' ? undefined : Number(r[2]), t1: r[3] ? String(r[3]) : undefined,
    b2: r[4] === '' ? undefined : Number(r[4]), t2: r[5] ? String(r[5]) : undefined
  }));
}

// 讀排行榜（公開資料只回傳暱稱與分數，裝置 ID 只回傳雜湊後的短碼）
function doGet() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('board');
  if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  const body = JSON.stringify({ ok: true, rows: rows_().map(r => ({ ...r, id: tag_(r.id) })) });
  if (body.length < 90000) cache.put('board', body, CACHE_SEC);
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

// 送分：{id, name, lv, score}
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, err: 'bad json' }); }
  const id = String(d.id || ''), lv = Number(d.lv), score = Number(d.score);
  const name = String(d.name || '匿名').replace(/[\u0000-\u001f<>]/g, '').trim().slice(0, 12) || '匿名';
  if (!/^[a-z0-9-]{16,40}$/.test(id)) return out_({ ok: false, err: 'bad id' });
  if (lv !== 1 && lv !== 2) return out_({ ok: false, err: 'bad level' });
  if (!Number.isInteger(score) || score < 0 || score > MAX_SCORE) return out_({ ok: false, err: 'bad score' });

  const cache = CacheService.getScriptCache();
  if (cache.get('gap:' + id)) return out_({ ok: false, err: 'too fast' });

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(8000)) return out_({ ok: false, err: 'busy' });
  try {
    const sh = sheet_(), v = sh.getDataRange().getValues(), now = new Date().toISOString();
    let rowIdx = -1;
    for (let i = 1; i < v.length; i++) if (String(v[i][0]) === id) { rowIdx = i; break; }
    const bc = lv === 1 ? 2 : 4;                       // b1 在第 3 欄（索引 2）、b2 在第 5 欄（索引 4）
    if (rowIdx < 0) {
      const r = [id, name, '', '', '', '', now]; r[bc] = score; r[bc + 1] = now;
      sh.appendRow(r);
    } else {
      const r = v[rowIdx]; r[1] = name; r[6] = now;
      if (r[bc] === '' || Number(r[bc]) < score) { r[bc] = score; r[bc + 1] = now; }
      sh.getRange(rowIdx + 1, 1, 1, HEAD.length).setValues([r]);
    }
    cache.put('gap:' + id, '1', MIN_GAP_SEC);
    cache.remove('board');
  } finally { lock.releaseLock(); }
  return out_({ ok: true, me: tag_(id) });
}

function tag_(id) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'pe:' + id);
  return b.slice(0, 6).map(x => ((x + 256) % 256).toString(16).padStart(2, '0')).join('');
}
