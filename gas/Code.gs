/**
 * KURFC マルチアングル 短縮リンク ＋ みんなのメモ
 *
 *  POST 本文が "z..." の文字列             → 共有データを保存して {"k":"Xy7Pq2","b":"Xy7Pq2"}（古い形式）
 *  POST {"op":"save","d":"z...","from":"Ab3x9Q"}
 *                                          → 保存。from と同じ「メモ帳」(b) を使う → {"k","b"}
 *  POST {"op":"note","b","id","t","cam","text","by","re","own"} → みんなのメモを1件追加
 *  POST {"op":"del","b","id","own"}        → 自分が書いたメモを消す（own が一致したときだけ）
 *  GET  ?k=Xy7Pq2[&o=端末ID]               → {"d":"z...","b":"...","notes":[...]}
 *  GET  ?b=メモ帳ID[&o=端末ID]              → {"notes":[...]}
 *
 * スプレッドシート「KURFCマルチアングル 短縮リンク」
 *  links : A=ID  B=作成日時  C=指紋(MD5)  D以降=データ（長い場合は分割）
 *  boards: A=ID  B=メモ帳ID（同じ分析から作ったリンクは同じメモ帳）
 *  notes : A=メモ帳ID B=メモID C=時刻(秒) D=カメラ E=メモ F=名前 G=返信先 H=書いた日時 I=端末 J=削除
 */
const SSID = '1iU_W3YgJmshZ3oKep6ffJiSb303LpdKrev19Wbf4YFs'; // スプレッドシート「KURFCマルチアングル 短縮リンク」
const SHEET = 'links';
const CHUNK = 45000;          // 1セルの上限(5万文字)より少し小さく
const MAX_LEN = 400000;       // 異常に大きいデータは受け付けない
const ID_LEN = 6;
const ABC = 'abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 紛らわしい l/1/O/0/I を除く
const MAX_NOTES = 3000;       // 1つのメモ帳に書けるメモの上限
const ID_RE = /^[A-Za-z0-9]{4,12}$/;
const TOKEN_RE = /^[A-Za-z0-9_-]{4,40}$/;

let ss_;
function book_() { return ss_ || (ss_ = SpreadsheetApp.openById(SSID)); }
function sheetOf_(name, header) {
  const ss = book_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(header);
    sh.setFrozenRows(1);
  }
  return sh;
}
const sheet_ = () => sheetOf_(SHEET, ['ID', '作成日時', '指紋', 'データ']);
const boards_ = () => sheetOf_('boards', ['ID', 'メモ帳ID']);
const notes_ = () => sheetOf_('notes', ['メモ帳ID', 'メモID', '時刻', 'カメラ', 'メモ', '名前', '返信先', '書いた日時', '端末', '削除']);

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function md5_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function newId_() {
  let s = '';
  for (let i = 0; i < ID_LEN; i++) s += ABC[Math.floor(Math.random() * ABC.length)];
  return s;
}
function find_(sh, col, value) {
  if (sh.getLastRow() < 2) return 0;
  const r = sh.getRange(2, col, sh.getLastRow() - 1, 1)
    .createTextFinder(value).matchEntireCell(true).matchCase(true).findNext();
  return r ? r.getRow() : 0;
}
// 先頭の ' は「文字として保存」の印（数字や式に化けないように）。読み出し時には付きません
const txt_ = v => "'" + String(v == null ? '' : v);

function boardOf_(k) {
  const sh = boards_();
  const row = find_(sh, 1, k);
  return row ? String(sh.getRange(row, 2).getValue()) : k; // 以前のリンクは自分のIDがメモ帳
}

/* ---------- 共有データ ---------- */
function saveLink_(d, from) {
  d = String(d || '').trim();
  if (!/^[zj][A-Za-z0-9_-]+$/.test(d) || d.length > MAX_LEN) return {error: 'bad data'};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const fp = md5_(d);
    const same = find_(sh, 3, fp);                     // 同じ内容なら同じIDを返す
    if (same) {
      const k = String(sh.getRange(same, 1).getValue());
      let b = boardOf_(k);
      // 同じ内容のリンクが前からあっても、まだメモの無いリンクなら from と同じメモ帳にそろえる
      if (ID_RE.test(from || '') && from !== k && b === k && find_(sh, 1, from)) {
        const fb = boardOf_(from);
        if (fb !== k && !allNotes_(k).length) {
          const bs = boards_(), row = find_(bs, 1, k);
          if (row) bs.getRange(row, 2).setValue(txt_(fb)); else bs.appendRow([txt_(k), txt_(fb)]);
          b = fb;
        }
      }
      return {k: k, b: b};
    }
    let id;
    do { id = newId_(); } while (find_(sh, 1, id));
    const parts = [];
    for (let i = 0; i < d.length; i += CHUNK) parts.push(d.slice(i, i + CHUNK));
    sh.appendRow([txt_(id), new Date(), txt_(fp)].concat(parts.map(txt_)));
    // 同じ分析から作り直したリンクは、前のリンクと同じメモ帳を使う
    let b = id;
    if (ID_RE.test(from || '') && find_(sh, 1, from)) b = boardOf_(from);
    boards_().appendRow([txt_(id), txt_(b)]);
    return {k: id, b: b};
  } finally {
    lock.releaseLock();
  }
}
function readLink_(k) {
  const sh = sheet_();
  const row = find_(sh, 1, k);
  if (!row) return null;
  const vals = sh.getRange(row, 4, 1, Math.max(sh.getLastColumn() - 3, 1)).getValues()[0];
  return vals.filter(String).join('');
}

/* ---------- みんなのメモ ---------- */
function allNotes_(b) {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('n:' + b);
  if (hit) return JSON.parse(hit);
  const sh = notes_();
  const list = [];
  if (sh.getLastRow() >= 2) {
    const found = sh.getRange(2, 1, sh.getLastRow() - 1, 1)
      .createTextFinder(b).matchEntireCell(true).matchCase(true).findAll();
    found.forEach(r => {
      const v = sh.getRange(r.getRow(), 1, 1, 10).getValues()[0];
      if (v[9]) return; // 削除済み
      list.push({id: String(v[1]), t: +v[2] || 0, cam: v[3] === '' ? -1 : +v[3], text: String(v[4]), by: String(v[5]),
        re: String(v[6] || ''), at: +v[7] || 0, own: String(v[8])});
    });
  }
  list.sort((x, y) => x.t - y.t || x.at - y.at);
  const s = JSON.stringify(list);
  if (s.length < 90000) cache.put('n:' + b, s, 21600);
  return list;
}
function publicNotes_(b, me) {
  return allNotes_(b).map(n => ({id: n.id, t: n.t, cam: n.cam, text: n.text, by: n.by, re: n.re, at: n.at, mine: !!me && n.own === me}));
}
function addNote_(o) {
  const b = String(o.b || ''), own = String(o.own || '');
  const text = String(o.text || '').replace(/\s+$/, '').slice(0, 500);
  const by = String(o.by || '').trim().slice(0, 30);
  const t = Number(o.t), cam = Math.round(Number(o.cam));
  const re = /^[A-Za-z0-9_-]{1,24}$/.test(o.re || '') ? String(o.re) : '';
  let id = /^[A-Za-z0-9_-]{4,24}$/.test(o.id || '') ? String(o.id) : '';
  if (!ID_RE.test(b) || !TOKEN_RE.test(own) || !text.trim() || !by || !isFinite(t) || t < -600 || t > 86400) return {error: 'bad note'};
  if (!find_(sheet_(), 1, b)) return {error: 'no board'};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const list = allNotes_(b);
    if (list.length >= MAX_NOTES) return {error: 'full'};
    if (!id || list.some(n => n.id === id)) id = Date.now().toString(36) + newId_();
    const at = Date.now();
    notes_().appendRow([txt_(b), txt_(id), t, isFinite(cam) && cam >= 0 && cam < 12 ? cam : -1, txt_(text), txt_(by), txt_(re), at, txt_(own), '']);
    CacheService.getScriptCache().remove('n:' + b);
    return {ok: true, note: {id: id, t: t, cam: cam, text: text, by: by, re: re, at: at, mine: true}};
  } finally {
    lock.releaseLock();
  }
}
function delNote_(o) {
  const b = String(o.b || ''), id = String(o.id || ''), own = String(o.own || '');
  if (!ID_RE.test(b) || !id || !TOKEN_RE.test(own)) return {error: 'bad'};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = notes_();
    if (sh.getLastRow() < 2) return {error: 'not found'};
    const found = sh.getRange(2, 2, sh.getLastRow() - 1, 1)
      .createTextFinder(id).matchEntireCell(true).matchCase(true).findAll();
    for (const r of found) {
      const v = sh.getRange(r.getRow(), 1, 1, 10).getValues()[0];
      if (String(v[0]) === b && String(v[8]) === own) {
        sh.getRange(r.getRow(), 10).setValue(new Date());
        CacheService.getScriptCache().remove('n:' + b);
        return {ok: true};
      }
    }
    return {error: 'not yours'};
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 入口 ---------- */
function doPost(e) {
  const raw = String((e && e.postData && e.postData.contents) || '').trim();
  if (raw[0] !== '{') return json_(saveLink_(raw, ''));
  let o;
  try { o = JSON.parse(raw); } catch (err) { return json_({error: 'bad json'}); }
  if (o.op === 'save') return json_(saveLink_(o.d, String(o.from || '')));
  if (o.op === 'note') return json_(addNote_(o));
  if (o.op === 'del') return json_(delNote_(o));
  return json_({error: 'bad op'});
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const me = TOKEN_RE.test(p.o || '') ? String(p.o) : '';
  if (ID_RE.test(p.k || '')) {
    const d = readLink_(String(p.k));
    if (!d) return json_({error: 'not found'});
    const b = boardOf_(String(p.k));
    return json_({d: d, b: b, notes: publicNotes_(b, me)});
  }
  if (ID_RE.test(p.b || '')) return json_({notes: publicNotes_(String(p.b), me)});
  return json_({ok: true, app: 'KURFC multi-angle short links'});
}

// 最初に1回だけ実行：権限の許可と各シートの作成
function setup() {
  sheet_(); boards_(); notes_();
}
