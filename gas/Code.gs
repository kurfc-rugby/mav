/**
 * KURFC マルチアングル 短縮リンク
 *  POST  本文 = 共有データ（"z..." の文字列）      → {"k":"Xy7Pq2"}
 *  GET   ?k=Xy7Pq2                                  → {"d":"z..."}
 * データはスプレッドシート「KURFCマルチアングル 短縮リンク」の「links」シートに 1行1件で保存します。
 * 列: A=ID  B=作成日時  C=指紋(MD5)  D以降=データ（長い場合は分割）
 */
const SSID = '1iU_W3YgJmshZ3oKep6ffJiSb303LpdKrev19Wbf4YFs'; // スプレッドシート「KURFCマルチアングル 短縮リンク」
const SHEET = 'links';
const CHUNK = 45000;          // 1セルの上限(5万文字)より少し小さく
const MAX_LEN = 400000;       // 異常に大きいデータは受け付けない
const ID_LEN = 6;
const ABC = 'abcdefghijkmnopqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // 紛らわしい l/1/O/0/I を除く

function sheet_() {
  const ss = SpreadsheetApp.openById(SSID);
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(['ID', '作成日時', '指紋', 'データ']);
    sh.setFrozenRows(1);
  }
  return sh;
}

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
  const r = sh.getRange(2, col, Math.max(sh.getLastRow() - 1, 1), 1)
    .createTextFinder(value).matchEntireCell(true).matchCase(true).findNext();
  return r ? r.getRow() : 0;
}

function doPost(e) {
  const d = String((e && e.postData && e.postData.contents) || '').trim();
  if (!/^[zj][A-Za-z0-9_-]+$/.test(d) || d.length > MAX_LEN) return json_({error: 'bad data'});
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const fp = md5_(d);
    const same = find_(sh, 3, fp);                     // 同じ内容なら同じIDを返す
    if (same) return json_({k: String(sh.getRange(same, 1).getValue())});
    let id;
    do { id = newId_(); } while (find_(sh, 1, id));
    const parts = [];
    for (let i = 0; i < d.length; i += CHUNK) parts.push(d.slice(i, i + CHUNK));
    // 先頭の ' は「文字として保存」の印（数字や式に化けないように）。読み出し時には付きません
    sh.appendRow(["'" + id, new Date(), "'" + fp].concat(parts.map(x => "'" + x)));
    return json_({k: id});
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const id = String((e && e.parameter && e.parameter.k) || '');
  if (!/^[A-Za-z0-9]{4,12}$/.test(id)) return json_({ok: true, app: 'KURFC multi-angle short links'});
  const sh = sheet_();
  const row = find_(sh, 1, id);
  if (!row) return json_({error: 'not found'});
  const vals = sh.getRange(row, 4, 1, Math.max(sh.getLastColumn() - 3, 1)).getValues()[0];
  return json_({d: vals.filter(String).join('')});
}

// 最初に1回だけ実行：権限の許可と「links」シートの作成
function setup() {
  sheet_();
}
