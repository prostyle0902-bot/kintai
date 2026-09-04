/**
 * シフト作成ツールの共有用 Google Apps Script
 *
 * これを新しい Apps Script プロジェクトに貼り付けて「ウェブアプリ」として
 * デプロイすると、shift.html の設定・希望休・作成したシフトを
 * スプレッドシートに保存して、全員で同じ内容を見られるようになります。
 *
 * --- 準備 ---------------------------------------------------------------
 * 1. 保存用のスプレッドシートを新規作成する（名前は何でもよい）
 * 2. そのスプレッドシートで 拡張機能 → Apps Script を開く
 * 3. 開いたエディタの中身をすべて消して、このファイルの内容を貼り付ける
 * 4. 下の TOKEN を、自社だけが分かる合言葉に書き換える
 *    （shift.html 側の SYNC_TOKEN と同じ文字列にすること）
 * 5. 右上の「デプロイ」→「新しいデプロイ」
 *      種類 : ウェブアプリ
 *      次のユーザーとして実行 : 自分
 *      アクセスできるユーザー : 全員
 *    → 発行された「ウェブアプリのURL」を控える
 * 6. そのURLを shift.html の SYNC_API に設定する
 *
 * --- 注意 ---------------------------------------------------------------
 * 「アクセスできるユーザー: 全員」で公開するため、URLと合言葉を知っていれば
 * 誰でも読み書きできます。URLは社内だけで共有してください。
 * 合言葉を変えたいときは、この TOKEN と shift.html の SYNC_TOKEN の
 * 両方を同じ値に書き換えます。
 */

var TOKEN = 'prostyle-shift-2026';   // ★ 自社の合言葉に変えてください
var SHEET_NAME = 'shiftdata';
var CHUNK = 40000;                   // 1セルに入れる文字数（上限5万字より少なめ）

function doPost(e) {
  var p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) { p = e.parameter || {}; }
  return handle_(p, '');
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  return handle_(p, p.callback || '');
}

function handle_(p, cb) {
  try {
    if (String(p.token || '') !== TOKEN) {
      return json_({ status: 'error', message: '合言葉が違います' }, cb);
    }
    // 版数を見るだけなら、待たされないように鍵を取らない
    if (p.action === 'rev') return json_(revOnly_(getSheet_()), cb);

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) {
      return json_({ status: 'error', message: '混み合っています。少し待ってからもう一度お試しください' }, cb);
    }
    try {
      var sh = getSheet_();
      if (p.action === 'load') return json_(loadAll_(sh), cb);
      if (p.action === 'save') return json_(saveAll_(sh, p), cb);
      return json_({ status: 'error', message: '不明な操作です: ' + p.action }, cb);
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ status: 'error', message: String(err) }, cb);
  }
}

// callback が付いていれば JSONP で返す（ブラウザから読むときに使う）
function json_(obj, cb) {
  var body = JSON.stringify(obj);
  if (cb) {
    return ContentService.createTextOutput(cb + '(' + body + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

// 誰かが更新したかを見るだけの、軽い問い合わせ
function revOnly_(sh) {
  return {
    status: 'ok',
    rev: Number(sh.getRange('B1').getValue() || 0),
    updatedAt: String(sh.getRange('B2').getValue() || ''),
    updatedBy: String(sh.getRange('B3').getValue() || ''),
  };
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange('A1').setValue('rev');
    sh.getRange('A2').setValue('updatedAt');
    sh.getRange('A3').setValue('updatedBy');
    sh.getRange('A4').setValue('--- data ---');
    sh.getRange('B1').setValue(0);
    sh.setColumnWidth(2, 120);
  }
  return sh;
}

function loadAll_(sh) {
  var rev = Number(sh.getRange('B1').getValue() || 0);
  var last = sh.getLastRow();
  var data = '';
  if (last >= 5) {
    var vals = sh.getRange(5, 1, last - 4, 1).getValues();
    for (var i = 0; i < vals.length; i++) data += String(vals[i][0] || '');
  }
  return {
    status: 'ok',
    rev: rev,
    updatedAt: String(sh.getRange('B2').getValue() || ''),
    updatedBy: String(sh.getRange('B3').getValue() || ''),
    data: data,
  };
}

function saveAll_(sh, p) {
  var cur = Number(sh.getRange('B1').getValue() || 0);
  var base = Number(p.baseRev || 0);
  // 自分が読み込んだあとに他の人が保存していたら、上書きせずに知らせる
  if (base !== cur && String(p.force) !== 'true') {
    var info = loadAll_(sh);
    info.status = 'conflict';
    return info;
  }

  var data = String(p.data || '');
  var chunks = [];
  for (var i = 0; i < data.length; i += CHUNK) chunks.push([data.substr(i, CHUNK)]);
  if (!chunks.length) chunks = [['']];

  var last = sh.getLastRow();
  if (last >= 5) sh.getRange(5, 1, last - 4, 1).clearContent();
  sh.getRange(5, 1, chunks.length, 1).setValues(chunks);

  var rev = cur + 1;
  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
  sh.getRange('B1').setValue(rev);
  sh.getRange('B2').setValue(now);
  sh.getRange('B3').setValue(String(p.by || ''));

  return { status: 'ok', rev: rev, updatedAt: now, updatedBy: String(p.by || '') };
}
