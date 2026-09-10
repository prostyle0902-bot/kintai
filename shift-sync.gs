/**
 * シフト作成ツールの共有用 Google Apps Script
 *
 * これを新しい Apps Script プロジェクトに貼り付けて「ウェブアプリ」として
 * デプロイすると、shift.html の設定・希望休・作成したシフトを
 * スプレッドシートに保存して、全員で同じ内容を見られるようになります。
 *
 * あわせて「共通スタッフ名簿」も同じスプレッドシートで預かります。
 * 入社登録アプリ（onboarding.html）で登録した人が、
 * PA勤怠（kintai_part）・勤怠（index.html）・シフト作成（shift.html）へ
 * 自動で行きわたるようにするための土台です。
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

/* ===== 最初に1回だけ：権限を許可する =====================================
 * 給与一覧のスプレッドシートを探して開くため、このスクリプトは
 * Google ドライブを見にいきます。そのぶん、これまでより広い許可が必要です。
 *
 * 貼り替えたあと、エディタの上にある関数の選択欄で
 *   authorizeOnce
 * を選んで「実行」を押し、出てくる画面で許可してください。
 * （許可が済むまで、アプリからは「登録できませんでした」と出ます）
 * ======================================================================= */
function authorizeOnce() {
  var lines = [];

  try {
    lines.push('名簿の置き場所：' + SpreadsheetApp.getActiveSpreadsheet().getName());
  } catch (e) {
    lines.push('【要対応】名簿のスプレッドシートを開けません：' + String(e));
  }

  // 給与一覧を開けるか。ここが通れば、給与一覧への登録は動く
  try {
    var t = SpreadsheetApp.openById(PAYROLL_FILES[PAYROLL_FILES.length - 1].id);
    lines.push('給与一覧を開けました：' + t.getName());
  } catch (e) {
    lines.push('【要対応】給与一覧を開けません：' + String(e));
  }

  var msg = lines.join('\n');
  Logger.log(msg);
  return msg;
}

var TOKEN = 'prostyle-shift-2026';   // ★ 自社の合言葉に変えてください
var SHEET_NAME = 'shiftdata';
var CHUNK = 40000;                   // 1セルに入れる文字数（上限5万字より少なめ）

/* 何があっても JSON を返す。
   ここで throw すると Google がエラーのHTMLを返し、アプリ側には
   「Unexpected token '<'」という分かりにくい文言しか届かないため。 */
function safe_(fn, cb) {
  try {
    return fn();
  } catch (err) {
    var body = JSON.stringify({ status: 'error', message: 'スクリプトの中で問題が起きました：' + String(err) });
    if (cb) {
      return ContentService.createTextOutput(cb + '(' + body + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  return safe_(function () {
    var p = {};
    try { p = JSON.parse(e.postData.contents); } catch (err) { p = (e && e.parameter) || {}; }
    return handle_(p, '');
  }, '');
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  return safe_(function () { return handle_(p, p.callback || ''); }, p.callback || '');
}

function handle_(p, cb) {
  try {
    if (String(p.token || '') !== TOKEN) {
      return json_({ status: 'error', message: '合言葉が違います' }, cb);
    }
    /* 生きているかの確認。ブラウザでこのURLを開くだけで確かめられる。
       …/exec?action=ping&token=合言葉
       JSONが出れば動いている。Googleの画面が出るなら、デプロイか許可の問題。 */
    if (p.action === 'ping') {
      return json_({ status: 'ok', message: '動いています', at: nowStamp_(),
                     staffRev: staffRev_(), payrollFiles: PAYROLL_FILES.length }, cb);
    }

    // 版数を見るだけなら、待たされないように鍵を取らない
    if (p.action === 'rev') return json_(revOnly_(getSheet_()), cb);
    // 名簿を読むだけなら、待たされないように鍵を取らない
    if (p.action === 'staffList') return json_(staffList_(), cb);

    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) {
      return json_({ status: 'error', message: '混み合っています。少し待ってからもう一度お試しください' }, cb);
    }
    try {
      var sh = getSheet_();
      if (p.action === 'load') return json_(loadAll_(sh), cb);
      if (p.action === 'save') return json_(saveAll_(sh, p), cb);
      if (p.action === 'staffSave') return json_(staffSave_(p), cb);
      if (p.action === 'staffImport') return json_(staffImport_(p), cb);
      if (p.action === 'staffFill') return json_(staffFill_(p), cb);
      if (p.action === 'staffRetire') return json_(staffRetire_(p), cb);
      if (p.action === 'payrollFiles') return json_(payrollFiles_(), cb);
      if (p.action === 'payrollAdd') return json_(payrollAdd_(p), cb);
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
    staffRev: staffRev_(),
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

/* =========================================================
   共通スタッフ名簿
   入社登録アプリ（onboarding.html）が書き込み、
   PA勤怠・勤怠・シフト作成の3つが読み取る、社内で1つだけの名簿。
   同じスプレッドシートの「staff」シートに1人1行で入る。
   ========================================================= */

var STAFF_SHEET = 'staff';

/* 扱う列の一覧。
   読み書きはシートの見出し行を見て列を探すので、この並びどおりでなくてよい。
   ここに足した列がシートに無ければ、右端に自動で足す。 */
var STAFF_COLS = [
  'id',          // 社内の通し番号。あとから変えない
  'status',      // active（在籍） / retired（退職）
  'kind',        // part（パート・アルバイト） / staff（社員）
  'pin',         // 勤怠アプリの4桁PIN
  'name',        // 氏名（勤怠アプリでの表示名）
  'kana',        // フリガナ
  'dept',        // PA勤怠・勤怠アプリの所属表示
  'storeIds',    // シフト作成の現場ID。掛け持ちはカンマ区切り。空ならシフトには入れない
  'shiftName',   // シフト作成での表記（空なら name と同じ）
  'kyuyoName',   // 給与一覧での表記（空なら name と同じ）
  'payKind',     // hourly（時給） / daily（日給）
  'payRate',     // 単価（円）
  'commute',     // 通勤手当（円）
  'commuteKind', // daily（×出勤日数） / monthly（月固定） / none（なし）
  'locs',        // 打刻できる場所。JSONの配列。空文字なら場所を問わない
  'joinedAt',    // 入社日 YYYY-MM-DD
  'retiredAt',   // 退職日 YYYY-MM-DD
  'kyuyo',       // 出勤簿（給与）への転記対象か
  'days',        // 勤務できる曜日 例 "1,2,3,4,5"
  'holidayOk',   // 祝日に出られるか
  'startTime', 'endTime', 'startTimeWeekend', 'endTimeWeekend',
  'targetDays', 'maxDays', 'maxPerWeek',
  'note',
  'updatedAt', 'updatedBy'
];

// 数字として扱う列（空欄は0にする）
var STAFF_NUM_COLS = ['payRate', 'commute', 'targetDays', 'maxDays', 'maxPerWeek'];
// はい／いいえで扱う列
var STAFF_BOOL_COLS = ['kyuyo', 'holidayOk'];

function getStaffSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(STAFF_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STAFF_SHEET);
    sh.getRange(1, 1, 1, STAFF_COLS.length).setValues([STAFF_COLS]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, STAFF_COLS.length).setFontWeight('bold');
  }
  return sh;
}

/* シートの見出し行を返す。まだ無い列は右端に足す。
   こうしておけば、あとから列を増やしても、すでに入っている行はずれない。 */
function staffHeader_(sh) {
  var width = Math.max(sh.getLastColumn(), 1);
  var head = sh.getRange(1, 1, 1, width).getValues()[0].map(function (v) { return String(v || ''); });
  while (head.length && head[head.length - 1] === '') head.pop();
  if (!head.length) {
    sh.getRange(1, 1, 1, STAFF_COLS.length).setValues([STAFF_COLS]).setFontWeight('bold');
    return STAFF_COLS.slice();
  }
  var add = STAFF_COLS.filter(function (c) { return head.indexOf(c) < 0; });
  if (add.length) {
    sh.getRange(1, head.length + 1, 1, add.length).setValues([add]).setFontWeight('bold');
    head = head.concat(add);
  }
  return head;
}

function staffIndex_(head) {
  var ix = {};
  for (var i = 0; i < head.length; i++) if (head[i]) ix[head[i]] = i;
  return ix;
}

// 名簿が変わった回数。各アプリはこの数字を見て、変わったときだけ読み直す
function staffRev_() {
  return Number(PropertiesService.getScriptProperties().getProperty('staffRev') || 0);
}
function bumpStaffRev_() {
  var n = staffRev_() + 1;
  PropertiesService.getScriptProperties().setProperty('staffRev', String(n));
  return n;
}

// シートの1行を、アプリが扱いやすい形に直す
function staffRowToObj_(row, head) {
  var cols = head || STAFF_COLS;
  var o = {};
  for (var i = 0; i < cols.length; i++) {
    var key = cols[i];
    if (!key) continue;
    var v = row[i];
    if (STAFF_NUM_COLS.indexOf(key) >= 0) {
      o[key] = (v === '' || v === null) ? 0 : Number(v) || 0;
    } else if (STAFF_BOOL_COLS.indexOf(key) >= 0) {
      o[key] = (v === true || String(v).toLowerCase() === 'true' || String(v) === '1');
    } else {
      o[key] = String(v === null || v === undefined ? '' : v);
    }
  }
  return o;
}

function staffObjToRow_(o, head) {
  var cols = head || STAFF_COLS;
  var row = [];
  for (var i = 0; i < cols.length; i++) {
    var key = cols[i];
    if (!key) { row.push(''); continue; }
    var v = o[key];
    if (STAFF_NUM_COLS.indexOf(key) >= 0) {
      row.push((v === '' || v === null || v === undefined) ? '' : Number(v) || 0);
    } else if (STAFF_BOOL_COLS.indexOf(key) >= 0) {
      row.push(v === true || String(v).toLowerCase() === 'true' || String(v) === '1');
    } else {
      row.push(v === null || v === undefined ? '' : String(v));
    }
  }
  return row;
}

function readStaff_(sh, head) {
  head = head || staffHeader_(sh);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, head.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var o = staffRowToObj_(vals[i], head);
    if (!o.id && !o.name) continue;   // 空行は飛ばす
    o._row = i + 2;                   // 何行目にいるか（書き戻しに使う）
    out.push(o);
  }
  return out;
}

function staffList_() {
  var sh = getStaffSheet_();
  var list = readStaff_(sh, staffHeader_(sh));
  for (var i = 0; i < list.length; i++) delete list[i]._row;
  return { status: 'ok', staffRev: staffRev_(), staff: list };
}

function nowStamp_() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
}

// 次の通し番号を作る（p_0001 のような形）
function nextStaffId_(list) {
  var max = 0;
  for (var i = 0; i < list.length; i++) {
    var m = /^p_(\d+)$/.exec(String(list[i].id || ''));
    if (m) max = Math.max(max, Number(m[1]));
  }
  return 'p_' + ('0000' + (max + 1)).slice(-4);
}

// 誰とも重ならない4桁のPINを作る。分かりやすすぎる番号は避ける
function nextPin_(list, extraUsed) {
  var used = {};
  for (var i = 0; i < list.length; i++) if (list[i].pin) used[String(list[i].pin)] = true;
  var ex = extraUsed || [];
  for (var j = 0; j < ex.length; j++) if (ex[j]) used[String(ex[j])] = true;
  var ng = { '1111': 1, '2222': 1, '3333': 1, '4444': 1, '5555': 1, '6666': 1,
             '7777': 1, '8888': 1, '9999': 1, '0000': 1, '1234': 1, '4321': 1 };
  for (var t = 0; t < 5000; t++) {
    var pin = String(Math.floor(1000 + Math.random() * 9000));
    if (!used[pin] && !ng[pin]) return pin;
  }
  return '';
}

// 1人ぶんを登録・更新する
function staffSave_(p) {
  var rec = p.rec;
  if (typeof rec === 'string') { try { rec = JSON.parse(rec); } catch (e) { rec = null; } }
  if (!rec || !String(rec.name || '').trim()) {
    return { status: 'error', message: '氏名が入っていません' };
  }

  var sh = getStaffSheet_();
  var head = staffHeader_(sh);
  var list = readStaff_(sh, head);
  var target = null;
  if (rec.id) {
    for (var i = 0; i < list.length; i++) if (list[i].id === rec.id) { target = list[i]; break; }
    if (!target) return { status: 'error', message: 'その人は名簿にいません（id: ' + rec.id + '）' };
  }

  // 同じ名前の人がすでに在籍していないか（新規登録のときだけ見る）
  if (!target) {
    for (var k = 0; k < list.length; k++) {
      if (list[k].status !== 'retired' && list[k].name === String(rec.name).trim()) {
        return { status: 'error', message: String(rec.name) + ' さんはすでに登録されています' };
      }
    }
  }

  // PIN。指定がなければ自動で採番する
  var pin = String(rec.pin || '').trim();
  if (pin) {
    if (!/^\d{4}$/.test(pin)) return { status: 'error', message: 'PINは4桁の数字で入れてください' };
    for (var m = 0; m < list.length; m++) {
      if (list[m].pin === pin && (!target || list[m].id !== target.id)) {
        return { status: 'error', message: 'PIN ' + pin + ' は ' + list[m].name + ' さんが使っています' };
      }
    }
  } else {
    pin = target ? target.pin : '';
    if (!pin) pin = nextPin_(list, rec.reservedPins || []);
    if (!pin) return { status: 'error', message: 'PINを作れませんでした' };
  }

  var out = {};
  for (var c = 0; c < STAFF_COLS.length; c++) {
    var key = STAFF_COLS[c];
    if (rec[key] !== undefined) out[key] = rec[key];
    else if (target && target[key] !== undefined) out[key] = target[key];
    else out[key] = '';
  }
  out.id = target ? target.id : (rec.id || nextStaffId_(list));
  out.pin = pin;
  out.name = String(rec.name).trim();
  out.status = (out.status === 'retired') ? 'retired' : 'active';
  out.kind = (out.kind === 'staff') ? 'staff' : 'part';
  out.updatedAt = nowStamp_();
  out.updatedBy = String(p.by || '');

  if (target) {
    sh.getRange(target._row, 1, 1, head.length).setValues([staffObjToRow_(out, head)]);
  } else {
    sh.getRange(sh.getLastRow() + 1, 1, 1, head.length).setValues([staffObjToRow_(out, head)]);
  }
  return { status: 'ok', staffRev: bumpStaffRev_(), rec: out, created: !target };
}

// まとめて登録する（いまの名簿の取り込みに使う）。すでにいる人は上書きしない
function staffImport_(p) {
  var recs = p.recs;
  if (typeof recs === 'string') { try { recs = JSON.parse(recs); } catch (e) { recs = null; } }
  if (!recs || !recs.length) return { status: 'error', message: '取り込む内容がありません' };

  var sh = getStaffSheet_();
  var head = staffHeader_(sh);
  var list = readStaff_(sh, head);
  var byName = {}, byPin = {};
  for (var i = 0; i < list.length; i++) {
    byName[list[i].name] = list[i];
    if (list[i].pin) byPin[list[i].pin] = list[i];
  }

  var added = [], skipped = [], rows = [];
  for (var j = 0; j < recs.length; j++) {
    var rec = recs[j] || {};
    var name = String(rec.name || '').trim();
    if (!name) continue;
    if (byName[name]) { skipped.push(name); continue; }

    var made = rows.map(function (r) { return staffRowToObj_(r, head); });
    var pin = String(rec.pin || '').trim();
    if (!/^\d{4}$/.test(pin) || byPin[pin]) pin = nextPin_(list.concat(made), []);

    var out = {};
    for (var c = 0; c < STAFF_COLS.length; c++) {
      var key = STAFF_COLS[c];
      out[key] = (rec[key] !== undefined) ? rec[key] : '';
    }
    out.id = nextStaffId_(list.concat(made));
    out.pin = pin;
    out.name = name;
    out.status = (out.status === 'retired') ? 'retired' : 'active';
    out.kind = (out.kind === 'staff') ? 'staff' : 'part';
    out.updatedAt = nowStamp_();
    out.updatedBy = String(p.by || '');

    byName[name] = out;
    byPin[pin] = out;
    rows.push(staffObjToRow_(out, head));
    added.push(name);
  }

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, head.length).setValues(rows);
    bumpStaffRev_();
  }
  return { status: 'ok', staffRev: staffRev_(), added: added, skipped: skipped };
}

/* すでに名簿にいる人へ、空いている欄だけを埋める。
   取り込みのあとに項目を増やしたときに使う（単価など）。
   すでに何か入っている欄は、こちらの値があっても触らない。 */
function staffFill_(p) {
  var recs = p.recs;
  if (typeof recs === 'string') { try { recs = JSON.parse(recs); } catch (e) { recs = null; } }
  if (!recs || !recs.length) return { status: 'error', message: '入れる内容がありません' };

  // 触ってよい欄。氏名・PIN・状態など、間違えると困るものは対象にしない
  var FILLABLE = ['kana', 'kyuyoName', 'payKind', 'payRate', 'commute', 'commuteKind',
                  'shiftName', 'joinedAt', 'storeIds', 'dept'];

  var sh = getStaffSheet_();
  var head = staffHeader_(sh);
  var list = readStaff_(sh, head);
  var byName = {};
  for (var i = 0; i < list.length; i++) byName[list[i].name] = list[i];

  var filled = [], untouched = [], missing = [];
  for (var j = 0; j < recs.length; j++) {
    var rec = recs[j] || {};
    var name = String(rec.name || '').trim();
    if (!name) continue;
    var t = byName[name];
    if (!t) { missing.push(name); continue; }

    var changed = [];
    for (var k = 0; k < FILLABLE.length; k++) {
      var key = FILLABLE[k];
      if (rec[key] === undefined || rec[key] === '' || rec[key] === null) continue;
      var cur = t[key];
      var empty = (cur === '' || cur === null || cur === undefined || cur === 0);
      if (!empty) continue;              // すでに入っている欄は触らない
      t[key] = rec[key];
      changed.push(key);
    }
    if (!changed.length) { untouched.push(name); continue; }
    t.updatedAt = nowStamp_();
    t.updatedBy = String(p.by || '');
    sh.getRange(t._row, 1, 1, head.length).setValues([staffObjToRow_(t, head)]);
    filled.push(name);
  }

  if (filled.length) bumpStaffRev_();
  return { status: 'ok', staffRev: staffRev_(),
           filled: filled, untouched: untouched, missing: missing };
}

// 退職にする。行は消さずに status を retired にして履歴を残す
function staffRetire_(p) {
  var sh = getStaffSheet_();
  var head = staffHeader_(sh);
  var list = readStaff_(sh, head);
  var target = null;
  for (var i = 0; i < list.length; i++) if (list[i].id === String(p.id || '')) { target = list[i]; break; }
  if (!target) return { status: 'error', message: 'その人は名簿にいません' };

  target.status = 'retired';
  target.retiredAt = String(p.retiredAt || '') || Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
  target.updatedAt = nowStamp_();
  target.updatedBy = String(p.by || '');
  sh.getRange(target._row, 1, 1, head.length).setValues([staffObjToRow_(target, head)]);
  return { status: 'ok', staffRev: bumpStaffRev_(), rec: target };
}

/* =========================================================
   給与一覧への登録
   新しく入った人の「出勤簿」シートを作り、「給与一覧表」にも行を足す。

   計算式はこちらでは組み立てない。同じ現場の人のシートと行をまるごと複製して、
   氏名・単価・通勤手当だけを差し替える。こうすれば給与の計算式が変わっても
   このスクリプトを直さずに済み、式を取り違える心配もない。

   すでに同じ名前のシートがある期間には何もしない（上書きしない）。
   ========================================================= */

/* 給与一覧のスプレッドシート。
   ドライブを探しに行くと、そのぶん広い許可が必要になって動かなくなるため、
   場所をここに直接書いている。
   期間のファイルを新しく作ったら、ここに1行足してください。
   IDは、そのスプレッドシートのURLの /d/ と /edit の間の文字列です。 */
var PAYROLL_FILES = [
  { title: '給与一覧_2026_5_16-2026_6_15', id: '1iRes5ZY-EMu7dAeJnkYdl5KBoNPb8OI0cwkzg37pb7o' },
  { title: '給与一覧_2026_6_16-2026_7_15', id: '1yIkbj5iHAQ6VKWp0LtTcvF_HugELHWhy_fwbON1yR6M' },
  { title: '給与一覧_2026_7_16-2026_8_15', id: '1-v2ZJE46vG3GKTrwSUnJtduT0O8jKFyc9-YsybRF468' },
  { title: '給与一覧_2026_8_16-2026_9_15', id: '1kknpC_KPyaHrldOpXTxC3rQke_sbRXRgYGq2IXlayAw' },
  { title: '給与一覧_2026_9_16-2026_10_15', id: '1Dxr27E2tGwiPKESxWWB8HCGDaX0i2jwtSqfL5Lyc9tw' },
  { title: '給与一覧_2026_10_16-2026_11_15', id: '1SSRb1I2RD6cOkRmBqW0tcZBVIntf_4RHDzDpzqXHi4w' },
  { title: '給与一覧_2026_11_16-2026_12_15', id: '1ARP6ApVaj7VzSKBuz0P5TcdnbH8BQJ5J5IzB_AQgmBw' },
  { title: '給与一覧_2026_12_16-2027_1_15', id: '1MwAXLQfBF-tBQ0jMHyJqLaNlUSbImikp23IdvD9CDBs' },
  { title: '給与一覧_2027_1_16-2027_2_15', id: '1awRXj9U0C18D_gT0tM5u14AENE-zak4d6pafseX5gAM' },
  { title: '給与一覧_2027_2_16-2027_3_15', id: '1rju65K08_cefudwVkKcfw4gJh4uPAWc8kDgpqmeFYC4' },
  { title: '給与一覧_2027_3_16-2027_4_15', id: '18bugMUJZtAQucKRLLA8CAawEiw7IrndNGHhDBIfOpbs' },
  { title: '給与一覧_2027_4_16-2027_5_15', id: '1QUPZqD496tJqtLUh04su1FmmzD4Q8GmcN0yTS9YhsD0' },
  { title: '給与一覧_2027_5_16-2027_6_15', id: '1S5Oq5bHtxBM2JaQApzfkOk2BPu9_GN04MjVCGM5XDn8' },
  { title: '給与一覧_2027_6_16-2027_7_15', id: '1K2q7smAXcqW0VF12FbSmo4vWSjOEwewYgAKfzX6dzWU' },
  { title: '給与一覧_2027_7_16-2027_8_15', id: '1MrtFtRVKxG97mIHiRkX27fkVXxrQNYy9Y5-Q__oZE8Y' },
  { title: '給与一覧_2027_8_16-2027_9_15', id: '1TrgIpopQedhDhzCyx4I0-uqg38ravAP0lw1HnBd3Y6o' },
];

var PAYROLL_PREFIX = '給与一覧_';
var PAYROLL_LIST_SHEET = '給与一覧表';   // まとめの表。名前が違うときは先頭の候補を探す

// 給与一覧_2026_9_16-2026_10_15 → 期間の開始日 2026-09-16
function payrollPeriod_(title) {
  var m = /^給与一覧_(\d{4})_(\d{1,2})_(\d{1,2})-(\d{4})_(\d{1,2})_(\d{1,2})$/.exec(title);
  if (!m) return null;
  var pad = function (n) { return ('0' + n).slice(-2); };
  return {
    start: m[1] + '-' + pad(m[2]) + '-' + pad(m[3]),
    end: m[4] + '-' + pad(m[5]) + '-' + pad(m[6]),
    label: m[1] + '/' + Number(m[2]) + '/' + Number(m[3]) + '〜' + m[4] + '/' + Number(m[5]) + '/' + Number(m[6])
  };
}

/* 給与一覧のファイルを新しい順に並べて返す。
   うまくいかなかったことは errors に入れて返す。
   黙って0件を返すと、何が起きているのか分からなくなるため。 */
function payrollFiles_() {
  var out = [], seen = {}, errors = [];

  for (var k = 0; k < PAYROLL_FILES.length; k++) {
    var f0 = PAYROLL_FILES[k];
    var pr0 = payrollPeriod_(f0.title);
    if (!pr0 || seen[f0.id]) continue;
    seen[f0.id] = true;
    out.push({ id: f0.id, title: f0.title, start: pr0.start, end: pr0.end, label: pr0.label });
  }

  out.sort(function (a, b) { return a.start < b.start ? 1 : a.start > b.start ? -1 : 0; });
  if (!out.length) errors.push('PAYROLL_FILES に給与一覧のファイルが1つも書かれていません。');
  return { status: 'ok', files: out, errors: errors };
}

// まとめの表のシートを見つける
function payrollListSheet_(ss) {
  var sh = ss.getSheetByName(PAYROLL_LIST_SHEET);
  if (sh) return sh;
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    if (all[i].getName().indexOf('給与一覧') >= 0) return all[i];
  }
  return null;
}

/* 出勤簿シートを名前で探す。
   シート名は氏名だけにするが、前に「出勤簿　氏名」で作ったぶんも拾えるようにしておく。 */
function payrollStaffSheet_(ss, name) {
  var all = ss.getSheets();
  for (var i = 0; i < all.length; i++) {
    var n = all[i].getName();
    if (n === name || n === '出勤簿　' + name || n === '出勤簿 ' + name) return all[i];
  }
  return null;
}

// まとめの表から、その現場の見出し行と、そこに並ぶ人の行を調べる
function payrollGroup_(list, group) {
  var vals = list.getRange(1, 1, list.getLastRow(), Math.min(8, list.getLastColumn())).getValues();
  var head = -1, last = -1, sample = -1;
  for (var r = 0; r < vals.length; r++) {
    var joined = vals[r].join(' ');
    if (head < 0) {
      if (joined.indexOf('▶') >= 0 && joined.indexOf(group) >= 0) head = r + 1;
      continue;
    }
    // 見出しのあとに続く、番号の入った行がその現場の人
    if (joined.indexOf('▶') >= 0) break;              // 次の現場に入った
    var no = String(vals[r][0] || '').trim();
    if (/^\d+$/.test(no)) { last = r + 1; if (sample < 0) sample = r + 1; }
  }
  if (head < 0) return null;
  return { head: head, last: last > 0 ? last : head, sample: sample };
}

function payrollAdd_(p) {
  var rec = p.rec;
  if (typeof rec === 'string') { try { rec = JSON.parse(rec); } catch (e) { rec = null; } }
  if (!rec || !String(rec.name || '').trim()) return { status: 'error', message: '氏名が入っていません' };
  var ids = p.fileIds;
  if (typeof ids === 'string') { try { ids = JSON.parse(ids); } catch (e) { ids = null; } }
  if (!ids || !ids.length) return { status: 'error', message: '登録する期間が選ばれていません' };

  var name = String(rec.name).trim();
  var group = String(rec.group || '').trim();
  if (!group) return { status: 'error', message: '現場が指定されていません' };

  var done = [], skipped = [], failed = [];

  for (var i = 0; i < ids.length; i++) {
    try {
      var ss = SpreadsheetApp.openById(String(ids[i]));
      if (payrollStaffSheet_(ss, name)) { skipped.push(ss.getName() + '（すでにあります）'); continue; }

      var list = payrollListSheet_(ss);
      if (!list) { failed.push(ss.getName() + '：給与一覧表が見つかりません'); continue; }

      var g = payrollGroup_(list, group);
      if (!g || g.sample < 0) {
        failed.push(ss.getName() + '：' + group + ' に見本になる人がいないため作れません');
        continue;
      }

      // 同じ現場の人のシートを雛形にする
      var sampleName = String(list.getRange(g.last, 2).getValue() || '').trim();
      var tpl = payrollStaffSheet_(ss, sampleName);
      if (!tpl) { failed.push(ss.getName() + '：' + sampleName + ' の出勤簿が見つかりません'); continue; }

      // --- 出勤簿シートを作る ---
      var sheet = tpl.copyTo(ss).setName(name);
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(tpl.getIndex() + 1);
      payrollFillSheet_(sheet, sampleName, name, rec);

      // --- まとめの表に行を足す ---
      list.insertRowsAfter(g.last, 1);
      list.getRange(g.last, 1, 1, list.getLastColumn()).copyTo(list.getRange(g.last + 1, 1));
      list.getRange(g.last + 1, 2).setValue(name);
      payrollRenumber_(list);

      done.push(ss.getName());
    } catch (e) {
      failed.push(String(ids[i]) + '：' + String(e));
    }
  }

  return { status: 'ok', done: done, skipped: skipped, failed: failed };
}

// 複製したシートの、氏名・単価・通勤手当を差し替えて、打刻の中身を消す
function payrollFillSheet_(sheet, fromName, toName, rec) {
  var rows = Math.min(sheet.getLastRow(), 60);
  var cols = Math.min(sheet.getLastColumn(), 15);
  var rng = sheet.getRange(1, 1, rows, cols);
  var vals = rng.getValues();
  var formulas = rng.getFormulas();

  var payKind = (rec.payKind === 'daily') ? '日給' : '時給';
  var rate = Number(rec.payRate) || 0;
  var commute = Number(rec.commute) || 0;
  var ck = rec.commuteKind;
  var commuteTxt = (ck === 'none' || !commute) ? '通勤手当：なし'
    : (ck === 'monthly') ? '通勤手当：' + commute + '円（月固定）'
    : '通勤手当：' + commute + '円×出勤日数';

  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < cols; c++) {
      if (formulas[r][c]) continue;               // 計算式はそのまま残す
      var v = vals[r][c];
      if (typeof v === 'string' && v) {
        if (v.indexOf(fromName) >= 0) {
          sheet.getRange(r + 1, c + 1).setValue(v.split(fromName).join(toName));
        } else if (/^(時給|日給)：/.test(v)) {
          sheet.getRange(r + 1, c + 1).setValue(payKind + '：' + rate + '円');
        } else if (v.indexOf('通勤手当：') === 0) {
          sheet.getRange(r + 1, c + 1).setValue(commuteTxt + v.replace(/^通勤手当：[^　]*/, ''));
        }
      }
    }
  }

  // 単価と手当の数値セル（見出しの右にある）を入れ替える
  payrollSetNumberAfter_(sheet, vals, formulas, rows, cols, /^(時給|日給)：/, rate);
  payrollSetNumberAfter_(sheet, vals, formulas, rows, cols, /^通勤手当：/, commute);

  // 前の人の出勤・退勤・備考を消す（計算式の入ったセルには触らない）
  payrollClearPunches_(sheet);
}

// 「時給：1094円」のような見出しと同じ行にある数値セルを書き換える
function payrollSetNumberAfter_(sheet, vals, formulas, rows, cols, re, value) {
  for (var r = 0; r < rows; r++) {
    var found = -1;
    for (var c = 0; c < cols; c++) {
      if (typeof vals[r][c] === 'string' && re.test(vals[r][c])) { found = c; break; }
    }
    if (found < 0) continue;
    for (var c2 = found + 1; c2 < cols; c2++) {
      if (formulas[r][c2]) continue;
      if (typeof vals[r][c2] === 'number') { sheet.getRange(r + 1, c2 + 1).setValue(value); return; }
    }
  }
}

// 日付の表の「出勤・退勤・備考」を空にする
function payrollClearPunches_(sheet) {
  var rows = sheet.getLastRow();
  var cols = sheet.getLastColumn();
  if (rows < 2 || cols < 2) return;
  var vals = sheet.getRange(1, 1, rows, cols).getValues();

  // 見出し行（日付・曜・出勤・退勤…）を探す
  var hr = -1, cIn = -1, cOut = -1, cMemo = -1;
  for (var r = 0; r < rows && hr < 0; r++) {
    for (var c = 0; c < cols; c++) {
      if (String(vals[r][c]).trim() === '出勤') {
        hr = r; cIn = c;
        for (var c2 = c; c2 < cols; c2++) {
          var t = String(vals[r][c2]).trim();
          if (t === '退勤') cOut = c2;
          if (t === '備考') cMemo = c2;
        }
        break;
      }
    }
  }
  if (hr < 0) return;

  var formulas = sheet.getRange(1, 1, rows, cols).getFormulas();
  for (var r2 = hr + 1; r2 < rows; r2++) {
    [cIn, cOut, cMemo].forEach(function (c3) {
      if (c3 < 0) return;
      if (formulas[r2][c3]) return;          // 計算式は残す
      sheet.getRange(r2 + 1, c3 + 1).clearContent();
    });
  }
}

// まとめの表の No を振り直す
function payrollRenumber_(list) {
  var rows = list.getLastRow();
  var vals = list.getRange(1, 1, rows, 2).getValues();
  var n = 0;
  for (var r = 0; r < rows; r++) {
    var no = String(vals[r][0] || '').trim();
    var nm = String(vals[r][1] || '').trim();
    if (/^\d+$/.test(no) && nm) { n++; if (Number(no) !== n) list.getRange(r + 1, 1).setValue(n); }
  }
}

/* ===== 一度だけ：シート名を直す =========================================
 * 「出勤簿　氏名」で作ってしまったシートを「氏名」に直します。
 * 他のシートと名前の付け方が揃い、タブが読みやすくなります。
 *
 * エディタの関数の選択欄で fixSheetNames を選んで「実行」を押すだけです。
 * すでに同じ氏名のシートがある場合は、間違えないようそのままにします。
 * ===================================================================== */
function fixSheetNames() {
  var lines = [];
  var total = 0;

  for (var i = 0; i < PAYROLL_FILES.length; i++) {
    var f = PAYROLL_FILES[i];
    try {
      var ss = SpreadsheetApp.openById(f.id);
      var sheets = ss.getSheets();
      var names = {};
      for (var n = 0; n < sheets.length; n++) names[sheets[n].getName()] = true;

      var done = [];
      for (var j = 0; j < sheets.length; j++) {
        var cur = sheets[j].getName();
        if (cur.indexOf('出勤簿　') !== 0 && cur.indexOf('出勤簿 ') !== 0) continue;
        var to = cur.replace(/^出勤簿[　 ]/, '').trim();
        if (!to || names[to]) continue;      // 同じ氏名のシートがあるときは触らない
        sheets[j].setName(to);
        names[to] = true;
        done.push(cur + ' → ' + to);
        total++;
      }
      if (done.length) lines.push(f.title + '：' + done.join('、'));
    } catch (e) {
      lines.push(f.title + '：開けませんでした（' + String(e) + '）');
    }
  }

  var msg = total ? (total + ' 件のシート名を直しました。\n' + lines.join('\n'))
                  : '直すシートはありませんでした。';
  Logger.log(msg);
  return msg;
}
