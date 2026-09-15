/* =============================================================
   出勤簿GAS：入社登録アプリの名簿と、ひとつなぎにするパッチ

   これを入れると、新しい人が入ったときに出勤簿GASを手で直す必要が
   なくなる。入社登録アプリで登録するだけでよくなる。

   これまで手作業だったこと        →  このパッチが自動でやること
   ------------------------------------------------------------
   STAFF_MAP に名前を書き足す      →  名簿を見るので不要
   出勤簿シートを作る              →  同じ現場の人を雛形に自動で作る
   STORE_STAFF に書き足す          →  名簿の配属先から自動で足す
   時給・通勤手当を入れる          →  名簿の値をそのまま入れる

   ―――― 入れ方 ――――
   1. 出勤簿GASのいちばん下に、このファイルの中身を貼り付ける
      （前の版を貼ってある場合は、そこを消してから貼る）
   2. 次の1行を書き換える（writeShiftToKyuyo の中・1か所だけ）

        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) return { status: 'skip', reason: 'sheet_not_found' };
        ↓
        var sheet = ss.getSheetByName(sheetName) || ensureStaffSheet_(ss, sheetName, ev);
        if (!sheet) return { status: 'skip', reason: 'sheet_not_found' };

   3. 保存 → デプロイを管理 → 鉛筆 → バージョン「新バージョン」→ デプロイ
   4. checkRoster() を実行して、【要対応】が出ないか見る
   5. 毎日ひとりでに追いつくようにするなら installRosterTrigger() を1回実行

   ※ STAFF_MAP・STORE_STAFF・STAFF_WAGES は消さないこと。
     名簿に無い人（過去の在籍者など）の受け皿として残す。
   ============================================================= */

// 入社登録アプリの名簿（シフト自動作成アプリ用スプレッドシートの staff シート）
var ROSTER_SS_ID     = '1ysD43TtjdX9Cs0NfjsTS1uqeTfAfZJ6sdstgbBJhoXY';
var ROSTER_SHEET     = 'staff';
var ROSTER_CACHE     = 'kyuyo_roster_v2';
var ROSTER_CACHE_SEC = 300;   // 5分。名簿を直したら最大5分で反映される
var SYNC_MAX_BOOKS   = 3;     // 一度に面倒を見る給与一覧の冊数（今の期間から先へ）

/* シフト作成の現場ID → 給与一覧の店舗グループ（STORE_STAFF のキー）。
   名簿は現場IDで持っているので、ここで読み替える。 */
var STORE_BY_ID = {
  s_sawara13:  '十三里屋',
  s_ryuchan:   'りゅうちゃん',
  s_takohai:   'タコハイ',
  s_bussing:   'バッシング',
  s_beer:      'バッシング',
  s_seiso:     '横丁清掃',
  s_yakitate:  '焼きたて屋',
  s_totalstay: 'トータルステイ',
  s_ichikawa:  '市川東病院',
  s_kaiji:     '海事・横河',
  s_chibashoyu:'ちば醬油',
  s_yuya:      'ゆう屋',
  s_jingu:     '神宮・亀甲堂',
  s_katorijingu:'香取神宮',
  s_nissei:    'ニッセーデリカ',
  s_honsha:    '本社'
};


/* =============================================================
   1. 転記先のシート名を決める
   ============================================================= */

/* ①名簿 → ②名簿（表記ゆれを吸収） → ③STAFF_MAP → ④STAFF_MAP（同）の順に探す */
function staffSheetName_(appName) {
  var key = String(appName || '').trim();
  if (!key) return null;

  var roster = rosterMap_();
  if (roster[key]) return roster[key];

  var n = normName_(key);
  for (var rk in roster) if (normName_(rk) === n) return roster[rk];

  if (STAFF_MAP[key]) return STAFF_MAP[key];
  for (var sk in STAFF_MAP) if (normName_(sk) === n) return STAFF_MAP[sk];

  return null;
}


/* =============================================================
   2. 名簿を読む
   ============================================================= */

/* 在籍していて「出勤簿へ転記する」が入っている人を、そのまま返す */
function rosterRecords_() {
  var out = [];
  try {
    var sh = SpreadsheetApp.openById(ROSTER_SS_ID).getSheetByName(ROSTER_SHEET);
    if (!sh || sh.getLastRow() < 2) return out;

    var values = sh.getDataRange().getValues();
    var head = values[0];
    var col = {};
    for (var c = 0; c < head.length; c++) col[String(head[c]).trim()] = c;
    if (col.name === undefined || col.kyuyo === undefined) return out;

    var get = function (row, key) {
      return col[key] === undefined ? '' : row[col[key]];
    };
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var name = String(get(row, 'name')).trim();
      if (!name) continue;
      if (String(get(row, 'status')).trim() !== 'active') continue;
      if (!isTrue_(get(row, 'kyuyo'))) continue;

      out.push({
        name:        name,
        shiftName:   String(get(row, 'shiftName')).trim(),
        kyuyoName:   String(get(row, 'kyuyoName')).trim(),
        storeIds:    String(get(row, 'storeIds')).trim(),
        dept:        String(get(row, 'dept')).trim(),
        payKind:     String(get(row, 'payKind')).trim(),
        payRate:     Number(get(row, 'payRate')) || 0,
        commute:     Number(get(row, 'commute')) || 0,
        commuteKind: String(get(row, 'commuteKind')).trim() || 'none'
      });
    }
  } catch (e) {
    Logger.log('名簿を読めませんでした: ' + e.message);
  }
  return out;
}

/* 「シフト表での名前 → 出勤簿でのシート名」だけを取り出したもの。
   転記のたびに使うので、5分だけ控えておく。 */
function rosterMap_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(ROSTER_CACHE);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) { /* 壊れていたら読み直す */ }
  }
  var map = {};
  var recs = rosterRecords_();
  for (var i = 0; i < recs.length; i++) {
    map[recs[i].shiftName || recs[i].name] = recs[i].kyuyoName || recs[i].name;
  }
  try { cache.put(ROSTER_CACHE, JSON.stringify(map), ROSTER_CACHE_SEC); } catch (e) {}
  return map;
}

// 出勤簿でのシート名
function sheetNameOf_(rec) { return rec.kyuyoName || rec.name; }

// その人が給与一覧のどの店舗グループに入るか
function storeGroupOf_(rec) {
  var ids = String(rec.storeIds || '').split(',');
  for (var i = 0; i < ids.length; i++) {
    var g = STORE_BY_ID[String(ids[i]).trim()];
    if (g && STORE_STAFF[g]) return g;
  }
  // 現場IDで分からなければ、所属の文字で当てにいく
  var dept = String(rec.dept || '');
  for (var st in STORE_STAFF) {
    if (dept && (dept.indexOf(st) >= 0 || st.indexOf(dept) >= 0)) return st;
  }
  return null;
}

// チェックボックスは true / 'TRUE' / 1 のどれで入っていても拾う
function isTrue_(v) {
  if (v === true) return true;
  var s = String(v).trim().toLowerCase();
  return s === 'true' || s === '1' || s === 'はい' || s === '○';
}

// 空白・異体字の違いを吸収してから比べる
function normName_(s) {
  return String(s || '')
    .replace(/[\s　・･]/g, '')
    .replace(/﨑|嵜/g, '崎').replace(/髙/g, '高').replace(/濵/g, '浜')
    .replace(/眞/g, '真').replace(/齊|齋/g, '斉');
}

// 名簿を直したあと、5分待たずに反映させたいとき
function clearRosterCache() {
  try { CacheService.getScriptCache().remove(ROSTER_CACHE); } catch (e) {}
  Logger.log('名簿の控えを消しました。次の転記で読み直します。');
}


/* =============================================================
   3. 出勤簿シートを自動で作る
   ============================================================= */

/* 転記しようとしてシートが無かったときに呼ばれる。
   同じ現場の人のシートを雛形にするので、その現場ならではの作り
   （トータルステイの土日祝加算など）もそのまま写る。 */
function ensureStaffSheet_(ss, sheetName, ev) {
  var rec = rosterBySheetName_(sheetName);
  if (!rec) return null;   // 名簿にいない人には作らない
  return createStaffSheet_(ss, sheetName, rec, ev && ev.store);
}

function rosterBySheetName_(sheetName) {
  var recs = rosterRecords_();
  var n = normName_(sheetName);
  for (var i = 0; i < recs.length; i++) {
    if (normName_(sheetNameOf_(recs[i])) === n) return recs[i];
  }
  return null;
}

/* 雛形にするシートを選ぶ。
   通勤手当の出し方（日額／月固定／なし）が同じ人を優先する。
   明細の式が雛形ゆずりになるため、ここが違うと金額がずれる。 */
function templateSheetFor_(ss, group, commuteKind) {
  var list = (group && STORE_STAFF[group]) ? STORE_STAFF[group] : [];
  var want = commuteKind === 'monthly' ? 'fixed'
           : commuteKind === 'daily'   ? 'daily' : 'none';

  var best = null, plain = null, any = null;
  for (var i = 0; i < list.length; i++) {
    var sh = ss.getSheetByName(list[i]);
    if (!sh) continue;
    if (!any) any = sh;
    var w = (typeof STAFF_WAGES !== 'undefined') ? STAFF_WAGES[list[i]] : null;
    /* 店長手当・職務手当などが付いている人は雛形にしない。
       明細の行ごと写ってしまい、新しい人にその手当が付いてしまう。 */
    if (w && w.allowances && w.allowances.length) continue;
    if (!plain) plain = sh;
    if (w && w.commuteType === want && !best) best = sh;
  }
  if (best) return best;
  if (plain) {
    Logger.log('※通勤手当の出し方が同じ人が' + group + 'にいないため、' + plain.getName()
               + 'を雛形にします。明細の「通勤手当」の行を確かめてください。');
    return plain;
  }
  if (any) {
    Logger.log('【要対応】' + group + 'は手当が付いている人しかいないため、' + any.getName()
               + 'を雛形にします。その手当の行が写るので、必ず消してください。');
    return any;
  }

  // 同じ現場に誰もいなければ、どこかの出勤簿タブを使う
  var sheets = ss.getSheets();
  for (var j = 0; j < sheets.length; j++) {
    var nm = sheets[j].getName();
    if (nm.indexOf('給与一覧') >= 0 || nm === ID_SHEET_NAME || nm === '_時刻リスト') continue;
    Logger.log('【要対応】' + (group || '現場不明') + 'に雛形がないため、' + nm
               + 'を雛形にします。明細を必ず確かめてください。');
    return sheets[j];
  }
  return null;
}

function createStaffSheet_(ss, sheetName, rec, storeHint) {
  var already = ss.getSheetByName(sheetName);
  if (already) return already;

  var group = storeGroupOf_(rec);
  if (!group && storeHint) {
    for (var st in STORE_STAFF) {
      if (String(storeHint).indexOf(st) >= 0 || st.indexOf(String(storeHint)) >= 0) { group = st; break; }
    }
  }
  var tmpl = templateSheetFor_(ss, group, rec.commuteKind);
  if (!tmpl) { Logger.log('雛形が見つかりません: ' + ss.getName() + ' / ' + sheetName); return null; }

  var sheet = tmpl.copyTo(ss);
  sheet.setName(sheetName);

  var kindLabel = (rec.payKind === 'daily') ? '日給' : '時給';
  sheet.getRange(3, 11).setValue(rec.payRate);                       // K3 単価
  sheet.getRange(3, 13).setValue(rec.commuteKind === 'none' ? 0 : rec.commute);  // M3 通勤手当
  sheet.getRange(3, 2).setValue(kindLabel + '：' + rec.payRate + '円');

  var ctxt = rec.commuteKind === 'daily'   ? '通勤手当：' + rec.commute + '円×出勤日数'
           : rec.commuteKind === 'monthly' ? '通勤手当：' + rec.commute + '円（月固定）'
           : '通勤手当：なし';
  sheet.getRange(3, 6).setValue(ctxt + '　深夜割増：22:00〜翌5:00（法定）');

  // 1行目の見出しの氏名を差し替える
  var title = String(sheet.getRange(1, 2).getValue())
    .replace(/出勤簿　.*?（/, '出勤簿　' + sheetName + '（');
  sheet.getRange(1, 2).setValue(title);

  // 出退勤の中身だけ消す（日付と曜日はその期間のまま使う）
  var vals = sheet.getDataRange().getValues();
  for (var r = 0; r < vals.length; r++) {
    var cell = vals[r][1];
    if (!((cell instanceof Date) || /^\d+\/\d+$/.test(String(cell).trim()))) continue;
    sheet.getRange(r + 1, 4).setValue('');   // 出勤
    sheet.getRange(r + 1, 5).setValue('');   // 退勤
    sheet.getRange(r + 1, 6).setValue(0);    // 休憩
    sheet.getRange(r + 1, 9).setValue('');   // 備考
  }

  ss.setActiveSheet(sheet);
  ss.moveActiveSheet(tmpl.getIndex() + 1);

  Logger.log('シート作成: ' + ss.getName() + ' / ' + sheetName +
             '（' + (group || '現場不明') + '・' + tmpl.getName() + 'を雛形）');
  if (rec.payKind === 'daily') {
    Logger.log('  ※日給の方です。明細の「基本賃金」の式が時給向けのままなので確認してください');
  }
  return sheet;
}


/* =============================================================
   4. 名簿を出勤簿にまるごと反映する（これ1つ実行すればよい）
   ============================================================= */

/* 入社登録アプリで人を登録したあと、これを実行する。
   installRosterTrigger() を入れておけば、1日1回ひとりでに動く。

   ・出勤簿シートが無い人のシートを作る（今の期間から SYNC_MAX_BOOKS 冊ぶん）
   ・給与一覧の集計表に載っていない人を載せる
   何も足りていなければ、何もしない。 */
function syncRoster() {
  var t0 = new Date().getTime();
  clearRosterCache();

  var recs = rosterRecords_();
  if (!recs.length) { Logger.log('【要対応】名簿を1人も読めていません。checkRoster() で確かめてください。'); return; }
  Logger.log('名簿の転記対象：' + recs.length + '人');

  var books = booksFromNow_().slice(0, SYNC_MAX_BOOKS);
  if (!books.length) { Logger.log('見にいける給与一覧がありません'); return; }

  var madeIn = {};   // シートを作った冊。あとで集計表を組み直す
  var made = 0;
  for (var b = 0; b < books.length; b++) {
    if (new Date().getTime() - t0 > 4 * 60 * 1000) {
      Logger.log('時間切れです。もう一度 syncRoster() を実行して続きをやってください。');
      break;
    }
    var ss = SpreadsheetApp.openById(books[b].id);
    for (var i = 0; i < recs.length; i++) {
      var nm = sheetNameOf_(recs[i]);
      if (ss.getSheetByName(nm)) continue;
      if (createStaffSheet_(ss, nm, recs[i], null)) {
        madeIn[books[b].id] = books[b].label;
        made++;
        Utilities.sleep(500);
      }
    }
  }

  // 給与一覧の集計表に、名簿の人を載せる
  var added = mergeRosterIntoStoreStaff_();
  for (var a = 0; a < added.length; a++) Logger.log('集計表に追加: ' + added[a]);

  var ids = Object.keys(madeIn);
  if (added.length || ids.length) {
    // 集計表はどの冊でも組み直す必要がある（並び順と合計式が変わるため）
    for (var k = 0; k < books.length; k++) {
      if (new Date().getTime() - t0 > 5 * 60 * 1000) {
        Logger.log('時間切れです。もう一度 syncRoster() を実行してください。');
        break;
      }
      try {
        fixKyuyoIchiranInSS(books[k].id);
        Logger.log('集計表を組み直しました: ' + books[k].label);
      } catch (e) {
        Logger.log('【要対応】集計表を組み直せませんでした: ' + books[k].label + ' / ' + e.message);
      }
      Utilities.sleep(800);
    }
  }

  Logger.log('■完了：シート ' + made + '枚作成 / 集計表に ' + added.length + '人追加');
  if (!made && !added.length) Logger.log('（足りないものはありませんでした）');
}

/* 名簿の人を STORE_STAFF に足す。
   このスクリプトが動いている間だけの追加で、コードは書き換えない。
   退職した人は消さない（過去の集計表を変えないため）。 */
function mergeRosterIntoStoreStaff_() {
  var recs = rosterRecords_();
  var added = [];
  for (var i = 0; i < recs.length; i++) {
    var nm = sheetNameOf_(recs[i]);

    var already = false;
    for (var st in STORE_STAFF) {
      if (STORE_STAFF[st].indexOf(nm) >= 0) { already = true; break; }
    }
    if (already) continue;

    var g = storeGroupOf_(recs[i]);
    if (!g || !STORE_STAFF[g]) {
      Logger.log('【要対応】' + nm + ' の現場が分からないため、集計表に載せられません' +
                 '（名簿の配属先を確かめてください）');
      continue;
    }
    STORE_STAFF[g].push(nm);
    added.push(nm + ' → ' + g);
  }
  return added;
}

/* 1日1回、ひとりでに syncRoster() を動かす。1回実行すれば入る。 */
function installRosterTrigger() {
  removeRosterTrigger();
  ScriptApp.newTrigger('syncRoster').timeBased().everyDays(1).atHour(5).create();
  Logger.log('毎朝5時ごろに名簿を反映するようにしました。');
  Logger.log('やめるときは removeRosterTrigger() を実行してください。');
}

function removeRosterTrigger() {
  var ts = ScriptApp.getProjectTriggers();
  var n = 0;
  for (var i = 0; i < ts.length; i++) {
    if (ts[i].getHandlerFunction() === 'syncRoster') { ScriptApp.deleteTrigger(ts[i]); n++; }
  }
  if (n) Logger.log('自動実行を ' + n + '件やめました。');
}


/* =============================================================
   5. 点検
   ============================================================= */

/* 名簿が読めているか、誰が転記対象か、食い違いが無いかを確かめる。
   エディタでこの関数を選んで実行し、実行ログを見る。 */
function checkRoster() {
  clearRosterCache();
  var recs = rosterRecords_();
  Logger.log('名簿から読めた転記対象：' + recs.length + '人');
  if (!recs.length) {
    Logger.log('【要対応】名簿を1人も読めていません。');
    Logger.log('  ・ROSTER_SS_ID が合っているか');
    Logger.log('  ・staff シートがあるか');
    Logger.log('  ・この関数を1回実行して、権限の確認を許可したか');
    return;
  }
  for (var i = 0; i < recs.length; i++) {
    var appNm = recs[i].shiftName || recs[i].name;
    var shNm  = sheetNameOf_(recs[i]);
    var mark  = STAFF_MAP[appNm] ? '' : '　← 名簿だけにいる人';
    Logger.log('  ' + appNm + ' → ' + shNm + mark);
  }

  /* 名簿では在籍なのに、GAS側の一覧で「退職」や「対象外」になっている人を探す。
     ここが食い違うと、集計表から外れたり、シートを消されたりする。 */
  Logger.log('―――― GAS側の一覧と食い違っていないか ――――');
  var inStore = {};
  for (var st in STORE_STAFF) {
    for (var si = 0; si < STORE_STAFF[st].length; si++) inStore[STORE_STAFF[st][si]] = st;
  }
  var ng = 0;
  for (var k = 0; k < recs.length; k++) {
    var nm = sheetNameOf_(recs[k]);
    if (typeof RETIRED_STAFF !== 'undefined' && RETIRED_STAFF.indexOf(nm) >= 0) {
      Logger.log('【要対応】' + nm + ' は名簿では在籍ですが、RETIRED_STAFF に入っています');
      Logger.log('    → このまま集計表を組み直すと、載らなくなります');
      ng++;
    }
    if (!inStore[nm]) {
      var g = storeGroupOf_(recs[k]);
      Logger.log((g ? '　' : '【要対応】') + nm + ' は STORE_STAFF にいません' +
                 (g ? '（syncRoster() で ' + g + ' に自動で足されます）'
                    : '（現場が分からないので自動で足せません。名簿の配属先を確かめてください）'));
      if (!g) ng++;
    }
  }
  if (!ng) Logger.log('食い違いはありません。');

  Logger.log('―――― 出勤簿にシートがあるか（今の期間から先を全部） ――――');
  var books = booksFromNow_();
  if (!books.length) { Logger.log('見にいける給与一覧がありませんでした'); return; }
  var missing = 0;
  for (var b = 0; b < books.length; b++) {
    var ss = SpreadsheetApp.openById(books[b].id);
    var lack = [];
    for (var j = 0; j < recs.length; j++) {
      if (!ss.getSheetByName(sheetNameOf_(recs[j]))) lack.push(sheetNameOf_(recs[j]));
    }
    if (lack.length) {
      Logger.log('シートがありません：' + books[b].label + ' → ' + lack.join('、'));
      missing += lack.length;
    }
    Utilities.sleep(300);
  }
  if (!missing) {
    Logger.log('シートは全期間そろっています。');
  } else {
    Logger.log('※ syncRoster() を実行すると、今の期間から ' + SYNC_MAX_BOOKS + '冊ぶんは自動で作られます。');
    Logger.log('　 先の期間は、その人のシフトを転記したときに自動で作られます。');
  }
  Logger.log('確認おわり。【要対応】が出ていなければ大丈夫です。');
}

/* 今の期間から先の給与一覧を、古い順に返す */
function booksFromNow_() {
  var today = new Date();
  var todayNum = today.getFullYear() * 10000 + (today.getMonth() + 1) * 100 + today.getDate();
  var idSheet = getOrCreateIdSheet();
  var out = [];
  for (var i = 0; i < PERIODS.length; i++) {
    var p = PERIODS[i];
    var e = p.end.split('/');
    var eNum = (+e[0]) * 10000 + (+e[1]) * 100 + (+e[2]);
    if (eNum < todayNum) continue;          // もう終わった期間は触らない
    var id = getSSIdForPeriod(idSheet, p.label);
    if (id) out.push({ label: p.label, id: id });
  }
  return out;
}
