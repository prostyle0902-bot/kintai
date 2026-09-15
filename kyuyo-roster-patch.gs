/* =============================================================
   出勤簿GAS：転記先を「入社登録アプリの名簿」から自動で決める

   これまでは STAFF_MAP という手書きの一覧に載っている人だけが
   転記されていた。新しく入った人は、ここに書き足すまで
   だまって捨てられる（status:'skip' / reason:'not_part_time'）。

   このファイルを足すと、まず名簿を見にいくようになる。
   名簿で「出勤簿（給与）へ転記する」が入っていれば、
   STAFF_MAP に書かなくても転記される。
   名簿で見つからない人は、これまでどおり STAFF_MAP を見る。

   ―――― 入れ方 ――――
   1. このファイルの中身を、出勤簿GASのいちばん下に貼り付ける
   2. 既にある次の2行を、それぞれ書き換える（2か所だけ）

        var sheetName = STAFF_MAP[ev.staff];
        ↓
        var sheetName = staffSheetName_(ev.staff);

      1つは writeShiftToKyuyo の中、もう1つは deleteShiftFromKyuyo の中。
   3. 保存して、デプロイ →「デプロイを管理」→ 鉛筆 → バージョン「新バージョン」→ デプロイ
   4. 一度 checkRoster() を実行して、名簿が読めているか確かめる

   ※ STAFF_MAP は消さないこと。名簿に無い人の受け皿として残す。
   ============================================================= */

// 入社登録アプリの名簿（シフト自動作成アプリ用スプレッドシートの staff シート）
var ROSTER_SS_ID  = '1ysD43TtjdX9Cs0NfjsTS1uqeTfAfZJ6sdstgbBJhoXY';
var ROSTER_SHEET  = 'staff';
var ROSTER_CACHE  = 'kyuyo_roster_v1';
var ROSTER_CACHE_SEC = 300;   // 5分。名簿を直したら最大5分で反映される

/* 転記先のシート名を決める。
   ①名簿 → ②名簿（表記ゆれを吸収） → ③STAFF_MAP → ④STAFF_MAP（同）の順に探す */
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

/* 名簿を読んで「シフト表での名前 → 出勤簿でのシート名」を作る。
   在籍していて、かつ「出勤簿へ転記する」が入っている人だけ。 */
function rosterMap_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(ROSTER_CACHE);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) { /* 壊れていたら読み直す */ }
  }

  var map = {};
  try {
    var sh = SpreadsheetApp.openById(ROSTER_SS_ID).getSheetByName(ROSTER_SHEET);
    if (sh && sh.getLastRow() > 1) {
      var values = sh.getDataRange().getValues();
      var head = values[0];
      var col = {};
      for (var c = 0; c < head.length; c++) col[String(head[c]).trim()] = c;

      // 名簿の欄がそろっているときだけ使う
      if (col.name !== undefined && col.kyuyo !== undefined) {
        for (var r = 1; r < values.length; r++) {
          var row = values[r];
          var name = String(row[col.name] || '').trim();
          if (!name) continue;
          if (col.status !== undefined && String(row[col.status]).trim() !== 'active') continue;
          if (!isTrue_(row[col.kyuyo])) continue;

          var shiftName = col.shiftName !== undefined
            ? String(row[col.shiftName] || '').trim() : '';
          var kyuyoName = col.kyuyoName !== undefined
            ? String(row[col.kyuyoName] || '').trim() : '';

          map[shiftName || name] = kyuyoName || name;
        }
      }
    }
  } catch (e) {
    // 名簿が読めないときは空で返す。STAFF_MAP があるので転記は止まらない
    Logger.log('名簿を読めませんでした: ' + e.message);
    return {};
  }

  try { cache.put(ROSTER_CACHE, JSON.stringify(map), ROSTER_CACHE_SEC); } catch (e) {}
  return map;
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
    .replace(/嵜/g, '崎').replace(/髙/g, '高').replace(/濵/g, '浜')
    .replace(/眞/g, '真').replace(/齊|齋/g, '斉');
}

/* 名簿が読めているか、誰が転記対象になっているかを確かめる。
   エディタでこの関数を選んで実行し、実行ログを見る。
   今の期間から先の給与一覧を全部見て、シートが足りない人を挙げる。 */
function checkRoster() {
  clearRosterCache();
  var map = rosterMap_();
  var names = Object.keys(map);
  Logger.log('名簿から読めた転記対象：' + names.length + '人');
  if (!names.length) {
    Logger.log('【要対応】名簿を1人も読めていません。');
    Logger.log('  ・ROSTER_SS_ID が合っているか');
    Logger.log('  ・staff シートがあるか');
    Logger.log('  ・この関数を1回実行して、権限の確認を許可したか');
    return;
  }
  for (var i = 0; i < names.length; i++) {
    var onlyRoster = !STAFF_MAP[names[i]] ? '　← 名簿だけにいる人' : '';
    Logger.log('  ' + names[i] + ' → ' + map[names[i]] + onlyRoster);
  }
  /* 名簿では在籍なのに、GAS側の一覧で「退職」や「対象外」になっている人を探す。
     ここが食い違うと、集計表から外れたり、シートを消されたりする。 */
  Logger.log('―――― GAS側の一覧と食い違っていないか ――――');
  var inStore = {};
  for (var st in STORE_STAFF) {
    for (var si = 0; si < STORE_STAFF[st].length; si++) inStore[STORE_STAFF[st][si]] = st;
  }
  var ng = 0;
  for (var k = 0; k < names.length; k++) {
    var sheetNm = map[names[k]];
    if (typeof RETIRED_STAFF !== 'undefined' && RETIRED_STAFF.indexOf(sheetNm) >= 0) {
      Logger.log('【要対応】' + sheetNm + ' は名簿では在籍ですが、RETIRED_STAFF に入っています');
      Logger.log('    → このまま fixAllKyuyoIchiran() を流すと、給与一覧の集計から外れます');
      ng++;
    }
    if (!inStore[sheetNm]) {
      Logger.log('【要対応】' + sheetNm + ' が STORE_STAFF のどの現場にも入っていません');
      Logger.log('    → 給与一覧の集計表に出ません（出勤簿シートへの転記はされます）');
      ng++;
    }
  }
  if (!ng) Logger.log('食い違いはありません。');

  Logger.log('―――― 出勤簿に実際のシートがあるかも見る（今の期間から先を全部） ――――');
  var books = booksFromNow_();
  if (!books.length) { Logger.log('見にいける給与一覧がありませんでした'); return; }
  var missing = 0;
  for (var b = 0; b < books.length; b++) {
    var ss = SpreadsheetApp.openById(books[b].id);
    var lack = [];
    for (var j = 0; j < names.length; j++) {
      if (!ss.getSheetByName(map[names[j]])) lack.push(map[names[j]]);
    }
    if (lack.length) {
      Logger.log('【要対応】' + books[b].label + '：シートがありません → ' + lack.join('、'));
      missing += lack.length;
    }
    Utilities.sleep(300);
  }
  if (!missing) Logger.log('シートは全期間そろっています。');
  Logger.log('確認おわり。【要対応】が出ていなければ大丈夫です。');
  if (missing) {
    Logger.log('※ シートを作るには addStaffSheetToBooks() を使ってください。');
  }
}

/* 今の期間から先の給与一覧を、期間の新しい順ではなく古い順に並べて返す */
function booksFromNow_() {
  var today = new Date();
  var todayNum = today.getFullYear() * 10000 + (today.getMonth() + 1) * 100 + today.getDate();
  var idSheet = getOrCreateIdSheet();
  var out = [];
  for (var i = 0; i < PERIODS.length; i++) {
    var p = PERIODS[i];
    var e = p.end.split('/');
    var eNum = (+e[0]) * 10000 + (+e[1]) * 100 + (+e[2]);
    if (eNum < todayNum) continue;          // もう終わった期間は見ない
    var id = getSSIdForPeriod(idSheet, p.label);
    if (id) out.push({ label: p.label, id: id });
  }
  return out;
}

/* =============================================================
   新しく入った人の出勤簿シートを、今の期間から先にまとめて作る

   使い方（エディタでこの関数の中身を書き換えて実行）：
     name     … 出勤簿でのシート名（名簿の「給与一覧での表記」と同じにする）
     hourly   … 時給（円）
     commute  … 通勤手当（1日あたりの円）。無いときは 0
     template … 同じ現場の人のシート名。書式と手当の作りをそのまま写す

   すでにシートがある期間は、何もしません（上書きしません）。
   ============================================================= */
function addStaffSheetToBooks() {
  var name     = '岡田 梨沙';
  var hourly   = 1080;
  var commute  = 1;
  var template = '松本めぐみ';   // 同じトータルステイの人

  var books = booksFromNow_();
  var made = 0, skipped = 0;
  for (var b = 0; b < books.length; b++) {
    var ss = SpreadsheetApp.openById(books[b].id);
    if (ss.getSheetByName(name)) {
      Logger.log('あるのでそのまま: ' + books[b].label);
      skipped++;
      continue;
    }
    var tmpl = ss.getSheetByName(template);
    if (!tmpl) {
      Logger.log('【要対応】雛形が見つかりません: ' + books[b].label + ' / ' + template);
      continue;
    }
    var sheet = tmpl.copyTo(ss);
    sheet.setName(name);

    sheet.getRange(3, 11).setValue(hourly);    // K3 時給
    sheet.getRange(3, 13).setValue(commute);   // M3 通勤手当
    sheet.getRange(3, 2).setValue('時給：' + hourly + '円');
    sheet.getRange(3, 6).setValue(
      (commute > 0 ? '通勤手当：' + commute + '円×出勤日数' : '通勤手当：なし')
      + '　深夜割増：22:00〜翌5:00（法定）');

    // 1行目の見出しの氏名を差し替える
    var title = String(sheet.getRange(1, 2).getValue())
      .replace(/出勤簿　.*?（/, '出勤簿　' + name + '（');
    sheet.getRange(1, 2).setValue(title);

    // 出退勤の中身だけ消す（日付と曜日はその期間のまま使う）
    var vals = sheet.getDataRange().getValues();
    for (var r = 0; r < vals.length; r++) {
      var cell = vals[r][1];
      var isDateRow = (cell instanceof Date) || /^\d+\/\d+$/.test(String(cell).trim());
      if (!isDateRow) continue;
      sheet.getRange(r + 1, 4).setValue('');   // 出勤
      sheet.getRange(r + 1, 5).setValue('');   // 退勤
      sheet.getRange(r + 1, 6).setValue(0);    // 休憩
      sheet.getRange(r + 1, 9).setValue('');   // 備考
    }

    // 雛形にした人の隣に置く
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(tmpl.getIndex() + 1);

    Logger.log('作成: ' + books[b].label + ' → ' + name);
    made++;
    Utilities.sleep(800);
  }
  Logger.log('■完了: ' + made + '件作成 / ' + skipped + '件はもうありました');
  if (made) Logger.log('※ このあと fixAllKyuyoIchiran() を流すと、給与一覧の集計にも出ます。');
}

// 名簿を直したあと、5分待たずに反映させたいとき
function clearRosterCache() {
  try { CacheService.getScriptCache().remove(ROSTER_CACHE); } catch (e) {}
  Logger.log('名簿の控えを消しました。次の転記で読み直します。');
}
