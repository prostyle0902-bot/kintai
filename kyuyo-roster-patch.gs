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
   エディタでこの関数を選んで実行し、実行ログを見る。 */
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
  Logger.log('―――― 出勤簿に実際のシートがあるかも見る ――――');
  var ssId = findSSIdForDate(Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd'));
  if (!ssId) { Logger.log('今日の日付に対応する給与一覧が見つかりません'); return; }
  var ss = SpreadsheetApp.openById(ssId);
  for (var j = 0; j < names.length; j++) {
    if (!ss.getSheetByName(map[names[j]])) {
      Logger.log('【要対応】シートがありません：' + map[names[j]] + '（' + ss.getName() + '）');
    }
  }
  Logger.log('確認おわり。【要対応】が出ていなければ大丈夫です。');
}

// 名簿を直したあと、5分待たずに反映させたいとき
function clearRosterCache() {
  try { CacheService.getScriptCache().remove(ROSTER_CACHE); } catch (e) {}
  Logger.log('名簿の控えを消しました。次の転記で読み直します。');
}
