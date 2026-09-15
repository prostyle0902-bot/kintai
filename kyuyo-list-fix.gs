/* =============================================================
   出勤簿GAS：給与一覧（集計タブ）の抜けを直し、退職者を外す

   ※ 先に kyuyo-roster-patch.gs（名簿パッチ）を入れてください。
      このファイルは、その中の関数を使います。

   ―――― 抜けが出る理由 ――――
   集計タブは、一人ひとりの出勤簿シートを参照して作られています。
       ='松本めぐみ'!D37   のような式
   その人の出勤簿シートがそのブックに無いと、式が作られず、
   氏名だけ書かれて出勤日数・実働・支給額が空になります。
   新しく入った人は、まだシートが無い期間があるので、そこが抜けます。

   ―――― このファイルがやること ――――
   ① 名簿の在籍者で、出勤簿シートが無い人のシートを作る（抜けの元を断つ）
   ② 名簿で退職になっている人を、集計タブから外す
      ただし、その期間に出勤がある人は外しません（給与が消えるため）
   ③ 集計タブを組み直し、通し番号を振り直す
      （「コイララ　マダバ」のような空の重複行も、ここで消えます）

   出勤簿シートそのものは消しません。過去の記録が失われるためです。
   タブが多くて見づらいときは hideRetiredSheets() で隠せます（戻せます）。

   ―――― 使い方 ――――
   1. 出勤簿GASに新しいファイルを作って、この中身を貼り付ける
   2. previewKyuyoList() を実行して、何が変わるかログで確かめる
   3. fixKyuyoList() を実行する（1回で2冊ずつ。完了と出るまで繰り返す）
   ============================================================= */


// ① 何が変わるかを見るだけ（書き換えなし）
function previewKyuyoList() {
  var books = booksFromNow_();
  if (!books.length) { Logger.log('対象の給与一覧がありません'); return; }
  var actives = rosterRecords_();
  var retired = rosterRetired_();

  Logger.log('名簿：在籍 ' + actives.length + '人 / 退職 ' + retired.length + '人');
  for (var b = 0; b < books.length; b++) {
    var ss = SpreadsheetApp.openById(books[b].id);

    var lack = [];
    for (var i = 0; i < actives.length; i++) {
      var nm = sheetNameOf_(actives[i]);
      if (!ss.getSheetByName(nm)) lack.push(nm);
    }

    var out = [], keep = [];
    for (var j = 0; j < retired.length; j++) {
      var rn = retired[j];
      if (!inStoreStaff_(rn)) continue;
      var days = workedDays_(ss, rn);
      if (days > 0) keep.push(rn + '（' + days + '日出勤）'); else out.push(rn);
    }

    Logger.log(books[b].label
      + '\n    シートが無い人（抜けの元）：' + (lack.length ? lack.join('、') : 'なし')
      + '\n    集計から外す退職者：' + (out.length ? out.join('、') : 'なし')
      + '\n    出勤があるので残す退職者：' + (keep.length ? keep.join('、') : 'なし'));
    Utilities.sleep(300);
  }
  Logger.log('※確認のみ。直すのは fixKyuyoList() です。');
}


// ② 実際に直す。1回で2冊ずつ。完了と出るまで繰り返し実行する
function fixKyuyoList() {
  var props = PropertiesService.getScriptProperties();
  var lastIdx = parseInt(props.getProperty('kyuyoListIdx') || '0');
  var books = booksFromNow_();
  if (!books.length) { Logger.log('対象の給与一覧がありません'); return; }

  var base = deepCopyStoreStaff_();   // 元の STORE_STAFF を控えておく
  var end = Math.min(lastIdx + 2, books.length);

  for (var k = lastIdx; k < end; k++) {
    var t = books[k];
    try {
      restoreStoreStaff_(base);   // 冊ごとにまっさらから組み直す
      var ss = SpreadsheetApp.openById(t.id);

      // (1) 抜けの元を断つ：シートが無い在籍者のシートを作る
      var actives = rosterRecords_();
      var made = 0;
      for (var i = 0; i < actives.length; i++) {
        var nm = sheetNameOf_(actives[i]);
        if (ss.getSheetByName(nm)) continue;
        if (createStaffSheet_(ss, nm, actives[i], null)) { made++; Utilities.sleep(400); }
      }

      // (2) 名簿の在籍者を集計に載せる
      var added = mergeRosterIntoStoreStaff_();

      // (3) 退職者を外す。ただしその期間に出勤がある人は残す
      var removed = [], kept = [];
      var retired = rosterRetired_();
      for (var j = 0; j < retired.length; j++) {
        var rn = retired[j];
        if (!inStoreStaff_(rn)) continue;
        if (workedDays_(ss, rn) > 0) { kept.push(rn); continue; }
        dropFromStoreStaff_(rn);
        removed.push(rn);
      }

      // (4) 集計タブを組み直す
      fixKyuyoIchiranInSS(t.id);

      Logger.log(t.label
        + '｜シート作成 ' + made + '枚'
        + '／集計に追加 ' + (added.length ? added.join('、') : 'なし')
        + '／集計から外した退職者 ' + (removed.length ? removed.join('、') : 'なし')
        + (kept.length ? '／出勤があるので残した退職者 ' + kept.join('、') : ''));
    } catch (e) {
      Logger.log('エラー: ' + t.label + ' / ' + e.message);
    }
    Utilities.sleep(1000);
  }

  restoreStoreStaff_(base);
  props.setProperty('kyuyoListIdx', String(end));

  if (end >= books.length) {
    Logger.log('■全' + books.length + '冊 完了！');
    props.deleteProperty('kyuyoListIdx');
  } else {
    Logger.log('■' + end + '/' + books.length + '冊まで完了。'
               + 'もう一度 fixKyuyoList() を実行してください。');
  }
}

// 最初からやり直したいとき
function resetKyuyoListCheckpoint() {
  PropertiesService.getScriptProperties().deleteProperty('kyuyoListIdx');
  Logger.log('次回は最初から処理します。');
}


/* =============================================================
   退職者の出勤簿タブを隠す／戻す
   消さずに隠すだけなので、あとからいつでも戻せます。
   ============================================================= */
function hideRetiredSheets() {
  var retired = rosterRetired_();
  if (!retired.length) { Logger.log('名簿に退職者がいません'); return; }
  var books = booksFromNow_();
  var n = 0;
  for (var b = 0; b < books.length; b++) {
    var ss = SpreadsheetApp.openById(books[b].id);
    for (var i = 0; i < retired.length; i++) {
      var sh = ss.getSheetByName(retired[i]);
      if (!sh || sh.isSheetHidden()) continue;
      if (workedDays_(ss, retired[i]) > 0) {
        Logger.log('出勤があるので隠しません: ' + books[b].label + ' / ' + retired[i]);
        continue;
      }
      sh.hideSheet();
      n++;
    }
    Utilities.sleep(400);
  }
  Logger.log('■' + n + '枚のタブを隠しました。戻すときは showRetiredSheets() です。');
}

function showRetiredSheets() {
  var retired = rosterRetired_();
  var books = booksFromNow_();
  var n = 0;
  for (var b = 0; b < books.length; b++) {
    var ss = SpreadsheetApp.openById(books[b].id);
    for (var i = 0; i < retired.length; i++) {
      var sh = ss.getSheetByName(retired[i]);
      if (sh && sh.isSheetHidden()) { sh.showSheet(); n++; }
    }
    Utilities.sleep(400);
  }
  Logger.log('■' + n + '枚のタブを戻しました。');
}


/* ---------------- ここから中身 ---------------- */

/* 名簿で退職になっている人の、出勤簿でのシート名を集める */
function rosterRetired_() {
  var out = [];
  try {
    var sh = SpreadsheetApp.openById(ROSTER_SS_ID).getSheetByName(ROSTER_SHEET);
    if (!sh || sh.getLastRow() < 2) return out;
    var values = sh.getDataRange().getValues();
    var head = values[0];
    var col = {};
    for (var c = 0; c < head.length; c++) col[String(head[c]).trim()] = c;
    if (col.name === undefined || col.status === undefined) return out;

    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var name = String(row[col.name] || '').trim();
      if (!name) continue;
      if (String(row[col.status]).trim() !== 'retired') continue;
      var kn = col.kyuyoName === undefined ? '' : String(row[col.kyuyoName] || '').trim();
      out.push(kn || name);
    }
  } catch (e) {
    Logger.log('名簿を読めませんでした: ' + e.message);
  }
  return out;
}

/* その人が、そのブックの期間に何日出勤しているか。
   出勤簿シートの「合計」行の D列（例「12日出勤」）を見る。 */
function workedDays_(ss, sheetName) {
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return 0;
  var vals = sh.getRange(1, 2, Math.max(sh.getLastRow(), 1), 3).getValues();
  for (var r = 0; r < vals.length; r++) {
    if (String(vals[r][0]).trim() !== '合計') continue;
    var d = String(sh.getRange(r + 1, 4).getValue());
    var m = d.match(/(\d+)/);
    return m ? parseInt(m[1], 10) : 0;
  }
  return 0;
}

function inStoreStaff_(name) {
  for (var st in STORE_STAFF) if (STORE_STAFF[st].indexOf(name) >= 0) return true;
  return false;
}

function dropFromStoreStaff_(name) {
  for (var st in STORE_STAFF) {
    var i = STORE_STAFF[st].indexOf(name);
    if (i >= 0) STORE_STAFF[st].splice(i, 1);
  }
}

function deepCopyStoreStaff_() {
  var out = {};
  for (var st in STORE_STAFF) out[st] = STORE_STAFF[st].slice();
  return out;
}

function restoreStoreStaff_(base) {
  for (var st in STORE_STAFF) delete STORE_STAFF[st];
  for (var k in base) STORE_STAFF[k] = base[k].slice();
}
