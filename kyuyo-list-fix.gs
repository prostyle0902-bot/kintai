/* =============================================================
   出勤簿GAS：給与一覧（集計タブ）の抜けを直し、退職者を外す

   ※ 先に kyuyo-roster-patch.gs（名簿パッチ）を入れてください。
      このファイルは、その中の関数を使います。

   ―――― 抜けが出る理由（2種類）――――
   集計タブは、一人ひとりの出勤簿シートを参照する式で作られています。
       ='松本めぐみ'!D37 ←出勤日数     ='松本めぐみ'!G48 ←支給額

   (A) 出勤簿シートが無い人は、式が作られず、氏名だけ書かれて全部空になる。
       退職して出勤簿シートを消した人（田村あつ子さんなど）がこれ。

   (B) 支給額だけ空になる人がいる。
       集計を組む処理が、明細の「差引支給額」を"完全一致"で探しているため。
       衣幡千明さんは海事・横河との合算を入れたときに
       「差引支給額（トータルステイ）」に名前が変わり、見つからなくなった。
       同じ書き方のシートを写して作られた岡田梨沙さんも同じ。

   ―――― このファイルがやること ――――
   ① 名簿の在籍者で、出勤簿シートが無い人のシートを作る
   ② 名簿で退職になっている人を、集計タブから外す
      ただし、その期間に出勤がある人は外しません（給与が消えるため）
   ③ 集計タブを組み直し、通し番号を振り直す
   ④ そのあと、空いている項目に式を入れ直す
      「差引支給額（…）」のように名前が違っても見つけられるようにする

   出勤簿シートそのものは消しません。過去の記録が失われるためです。
   タブが多くて見づらいときは hideRetiredSheets() で隠せます（戻せます）。

   ―――― 使い方 ――――
   1. 出勤簿GASに新しいファイルを作って、この中身を貼り付ける
   2. previewKyuyoList() を実行して、何が変わるかログで確かめる
   3. fixKyuyoList() を実行する（1回で2冊ずつ。完了と出るまで繰り返す）
   ============================================================= */


/* 名簿には無いが、集計に載せたい出勤簿シート。
   ひとりが2つの現場で別々のシートを持つ場合など。
   桜庭京子さんのシートを作り替えた「衣幡千明_海事横河」がこれにあたる。
   ここに入れておかないと、その現場の支給額が総合計から抜ける。 */
var EXTRA_SHEETS = {
  '海事・横河': ['衣幡千明_海事横河']
};


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

    var blank = blankRows_(ss);
    Logger.log(books[b].label
      + '\n    いま項目が空の行：' + (blank.length ? blank.join('、') : 'なし')
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

      // (4) 名簿に無い特別なシートを載せる（衣幡千明_海事横河 など）
      var extra = mergeExtraSheets_(ss);

      // (5) 集計タブを組み直す
      fixKyuyoIchiranInSS(t.id);

      // (6) 空いている項目に式を入れ直す
      var patched = patchKyuyoFormulas_(ss);

      var all = added.concat(extra);
      Logger.log(t.label
        + '｜シート作成 ' + made + '枚'
        + '／集計に追加 ' + (all.length ? all.join('、') : 'なし')
        + '／集計から外した退職者 ' + (removed.length ? removed.join('、') : 'なし')
        + (kept.length ? '／出勤があるので残した退職者 ' + kept.join('、') : '')
        + '／式を入れ直した行 ' + patched);
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


/* 名簿に無い特別なシートを、集計に載せる */
function mergeExtraSheets_(ss) {
  var out = [];
  for (var g in EXTRA_SHEETS) {
    if (!STORE_STAFF[g]) STORE_STAFF[g] = [];
    var list = EXTRA_SHEETS[g];
    for (var i = 0; i < list.length; i++) {
      if (!ss.getSheetByName(list[i])) continue;   // そのブックに無ければ何もしない
      if (inStoreStaff_(list[i])) continue;
      STORE_STAFF[g].push(list[i]);
      out.push(list[i] + ' \u2192 ' + g);
    }
  }
  return out;
}

/* 集計タブの、空いている項目に式を入れ直す。

   組み直す処理は明細の「差引支給額」を完全一致で探すため、
   「差引支給額（トータルステイ）」のように名前が違うと見つからず、
   支給額だけ空になる。ここでは前方一致で探して入れ直す。
   合算支給額は使わない。海事・横河分は別の行で計上されるため、
   合算を使うと二重に足されてしまう。 */
function patchKyuyoFormulas_(ss) {
  var sheet = kyuyoSheetOf_(ss);
  if (!sheet || sheet.getLastRow() < 2) return 0;

  var vals = sheet.getRange(1, 1, sheet.getLastRow(), 6).getValues();
  var fixed = 0;

  for (var r = 0; r < vals.length; r++) {
    var nm = String(vals[r][1] || '').trim();
    if (!nm || !isNo_(vals[r][0])) continue;
    if (nm.indexOf("'") >= 0) continue;   // 式に使えない名前は触らない

    var st = ss.getSheetByName(nm);
    if (!st) continue;

    var sd = st.getRange(1, 2, Math.max(st.getLastRow(), 1), 1).getValues();
    var totalRow = -1, detailRow = -1;
    for (var k = 0; k < sd.length; k++) {
      var b = String(sd[k][0]).trim();
      if (b === '合計' && totalRow < 0) totalRow = k + 1;
      if (b.indexOf('差引支給額') === 0 && detailRow < 0) detailRow = k + 1;
    }

    var row = r + 1, did = false;
    if (totalRow > 0) {
      if (vals[r][2] === '') { sheet.getRange(row, 3).setFormula("='" + nm + "'!D" + totalRow); did = true; }
      if (vals[r][3] === '') { sheet.getRange(row, 4).setFormula("='" + nm + "'!G" + totalRow); did = true; }
      if (vals[r][4] === '') { sheet.getRange(row, 5).setFormula("='" + nm + "'!H" + totalRow); did = true; }
    }
    if (detailRow > 0 && vals[r][5] === '') {
      sheet.getRange(row, 6).setFormula("='" + nm + "'!G" + detailRow).setNumberFormat('#,##0');
      did = true;
    }
    if (did) { fixed++; Logger.log('  式を入れ直しました: ' + nm); }
  }
  return fixed;
}

/* いま集計タブで項目が空になっている人を挙げる（確認用） */
function blankRows_(ss) {
  var sheet = kyuyoSheetOf_(ss);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var vals = sheet.getRange(1, 1, sheet.getLastRow(), 6).getValues();
  var out = [];
  for (var r = 0; r < vals.length; r++) {
    var nm = String(vals[r][1] || '').trim();
    if (!nm || !isNo_(vals[r][0])) continue;
    var miss = [];
    if (vals[r][2] === '') miss.push('出勤日数');
    if (vals[r][3] === '') miss.push('実働');
    if (vals[r][4] === '') miss.push('深夜');
    if (vals[r][5] === '') miss.push('支給額');
    if (miss.length) out.push(nm + '（' + miss.join('・') + '）');
  }
  return out;
}

function kyuyoSheetOf_(ss) {
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf('給与一覧') >= 0) return sheets[i];
  }
  return null;
}

// A列が通し番号（1,2,3…）の行かどうか
function isNo_(v) {
  if (typeof v === 'number') return v > 0;
  return /^[0-9]+$/.test(String(v).trim());
}
