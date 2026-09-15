/* =============================================================
   出勤簿GAS：衣幡千明_海事横河 シートの見た目を、他のシートと揃える

   このシートは桜庭京子さんのシートを作り替えたもので、
   あとから全体にかけた書式そろえ（copyFormatsToTargetBooks）のときに
   参照元の 7/16-8/15 ブックに無かったため、ひとりだけ取り残された。

   あわせて、次の2つも直す。
     ・見出しが「衣幡千明（海事・横河）（パート・アルバイト）」と二重
     ・期間外の日付行「4/15」が1行だけ残っている
       （元の3/16〜4/15のブックから作り替えたときの消し忘れ。
        中身が空の行だけ消すので、打刻やシフトは消えない）

   ―――― 使い方 ――――
   1. 出勤簿GASに新しいファイルを作って、この中身を貼り付ける
      （エディタ左の「ファイル」＋ → スクリプト → 名前は kaiji など）
   2. previewKaijiDesign() を実行して、何が直るかログで確かめる
   3. fixKaijiDesign() を実行する（1回で3冊ずつ。完了と出るまで繰り返す）

   ※ 書式だけを写します。出勤・退勤・金額には触れません。
   ============================================================= */

var KAIJI_SHEET = '衣幡千明_海事横河';   // 直したいシート
var KAIJI_REF   = '衣幡千明';            // 見た目のお手本にするシート
var KAIJI_FROM  = 20260816;              // これ以降の期間だけ直す


// ① 何が直るかを見るだけ（書き換えなし）
function previewKaijiDesign() {
  var targets = kaijiTargets_();
  if (!targets.length) { Logger.log('対象の給与一覧が見つかりません'); return; }
  for (var i = 0; i < targets.length; i++) {
    var t = targets[i];
    try {
      var ss  = SpreadsheetApp.openById(t.id);
      var tgt = ss.getSheetByName(KAIJI_SHEET);
      var ref = ss.getSheetByName(KAIJI_REF);
      if (!tgt) { Logger.log(t.label + ' | ' + KAIJI_SHEET + ' なし'); continue; }
      if (!ref) { Logger.log(t.label + ' | お手本の ' + KAIJI_REF + ' なし'); continue; }

      var stray = strayDateRows_(tgt, t.m);
      var title = String(tgt.getRange(1, 2).getValue());
      Logger.log(t.label
        + ' | 期間外の空行:' + (stray.length ? stray.length + '行（消します）' : 'なし')
        + ' / 見出し:' + (title.indexOf('（パート') >= 0 && title.indexOf('海事') >= 0
                          ? '二重（直します）' : 'そのまま')
        + ' / 日付行 ' + dateRows_(tgt).length + '行（お手本 ' + dateRows_(ref).length + '行）');
    } catch (e) {
      Logger.log(t.label + ' | 開けません: ' + e.message);
    }
  }
  Logger.log('※確認のみ。直すのは fixKaijiDesign() です。');
}


// ② 実際に直す。1回で3冊ずつ。完了と出るまで繰り返し実行する
function fixKaijiDesign() {
  var props = PropertiesService.getScriptProperties();
  var lastIdx = parseInt(props.getProperty('kaijiFixIdx') || '0');
  var targets = kaijiTargets_();
  if (!targets.length) { Logger.log('対象の給与一覧が見つかりません'); return; }

  var end = Math.min(lastIdx + 3, targets.length);
  for (var k = lastIdx; k < end; k++) {
    try {
      fixOneKaiji_(targets[k]);
    } catch (e) {
      Logger.log('エラー: ' + targets[k].label + ' / ' + e.message);
    }
    Utilities.sleep(1000);
  }
  props.setProperty('kaijiFixIdx', String(end));

  if (end >= targets.length) {
    Logger.log('■全' + targets.length + '冊 完了！');
    props.deleteProperty('kaijiFixIdx');
  } else {
    Logger.log('■' + end + '/' + targets.length + '冊まで完了。'
               + 'もう一度 fixKaijiDesign() を実行してください。');
  }
}

// 最初からやり直したいとき
function resetKaijiCheckpoint() {
  PropertiesService.getScriptProperties().deleteProperty('kaijiFixIdx');
  Logger.log('次回は最初から処理します。');
}


/* ---------------- ここから中身 ---------------- */

function kaijiTargets_() {
  var idSheet = getOrCreateIdSheet();
  var data = idSheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var label = data[i][0], id = data[i][1];
    if (!label || !id) continue;
    var m = String(label).match(/給与一覧_(\d{4})_(\d{1,2})_(\d{1,2})-(\d{4})_(\d{1,2})_(\d{1,2})/);
    if (!m) continue;
    if ((+m[1]) * 10000 + (+m[2]) * 100 + (+m[3]) < KAIJI_FROM) continue;
    out.push({ label: label, id: id, m: m });
  }
  return out;
}

function fixOneKaiji_(t) {
  var ss  = SpreadsheetApp.openById(t.id);
  var tgt = ss.getSheetByName(KAIJI_SHEET);
  var ref = ss.getSheetByName(KAIJI_REF);
  if (!tgt) { Logger.log('シートなし: ' + t.label + ' / ' + KAIJI_SHEET); return; }
  if (!ref) { Logger.log('お手本なし: ' + t.label + ' / ' + KAIJI_REF); return; }

  // (1) 期間外の空の日付行を消す（下から消す。行番号がずれないように）
  var stray = strayDateRows_(tgt, t.m);
  for (var s = stray.length - 1; s >= 0; s--) tgt.deleteRow(stray[s]);
  if (stray.length) Logger.log('  期間外の空行を' + stray.length + '行消しました: ' + t.label);

  // (2) 見出しを整える
  var period = t.m[1] + '/' + t.m[2] + '/' + t.m[3] + '〜' + t.m[4] + '/' + t.m[5] + '/' + t.m[6];
  tgt.getRange(1, 2).setValue('出勤簿　衣幡千明（海事・横河）　' + period);

  // (3) 上の見出し（1〜5行目）の書式をお手本から写す
  copyFmt_(ref, 1, tgt, 1, 5);

  // (4) 日付ブロックの書式を写す（土曜が青・日曜が赤になる）
  var refRows = dateRows_(ref), tgtRows = dateRows_(tgt);
  if (refRows.length && tgtRows.length) {
    var n = Math.min(refRows.length, tgtRows.length);
    copyFmt_(ref, refRows[0], tgt, tgtRows[0], n);
  }

  // (5) 合計行
  var refTotal = rowOfLabel_(ref, '合計'), tgtTotal = rowOfLabel_(tgt, '合計');
  if (refTotal > 0 && tgtTotal > 0) copyFmt_(ref, refTotal, tgt, tgtTotal, 1);

  // (6) 給与明細は行数が違うので、同じ見出しの行どうしで写す
  var refLabels = labelRows_(ref), tgtLabels = labelRows_(tgt);
  for (var lb in tgtLabels) {
    var hit = refLabels[lb];
    if (hit === undefined) {
      // 「差引支給額」と「差引支給額（トータルステイ）」のような違いを拾う
      for (var rl in refLabels) {
        if (rl.indexOf(lb) === 0 || lb.indexOf(rl) === 0) { hit = refLabels[rl]; break; }
      }
    }
    if (hit !== undefined) copyFmt_(ref, hit, tgt, tgtLabels[lb], 1);
  }

  // (7) 基本賃金の備考を、他のシートと同じ書き方にする
  var kihon = rowOfLabelStarts_(tgt, '基本賃金');
  if (kihon > 0) {
    var rate = tgt.getRange(3, 11).getValue() || 0;
    tgt.getRange(kihon, 9).setValue('時給' + rate + '円×実働合計');
  }

  // (8) 列の幅と行の高さ
  for (var c = 1; c <= 13; c++) {
    try { tgt.setColumnWidth(c, ref.getColumnWidth(c)); } catch (e) {}
  }
  try { tgt.setRowHeight(1, ref.getRowHeight(1)); } catch (e) {}

  Logger.log('整えました: ' + t.label + ' / ' + KAIJI_SHEET);
}

// お手本の書式だけを写す（値・数式には触れない）
function copyFmt_(ref, refRow, tgt, tgtRow, rows) {
  var cols = Math.max(ref.getLastColumn(), tgt.getLastColumn());
  ref.getRange(refRow, 1, rows, cols)
     .copyTo(tgt.getRange(tgtRow, 1, rows, cols),
             SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

// B列が「M/D」または日付の行の、行番号を集める
function dateRows_(sheet) {
  var vals = sheet.getRange(1, 2, Math.max(sheet.getLastRow(), 1), 1).getValues();
  var out = [];
  for (var r = 0; r < vals.length; r++) {
    var v = vals[r][0];
    if ((v instanceof Date) || /^\d+\/\d+$/.test(String(v).trim())) out.push(r + 1);
  }
  return out;
}

/* その期間に無い日付の行を探す。
   出勤・退勤が入っている行は、間違いでも消さない（打刻を失わないため）。 */
function strayDateRows_(sheet, m) {
  var want = {};
  var d = new Date(+m[1], (+m[2]) - 1, +m[3]);
  var end = new Date(+m[4], (+m[5]) - 1, +m[6]);
  while (d <= end) {
    want[(d.getMonth() + 1) + '/' + d.getDate()] = true;
    d.setDate(d.getDate() + 1);
  }
  var rows = dateRows_(sheet);
  var out = [];
  for (var i = 0; i < rows.length; i++) {
    var v = sheet.getRange(rows[i], 2).getValue();
    var md = (v instanceof Date)
      ? (v.getMonth() + 1) + '/' + v.getDate()
      : String(v).trim();
    if (want[md]) continue;
    var inTime  = String(sheet.getRange(rows[i], 4).getValue()).trim();
    var outTime = String(sheet.getRange(rows[i], 5).getValue()).trim();
    if (inTime || outTime) {
      Logger.log('【要対応】期間外なのに中身がある行があります: ' + rows[i] + '行目（' + md + '）'
                 + ' → 消さずに残しました。中身を確かめてください。');
      continue;
    }
    out.push(rows[i]);
  }
  return out;
}

// B列の文字 → 行番号（給与明細の見出しを拾うため）
function labelRows_(sheet) {
  var vals = sheet.getRange(1, 2, Math.max(sheet.getLastRow(), 1), 1).getValues();
  var out = {};
  for (var r = 0; r < vals.length; r++) {
    var v = String(vals[r][0]).trim();
    if (!v) continue;
    if ((vals[r][0] instanceof Date) || /^\d+\/\d+$/.test(v)) continue;
    if (out[v] === undefined) out[v] = r + 1;
  }
  return out;
}

function rowOfLabel_(sheet, label) {
  var rows = labelRows_(sheet);
  return rows[label] === undefined ? -1 : rows[label];
}

function rowOfLabelStarts_(sheet, prefix) {
  var rows = labelRows_(sheet);
  for (var k in rows) if (k.indexOf(prefix) === 0) return rows[k];
  return -1;
}
