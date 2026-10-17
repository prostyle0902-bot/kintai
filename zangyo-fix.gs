/**
 * 出勤簿GAS：残業手当の計算式の抜けを直す
 *
 * ＜何のためのものか＞
 * kintai_YYYYMM（社員の出勤簿）の一人ひとりのシートには、給与明細の
 * 「残業手当」の行があります。ここに計算式が入っていない人がいて、
 * 残業したぶんが支給額に乗らないまま気づかれない、ということが起きました
 * （2026年9月分の石毛貴英さん）。
 *
 * ＜どう直すか＞
 * 計算式の中身は決め打ちにしません。会社ごと・時期ごとに変わるためです。
 * <b>同じファイルの中で、同じ行に正しい計算式が入っている人から写します。</b>
 *   ・「残業手当」の行が同じ行番号の人だけを見本にする
 *     （シートの作りが同じでないと、写した式が別の場所を見てしまうため）
 *   ・見本の式が他のシートを指している（! を含む）ときは写さない
 *   ・見本が1人もいない、見本どうしで式が違う、といったときも写さない
 * 写せなかったものは、理由を付けて全部ログに出します。黙って飛ばしません。
 *
 * ＜使い方＞
 * 1. 出勤簿GAS（kintai_YYYYMM を作っている Apps Script）に
 *    新しいファイルを作って、この中身を貼り付ける
 * 2. 関数の選択欄から  checkZangyo  を選んで「実行」
 *    → 読むだけ。何も書き換えません。実行ログに今の状態が出ます
 * 3. ログを見て問題なければ  fixZangyo  を選んで「実行」
 *    → 計算式を入れます
 *
 * 直すのは<b>計算式が空の人</b>だけです。すでに式が入っている人には触りません。
 */

/* 見に行くファイル。kintai_ で始まるスプレッドシートを全部さがす。
   名前を変えている場合は、ここを直してください。 */
var ZANGYO_FILE_PREFIX = 'kintai_';

/* 一人ひとりのシートではないもの（集計タブなど）。
   この文字が名前に含まれるシートは見ません。 */
var ZANGYO_SKIP = ['給与一覧', '_時刻リスト', 'SS_ID管理', 'シート'];

/* さがす行の名前。「残業手当」のほか、書き方のゆれも拾う。 */
var ZANGYO_LABELS = ['残業手当', '時間外手当', '割増手当'];


/* =========================================================
   調べる（読むだけ・何も変えない）
   ========================================================= */
function checkZangyo() {
  var out = zangyoRun_(true);
  Logger.log(out);
  return out;
}

/* =========================================================
   直す（計算式が空の人にだけ入れる）
   ========================================================= */
function fixZangyo() {
  var out = zangyoRun_(false);
  Logger.log(out);
  return out;
}


function zangyoRun_(dryRun) {
  var lines = [];
  lines.push(dryRun ? '■ 調べるだけ（何も書き換えません）' : '■ 直します');

  var files = zangyoFiles_();
  if (!files.length) {
    lines.push('「' + ZANGYO_FILE_PREFIX + '」で始まるスプレッドシートが見つかりませんでした。');
    return lines.join('\n');
  }
  lines.push('見つかったファイル：' + files.length + '件');

  var totalOk = 0, totalFixed = 0, totalLeft = 0;

  for (var i = 0; i < files.length; i++) {
    var f = files[i];
    var ss;
    try { ss = SpreadsheetApp.openById(f.getId()); }
    catch (e) { lines.push('\n【' + f.getName() + '】開けません：' + e); continue; }

    lines.push('\n【' + ss.getName() + '】');

    // --- 1) そのファイルの全員ぶんを調べる ---
    var people = [];
    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      var sh = sheets[s];
      if (zangyoSkip_(sh.getName())) continue;
      var info = zangyoFind_(sh);
      if (!info) continue;                       // 残業手当の行が無いシート
      people.push(info);
    }
    if (!people.length) { lines.push('　残業手当の行があるシートがありません。'); continue; }

    // --- 2) 見本になる人（計算式が入っている人）を集める ---
    var haveFormula = people.filter(function (p) { return p.formula; });
    var missing     = people.filter(function (p) { return !p.formula; });

    lines.push('　人数 ' + people.length + '／式あり ' + haveFormula.length
      + '／式なし ' + missing.length);
    totalOk += haveFormula.length;

    if (!missing.length) continue;

    for (var m = 0; m < missing.length; m++) {
      var me = missing[m];
      var why = '';

      /* 見本は「残業手当の行が同じ行番号の人」だけ。
         シートの作りが同じでないと、写した式が別の場所を見てしまう。
         式が空の人は金額の列も分からないので、見本の列をそのまま使う。 */
      var sames = haveFormula.filter(function (p) { return p.row === me.row; });
      var col = sames.length ? sames[0].col : me.col;

      if (!sames.length) {
        why = '同じ行（' + me.row + '行目）に式が入っている人がいません';
      } else if (!col) {
        why = '金額をどの列に入れればよいか分かりません';
      } else if (!sames.every(function (p) { return p.col === col; })) {
        why = '見本どうしで金額の列が違います（'
          + sames.slice(0, 3).map(function (p) { return p.name + '：' + zangyoCol_(p.col) + '列'; }).join(' / ') + '）';
      } else if (me.col && me.col !== col) {
        why = 'すでに ' + zangyoCol_(me.col) + '列 に何か入っています（見本は ' + zangyoCol_(col) + '列）';
      } else if (sames.some(function (p) { return p.formula.indexOf('!') >= 0; })) {
        why = '見本の式が他のシートを指しています（写すと取り違えます）';
      } else {
        var first = sames[0].formula;
        var allSame = sames.every(function (p) { return p.formula === first; });
        if (!allSame) {
          why = '見本どうしで式が違います（'
            + sames.slice(0, 3).map(function (p) { return p.name + '：' + p.formula; }).join(' / ') + '）';
        } else {
          if (dryRun) {
            lines.push('　・' + me.name + '（' + me.row + '行目 '
              + zangyoCol_(col) + '列）← ' + sames[0].name + ' から写します：' + first);
          } else {
            me.sheet.getRange(me.row, col).setFormula(first);
            lines.push('　・' + me.name + '（' + me.row + '行目 '
              + zangyoCol_(col) + '列）に入れました：' + first
              + '（' + sames[0].name + ' から）');
            totalFixed++;
          }
          continue;
        }
      }
      lines.push('　★ ' + me.name + '（' + me.row + '行目）は写せません：' + why);
      totalLeft++;
    }
  }

  lines.push('\n──────────');
  lines.push('式が入っていた：' + totalOk + '件');
  if (!dryRun) lines.push('入れた：' + totalFixed + '件');
  lines.push('人が手を入れないと直らない：' + totalLeft + '件');
  if (dryRun) lines.push('\n問題なければ fixZangyo を実行してください。');
  return lines.join('\n');
}


/* kintai_ で始まるスプレッドシートを、名前の順に集める */
function zangyoFiles_() {
  var out = [];
  var it = DriveApp.searchFiles(
    'title contains "' + ZANGYO_FILE_PREFIX + '"'
    + ' and mimeType = "application/vnd.google-apps.spreadsheet"'
    + ' and trashed = false');
  while (it.hasNext()) {
    var f = it.next();
    if (f.getName().indexOf(ZANGYO_FILE_PREFIX) !== 0) continue;   // 先頭が一致するものだけ
    out.push(f);
  }
  out.sort(function (a, b) { return a.getName() < b.getName() ? -1 : 1; });
  return out;
}

function zangyoSkip_(name) {
  for (var i = 0; i < ZANGYO_SKIP.length; i++) {
    if (String(name).indexOf(ZANGYO_SKIP[i]) >= 0) return true;
  }
  return false;
}

/* 1枚のシートから「残業手当」の行を見つけて、金額のセルの様子を返す。

   金額がどの列かはシートによって違うので、決め打ちにしない。
   その行の中で「計算式が入っているセル」、無ければ「数値のセル」を金額とみなす。
   どちらも無いときは、同じ行の人たちから列を決めるので、ここでは 0 を返す。 */
function zangyoFind_(sheet) {
  var lastRow = Math.min(sheet.getLastRow(), 120);
  var lastCol = Math.min(sheet.getLastColumn(), 14);
  if (lastRow < 2 || lastCol < 2) return null;

  var vals = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var fs   = sheet.getRange(1, 1, lastRow, lastCol).getFormulas();

  for (var r = 0; r < lastRow; r++) {
    var hit = -1;
    for (var c = 0; c < lastCol; c++) {
      var v = String(vals[r][c] || '').trim();
      if (!v) continue;
      for (var L = 0; L < ZANGYO_LABELS.length; L++) {
        if (v.indexOf(ZANGYO_LABELS[L]) >= 0) { hit = c; break; }
      }
      if (hit >= 0) break;
    }
    if (hit < 0) continue;

    // 見出しより右で、金額が入っていそうなセルをさがす
    var col = 0, formula = '';
    for (var c2 = hit + 1; c2 < lastCol; c2++) {
      if (fs[r][c2]) { col = c2 + 1; formula = fs[r][c2]; break; }
    }
    if (!col) {
      for (var c3 = hit + 1; c3 < lastCol; c3++) {
        if (typeof vals[r][c3] === 'number') { col = c3 + 1; break; }
      }
    }
    return {
      sheet: sheet, name: sheet.getName(), row: r + 1,
      labelCol: hit + 1, col: col, formula: formula,
    };
  }
  return null;
}

// 1 → A, 2 → B …（ログを読みやすくするためだけのもの）
function zangyoCol_(n) {
  var s = '';
  while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
  return s;
}
