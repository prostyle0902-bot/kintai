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
 * 同じファイルの中で、同じ行に正しい計算式が入っている人から写します。
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
 * 直すのは計算式が空の人だけです。すでに式が入っている人には触りません。
 */

/* 見に行くファイル。kintai_ で始まるスプレッドシートを全部さがす。
   名前を変えている場合は、ここを直してください。 */
var ZANGYO_FILE_PREFIX = 'kintai_';

/* 一人ひとりのシートではないもの（集計タブなど）。
   この文字が名前に含まれるシートは見ません。 */
var ZANGYO_SKIP = ['給与一覧', '_時刻リスト', 'SS_ID管理', 'シート'];

/* さがす行の名前。書き方のゆれを拾えるよう、部分一致で見る。
   見つからないときは、checkZangyo がそのシートの中身を出すので、
   そこに出た言葉をここに足してください。 */
var ZANGYO_LABELS = ['残業手当', '時間外手当', '時間外労働手当', '割増手当', '割増賃金', '残業', '時間外'];

/* 上の言葉が入っていても、残業手当の行ではないもの。
   ・日ごとの表の見出し（「残業(h)」など）
   ・深夜手当・休日手当（別の手当なので、取り違えると金額が狂う）
   「深夜割増」を残業手当と間違えないために、ここが要ります。 */
var ZANGYO_NOT = ['(h)', '（h）', '時間数', '合計時間', '深夜', '休日', '法定休'];


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
  var dumped = false;

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
    if (!people.length) {
      lines.push('　残業手当の行があるシートがありません。');
      /* どう書いてあるのか分からないと直しようがないので、
         1枚目の人のシートの中身をそのまま出す。
         ここに出た言葉を ZANGYO_LABELS に足せば、次から見つかる。 */
      if (!dumped) {
        var one = null;
        for (var d = 0; d < sheets.length; d++) {
          if (!zangyoSkip_(sheets[d].getName())) { one = sheets[d]; break; }
        }
        if (one) { lines.push(zangyoDump_(one)); dumped = true; }
      }
      continue;
    }

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
      if (!v || v.length > 20) continue;              // 長い説明文は見出しではない
      var bad = false;
      for (var N = 0; N < ZANGYO_NOT.length; N++) {
        if (v.indexOf(ZANGYO_NOT[N]) >= 0) { bad = true; break; }
      }
      if (bad) continue;
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


/* 1枚のシートに何が書いてあるかを、そのまま出す（読むだけ）。
   「残業手当」がどう書かれているのか分からないときに使う。
   文字が入っているセルだけを、場所つきで並べる。 */
function zangyoDump_(sheet) {
  var lastRow = Math.min(sheet.getLastRow(), 80);
  var lastCol = Math.min(sheet.getLastColumn(), 14);
  if (lastRow < 1 || lastCol < 1) return '　（中身がありません）';

  var vals = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var fs   = sheet.getRange(1, 1, lastRow, lastCol).getFormulas();
  var out = ['　── ' + sheet.getName() + ' の中身（' + lastRow + '行 × ' + lastCol + '列）──'];

  for (var r = 0; r < lastRow; r++) {
    var cells = [];
    for (var c = 0; c < lastCol; c++) {
      var f = fs[r][c];
      var v = vals[r][c];
      if (f) { cells.push(zangyoCol_(c + 1) + (r + 1) + '=' + f); continue; }
      if (v === '' || v === null) continue;
      var t = String(v);
      if (t.length > 28) t = t.slice(0, 28) + '…';
      cells.push(zangyoCol_(c + 1) + (r + 1) + ':' + t);
    }
    if (cells.length) out.push('　' + cells.join('  '));
  }
  out.push('　── ここまで ──');
  out.push('　※この中の「残業」にあたる言葉を、上の ZANGYO_LABELS に足してください。');
  return out.join('\n');
}

/* 1枚だけ中身を見たいとき用。シート名を書いて実行する。 */
function dumpOne() {
  var NAME = '石毛貴英';                       // ← 見たい人の名前に変える
  var files = zangyoFiles_();
  for (var i = 0; i < files.length; i++) {
    var ss = SpreadsheetApp.openById(files[i].getId());
    var sh = ss.getSheetByName(NAME);
    if (!sh) continue;
    var msg = '【' + ss.getName() + '】\n' + zangyoDump_(sh);
    Logger.log(msg);
    return msg;
  }
  var msg2 = NAME + ' のシートが見つかりませんでした。';
  Logger.log(msg2);
  return msg2;
}
