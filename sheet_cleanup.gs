/**
 * 出勤簿の日付欄を整えるスクリプト
 *
 * 対象：22期_PA勤怠フォルダの給与一覧スプレッドシート（2026/9/16〜2027/9/15 の12本）
 * 必要ファイル：attendance_lib.gs
 *
 * 事前作成された22期のスプレッドシートには、日付欄に2つの問題がある。
 *
 *   1. 期間外のゴミ行
 *      日次データの枠が31日分あり、30日以下の月は末尾が前のテンプレートの
 *      ままになっている。たとえば 2026/9/16〜10/15（30日）は R36 に「4/15」、
 *      2027/2/16〜3/15（28日）は R34〜R36 に「4/13」「4/14」「4/15」が残る。
 *      → cleanupExtraRows()
 *
 *   2. 日付の年が1年ずれている
 *      2027年になるはずの日付が2026年で入っている。表示書式が m/d なので
 *      画面では気づけないが、値としては1年前になっている。
 *      曜（C列）と集計期間（4行目）は正しいので、そこから正しい日付を入れ直す。
 *      → fixDateColumn()
 *
 * 使い方
 *   1. attendance_lib.gs と一緒に Apps Script プロジェクトに追加
 *   2. CLEAN.DRY_RUN = true のまま実行し、ログで対象行を確認
 *   3. 問題なければ CLEAN.DRY_RUN = false にして再実行
 */

const CLEAN = {
  // true の間は一切書き込まず、ログ出力だけ行う
  DRY_RUN: true,

  // 対象フォルダ。期が変わったら追加する
  FOLDER_IDS: [
    '1JQPlAe-jOMhCxg2jIbajcIyG0w7zXa6F', // 22期_PA勤怠
  ],

  // 個別に対象を指定したい場合はこちらに ID を入れる（FOLDER_IDS より優先）
  SPREADSHEET_IDS: [],

  // 消す列（日付〜備考）
  FIRST_COL: 2, // B
  LAST_COL: 9,  // I
};


/**
 * 集計期間からはみ出した行の中身を消す。
 *
 * 実働(G)・深夜(H) の式も一緒に消える。これらは G6/H6 を先頭とする共有数式なので
 * 先頭は残り、必要になったら1つ上の行から下方向にコピーすれば戻せる。
 */
function cleanupExtraRows() {
  eachBook(function (ss, log) {
    let tabs = 0, rows = 0;
    const bookPeriod = readBookPeriod(ss);
    ss.getSheets().forEach(function (sh) {
      const at = readAttendanceRows(sh);
      if (!at) return;
      const p = tabPeriod(sh, at, bookPeriod, log);
      if (!p) return;

      const first = at.dataStart + p.days;
      if (first > at.dataEnd) return;                 // ちょうど埋まっている月

      const n = at.dataEnd - first + 1;
      const range = sh.getRange(first, CLEAN.FIRST_COL, n, CLEAN.LAST_COL - CLEAN.FIRST_COL + 1);
      if (range.getValues().every(function (r) {
        return r.every(function (v) { return v === '' || v === null; });
      })) return;                                     // 既に空

      if (tabs === 0) {
        log.push('  期間 ' + fmtDate(p.start) + '〜' + fmtDate(p.end) + '（' + p.days + '日）'
          + ' → R' + first + '〜R' + at.dataEnd + ' の ' + n + '行が期間外');
      }
      tabs += 1; rows += n;
      if (!CLEAN.DRY_RUN) range.clearContent();
    });
    log.push('- 期間外の行を消したタブ: ' + tabs + '（計 ' + rows + '行）');
  });
}


/**
 * 日付欄（B列）を集計期間から入れ直す。
 * 曜（C列）が集計期間と食い違うタブは、構成が想定と違うので手を付けずに飛ばす。
 */
function fixDateColumn() {
  eachBook(function (ss, log) {
    let tabs = 0, cells = 0;
    let sample = '';
    const bookPeriod = readBookPeriod(ss);
    ss.getSheets().forEach(function (sh) {
      const at = readAttendanceRows(sh);
      if (!at) return;
      const p = tabPeriod(sh, at, bookPeriod, log);
      if (!p) return;

      const range = sh.getRange(at.dataStart, CLEAN.FIRST_COL, p.days, 1);
      const cur = range.getValues();
      let changed = 0;
      for (let i = 0; i < p.days; i++) {
        const want = addDays(p.start, i);
        if (dateKey(cur[i][0]) === dateKey(want)) continue;
        if (!sample) {
          sample = 'R' + (at.dataStart + i) + ' ' + describe(cur[i][0]) + ' → ' + fmtDate(want);
        }
        cur[i][0] = want;
        changed += 1;
      }
      if (!changed) return;

      tabs += 1; cells += changed;
      if (!CLEAN.DRY_RUN) range.setValues(cur);
    });
    if (sample) log.push('  例: ' + sample);
    log.push('- 日付を直したタブ: ' + tabs + '（計 ' + cells + 'セル）');
  });
}


/* ------------------------------------------------------------------ */

function eachBook(fn) {
  const log = [];
  const books = CLEAN.SPREADSHEET_IDS.length
    ? CLEAN.SPREADSHEET_IDS.map(function (id) { return { id: id, name: id }; })
    : listPayrollBooks(CLEAN.FOLDER_IDS, log);

  books.forEach(function (file) {
    let ss;
    try {
      ss = SpreadsheetApp.openById(file.id);
    } catch (e) {
      log.push('!! 開けません ' + file.name + ' : ' + e.message);
      return;
    }
    log.push('');
    log.push('========== ' + ss.getName() + ' ==========');
    try {
      fn(ss, log);
    } catch (e) {
      log.push('!! 中断: ' + e.message);
    }
  });

  log.push('');
  log.push(CLEAN.DRY_RUN
    ? '*** DRY_RUN です。書き込みは行っていません。***'
    : '*** 書き込みを実行しました。***');
  Logger.log(log.join('\n'));
}

function describe(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return fmtDate(v);
  return v === '' || v === null ? '(空)' : '「' + v + '」';
}
