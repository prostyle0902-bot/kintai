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
 *
 * 6分の実行時間制限で途中で止まった場合は、同じ関数をもう一度実行すれば
 * 続きから再開する。最初からやり直したいときは resetProgress() を実行する。
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

  // 1回の実行でここまで来たら中断し、次の実行で続きから再開する。
  // 実行時間の上限は個人アカウントで6分、Workspace で30分。
  TIME_LIMIT_MS: 240000,

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
  runOverBooks(cleanOpts('cleanupExtraRows'), function (ss, log) {
    let tabs = 0, rows = 0;
    const bookPeriod = readBookPeriod(ss);
    ss.getSheets().forEach(function (sh) {
      const grid = sheetGrid(sh);              // 1タブ1回だけ読む
      const at = readAttendanceRows(sh, grid);
      if (!at) return;
      const p = tabPeriod(sh, at, bookPeriod, log, grid);
      if (!p) return;

      const first = at.dataStart + p.days;
      if (first > at.dataEnd) return;                 // ちょうど埋まっている月

      const n = at.dataEnd - first + 1;
      let empty = true;                               // 中身の判定は grid で済ませる
      for (let r = first; r <= at.dataEnd && empty; r++) {
        for (let c = CLEAN.FIRST_COL; c <= CLEAN.LAST_COL; c++) {
          if (gv(grid, r, c) !== '') { empty = false; break; }
        }
      }
      if (empty) return;

      if (tabs === 0) {
        log.push('  期間 ' + fmtDate(p.start) + '〜' + fmtDate(p.end) + '（' + p.days + '日）'
          + ' → R' + first + '〜R' + at.dataEnd + ' の ' + n + '行が期間外');
      }
      tabs += 1; rows += n;
      if (!CLEAN.DRY_RUN) {
        sh.getRange(first, CLEAN.FIRST_COL, n, CLEAN.LAST_COL - CLEAN.FIRST_COL + 1).clearContent();
      }
    });
    log.push('- 期間外の行を消したタブ: ' + tabs + '（計 ' + rows + '行）');
  });
}


/**
 * 日付欄（B列）を集計期間から入れ直す。
 * 曜（C列）が集計期間と食い違うタブは、構成が想定と違うので手を付けずに飛ばす。
 */
function fixDateColumn() {
  runOverBooks(cleanOpts('fixDateColumn'), function (ss, log) {
    let tabs = 0, cells = 0;
    let sample = '';
    const bookPeriod = readBookPeriod(ss);
    ss.getSheets().forEach(function (sh) {
      const grid = sheetGrid(sh);              // 1タブ1回だけ読む
      const at = readAttendanceRows(sh, grid);
      if (!at) return;
      const p = tabPeriod(sh, at, bookPeriod, log, grid);
      if (!p) return;

      const cur = [];
      let changed = 0;
      for (let i = 0; i < p.days; i++) {
        const now = gv(grid, at.dataStart + i, CLEAN.FIRST_COL);
        const want = addDays(p.start, i);
        cur.push([dateKey(now) === dateKey(want) ? now : want]);
        if (dateKey(now) === dateKey(want)) continue;
        if (!sample) {
          sample = 'R' + (at.dataStart + i) + ' ' + describe(now) + ' → ' + fmtDate(want);
        }
        changed += 1;
      }
      if (!changed) return;

      tabs += 1; cells += changed;
      if (!CLEAN.DRY_RUN) {
        sh.getRange(at.dataStart, CLEAN.FIRST_COL, p.days, 1).setValues(cur);
      }
    });
    if (sample) log.push('  例: ' + sample);
    log.push('- 日付を直したタブ: ' + tabs + '（計 ' + cells + 'セル）');
  });
}


/* ------------------------------------------------------------------ */

function cleanOpts(tag) {
  return {
    dryRun: CLEAN.DRY_RUN,
    folderIds: CLEAN.FOLDER_IDS,
    spreadsheetIds: CLEAN.SPREADSHEET_IDS,
    timeLimitMs: CLEAN.TIME_LIMIT_MS,
    tag: tag,
  };
}

function describe(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return fmtDate(v);
  return v === '' || v === null ? '(空)' : '「' + v + '」';
}
