/**
 * 給与一覧と出勤簿のコピー崩れを直すスクリプト
 *
 * 対象：22期_PA勤怠フォルダの給与一覧スプレッドシート（2026/9/16〜2027/9/15 の12本）
 * 必要ファイル：attendance_lib.gs
 *
 * 新しい人の行やタブを既存の人のコピーで作った結果、参照や書式が
 * ずれたままになっているものを直す。
 *
 *   repointSummary()
 *     給与一覧タブの C〜F 列を、その人自身の出勤簿タブにつなぎ直す。
 *     行をコピーして挿入したため相対参照が1行下にずれており、たとえば
 *     岡田 梨沙の行が '衣幡千明'!D38、コイララ・マダバの行が
 *     '村田耕一'!D38 を見ていて、C〜F が空欄になっている。
 *
 *   removeCopiedDetailRows()
 *     コピー元から持ち込まれた「海事・横河分」「合算支給額」の行を消す。
 *     岡田 梨沙のタブは衣幡千明のコピーなので、R49 が
 *     ='衣幡千明_海事横河'!G46 のままになっている（他人の金額を拾っている）。
 *
 *   fixDetailColors()
 *     給与明細ブロックの背景色が抜けているタブを、同じ明細構成の他タブに合わせる。
 *     衣幡千明_海事横河は見出しが白背景・白文字で読めない状態になっている。
 *
 * 使い方
 *   1. attendance_lib.gs と一緒に Apps Script プロジェクトに追加
 *   2. FIX.DRY_RUN = true のまま実行し、ログで内容を確認
 *   3. 問題なければ FIX.DRY_RUN = false にして再実行
 *
 * 6分の実行時間制限で途中で止まった場合は、同じ関数をもう一度実行すれば
 * 続きから再開する。最初からやり直したいときは resetProgress() を実行する。
 */

const FIX = {
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

  // 「海事・横河分」「合算支給額」を持ってよいタブ（2か所で働く人）
  COMBINED_NAMES: ['衣幡千明'],

  // コピーで持ち込まれる、差引支給額より下の行
  COPIED_LABELS: ['海事・横河分', '合算支給額'],

  // true にすると、背景色が「一部だけ」見本と違うタブも見本に揃える。
  // 既定は false（全部白になっているタブだけを直す）。
  NORMALIZE_ALL: false,
};

// 給与一覧タブの列
const SUM_COL = { NO: 1, NAME: 2, DAYS: 3, HOURS: 4, NIGHT: 5, PAY: 6, NOTE: 7 };

function fixOpts(tag) {
  return {
    dryRun: FIX.DRY_RUN,
    folderIds: FIX.FOLDER_IDS,
    spreadsheetIds: FIX.SPREADSHEET_IDS,
    timeLimitMs: FIX.TIME_LIMIT_MS,
    tag: tag,
  };
}


/* ------------------------------------------------------------------ *
 * 1. 給与一覧の C〜F を本人の出勤簿につなぎ直す
 * ------------------------------------------------------------------ */
function repointSummary() {
  runOverBooks(fixOpts('repointSummary'), function (ss, log) {
    const summary = getSummarySheet(ss);
    if (!summary) { log.push('!! 給与一覧タブが見つかりません'); return; }

    const grid = sheetGrid(summary);
    const index = attendanceIndex(ss);
    const members = readSummaryMembers(summary, grid);

    // C〜F の数式はまとめて1回で読む
    const formulas = summary.getRange(1, SUM_COL.DAYS, grid.length, 4).getFormulas();
    let fixed = 0;

    members.forEach(function (m) {
      const hit = index[normalizeName(m.name)];
      if (!hit) { log.push('  ? 出勤簿タブが見つかりません: R' + m.row + ' ' + m.name); return; }

      const q = "'" + hit.sheet.getName().replace(/'/g, "''") + "'!";
      const want = [
        '=' + q + 'D' + hit.at.totalRow,   // 出勤日数
        '=' + q + 'G' + hit.at.totalRow,   // 実働(h)
        '=' + q + 'H' + hit.at.totalRow,   // 深夜(h)
        '=' + q + 'G' + hit.at.netRow,     // 差引支給額
      ];
      const now = formulas[m.row - 1];
      if (now.join(' ') === want.join(' ')) return;

      log.push('  · R' + m.row + ' ' + m.name + ': ' + (now[0] || '(値)') + ' → ' + want[0]
        + ' ほか（' + hit.at.netLabel + ' = G' + hit.at.netRow + '）');
      fixed += 1;
      if (!FIX.DRY_RUN) summary.getRange(m.row, SUM_COL.DAYS, 1, 4).setFormulas([want]);
    });

    log.push('- 給与一覧でつなぎ直した行: ' + fixed + ' / ' + members.length);
  });
}


/* ------------------------------------------------------------------ *
 * 2. コピー元から持ち込まれた明細行を消す
 * ------------------------------------------------------------------ */
function removeCopiedDetailRows() {
  const allowed = {};
  FIX.COMBINED_NAMES.forEach(function (n) { allowed[normalizeName(n)] = true; });

  runOverBooks(fixOpts('removeCopiedDetailRows'), function (ss, log) {
    const index = attendanceIndex(ss);
    let hits = 0;

    Object.keys(index).forEach(function (name) {
      if (allowed[name]) return;                   // 2か所で働く人は対象外
      const sh = index[name].sheet;
      const at = index[name].at;
      const grid = index[name].grid;

      // 差引支給額より下を数行だけ見る（読み取りは attendanceIndex の1回で済んでいる）
      const from = at.netRow + 1;
      const to = Math.min(at.netRow + FIX.COPIED_LABELS.length + 2, grid.length);

      for (let row = from; row <= to; row++) {
        const label = String(gv(grid, row, 2)).trim();
        if (FIX.COPIED_LABELS.indexOf(label) < 0) continue;

        const formula = sh.getRange(row, 7).getFormula();
        log.push('  · ' + sh.getName() + ' R' + row + ' 「' + label + '」を削除'
          + (formula ? '（' + formula + '）' : ''));
        hits += 1;
        if (!FIX.DRY_RUN) sh.getRange(row, 2, 1, 8).clearContent();
      }
    });

    log.push('- 消した明細行: ' + hits);
  });
}


/* ------------------------------------------------------------------ *
 * 3. 給与明細ブロックの背景色を他タブに合わせる
 * ------------------------------------------------------------------ */
function fixDetailColors() {
  runOverBooks(fixOpts('fixDetailColors'), function (ss, log) {
    const index = attendanceIndex(ss);
    const groups = {};

    // 明細の項目名の並びが同じタブごとにまとめ、それぞれの背景色を読む
    Object.keys(index).forEach(function (name) {
      const sh = index[name].sheet, at = index[name].at, grid = index[name].grid;
      if (!at.detailStart) return;
      const rows = at.netRow - at.detailStart + 1;
      const bg = sh.getRange(at.detailStart, 2, rows, 8).getBackgrounds();
      let colored = 0;
      bg.forEach(function (r) { r.forEach(function (c) { if (!isWhite(c)) colored += 1; }); });
      const sig = detailSignature(grid, at);
      (groups[sig] = groups[sig] || []).push({
        sheet: sh, at: at, rows: rows, bg: bg, colored: colored, key: JSON.stringify(bg),
      });
    });

    let fixed = 0;
    const partial = [];

    Object.keys(groups).forEach(function (sig) {
      const tabs = groups[sig];

      // 見本は「色が最も多く入っているパターン」。同数ならタブ数が多い方。
      // 単純な多数決にはしない。7行構成のタブは、色が10セルのものと21セルのものが
      // 10タブずつあり、多数決だと色の少ない方を見本にしてしまうため。
      const byKey = {};
      tabs.forEach(function (t) {
        const e = byKey[t.key] || (byKey[t.key] = { colored: t.colored, count: 0, sample: t });
        e.count += 1;
      });
      let best = null;
      Object.keys(byKey).forEach(function (k) {
        const e = byKey[k];
        if (!best || e.colored > best.colored
            || (e.colored === best.colored && e.count > best.count)) best = e;
      });
      if (!best || best.colored === 0) return;   // 色付きの見本がない

      tabs.forEach(function (t) {
        if (t.key === best.sample.key) return;

        if (t.colored > 0 && !FIX.NORMALIZE_ALL) {
          partial.push(t.sheet.getName());       // 一部だけ違うタブは既定では触らない
          return;
        }
        log.push('  · ' + t.sheet.getName() + ' R' + t.at.detailStart + '〜R' + t.at.netRow
          + ' の背景色を ' + best.sample.sheet.getName() + ' に合わせる'
          + (t.colored === 0 ? '（今は全部白）' : '（一部だけ違う）'));
        fixed += 1;
        if (FIX.DRY_RUN) return;

        // 罫線・文字色・数式には触れず、背景色だけを写す
        t.sheet.getRange(t.at.detailStart, 2, t.rows, 8).setBackgrounds(best.sample.bg);
      });
    });

    if (partial.length) {
      log.push('  ? 見本と一部だけ違うタブ（今回は変更しません。FIX.NORMALIZE_ALL = true で揃います）: '
        + partial.length + '件 … ' + partial.slice(0, 6).join('、'));
    }
    log.push('- 背景色を直したタブ: ' + fixed);
  });
}

/** 明細の項目名を縦に連結したもの。同じ構成のタブを見分けるのに使う */
function detailSignature(grid, at) {
  const out = [];
  for (let r = at.detailStart; r <= at.netRow; r++) out.push(String(gv(grid, r, 2)).trim());
  return out.join('|');
}

function isWhite(bg) {
  const c = String(bg).toLowerCase();
  return c === '' || c === '#ffffff' || c === '#fff' || c === 'white';
}
