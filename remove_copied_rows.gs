/**
 * コピー元から持ち込まれた明細行を消すスクリプト（単独で動きます）
 *
 * 対象：22期_PA勤怠フォルダの給与一覧スプレッドシート（2026/9/16〜2027/9/15 の12本）
 *
 * 新しい人の出勤簿タブを既存の人のコピーで作った結果、コピー元にしかない
 * 「海事・横河分」「合算支給額」の行がそのまま残っている。
 * 岡田 梨沙のタブは衣幡千明のコピーなので、
 *   R49 海事・横河分 = ='衣幡千明_海事横河'!G46   ← 他人の金額を拾っている
 *   R50 合算支給額   = =G48+G49
 * となっていて、岡田さんの合算支給額に衣幡さんの海事横河分が足されている。
 *
 * このファイルは attendance_lib.gs などを必要としない。
 * 他のファイルと関数名がぶつからないよう、補助関数は rcd で始める名前にしてある。
 * ただし removeCopiedDetailRows だけは同名なので、summary_fix.gs を
 * 同じプロジェクトに入れている場合はそちらを削除すること。
 *
 * 使い方
 *   1. Apps Script で新しいファイルを作ってこれを貼る
 *   2. RCD.DRY_RUN = true のまま removeCopiedDetailRows() を実行し、
 *      実行ログで「どのタブのどの行を消すか」を確認する
 *   3. 問題なければ RCD.DRY_RUN = false にして再実行
 *
 * 6分の実行時間制限で途中で止まった場合は、もう一度実行すれば続きから再開する。
 * 最初からやり直したいときは rcdResetProgress() を実行する。
 */

const RCD = {
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

  // これらの行を持っていてよい人（2か所で働くため本来必要）
  KEEP_NAMES: ['衣幡千明'],

  // 消す対象の項目名。B列がこの文字列に一致する行の B〜I を消す
  TARGET_LABELS: ['海事・横河分', '合算支給額'],
};


function removeCopiedDetailRows() {
  const log = [];
  const books = rcdBooks(log);
  if (!books.length) { Logger.log('対象スプレッドシートが見つかりません'); return; }

  const props = PropertiesService.getScriptProperties();
  const KEY = 'rcd:progress';
  let from = Number(props.getProperty(KEY) || 0);
  if (from >= books.length) from = 0;
  if (from > 0) log.push('前回の続き（' + (from + 1) + '本目）から再開します');

  const keep = {};
  RCD.KEEP_NAMES.forEach(function (n) { keep[rcdNorm(n)] = true; });

  const t0 = Date.now();
  let i = from, total = 0;

  for (; i < books.length; i++) {
    // 1本目は必ず処理する。2本目以降は残り時間を見て打ち切る
    if (i > from && Date.now() - t0 > RCD.TIME_LIMIT_MS) break;

    let ss;
    try {
      ss = SpreadsheetApp.openById(books[i].id);
    } catch (e) {
      log.push('!! 開けません ' + books[i].name + ' : ' + e.message);
      continue;
    }

    log.push('');
    log.push('========== [' + (i + 1) + '/' + books.length + '] ' + ss.getName() + ' ==========');
    let hits = 0, checked = 0;

    ss.getSheets().forEach(function (sh) {
      // 往復を増やさないため、タブの中身は1回だけ読む
      const grid = sh.getDataRange().getValues();

      const staff = rcdStaffName(grid);
      if (!staff) return;                      // 出勤簿タブではない
      checked += 1;
      if (keep[rcdNorm(staff)]) return;        // 2か所で働く人は対象外

      for (let r = 0; r < grid.length; r++) {
        const label = rcdText(grid[r][1]);     // B列
        if (RCD.TARGET_LABELS.indexOf(label) < 0) continue;

        const row = r + 1;
        const formula = sh.getRange(row, 7).getFormula();   // G列。証拠としてログに残す
        log.push('  · ' + sh.getName() + ' R' + row + ' 「' + label + '」を削除'
          + '（G' + row + ' = ' + (formula || rcdText(grid[r][6]) || '空') + '）');
        hits += 1;
        if (!RCD.DRY_RUN) sh.getRange(row, 2, 1, 8).clearContent();
      }
    });

    log.push('- 出勤簿 ' + checked + 'タブを確認し、' + hits + '行が該当');
    total += hits;
    SpreadsheetApp.flush();
  }

  log.push('');
  if (i < books.length) {
    props.setProperty(KEY, String(i));
    log.push('*** 時間切れのため ' + i + '本目までで中断しました。');
    log.push('    もう一度 removeCopiedDetailRows() を実行すると ' + (i + 1) + '本目から続きます。***');
  } else {
    props.deleteProperty(KEY);
    log.push('*** 全' + books.length + '本 完了。該当行は合計 ' + total + '行 ***');
  }
  log.push(RCD.DRY_RUN
    ? '*** DRY_RUN です。書き込みは行っていません。***'
    : '*** 書き込みを実行しました。***');
  Logger.log(log.join('\n'));
}


/** 進捗の記録を消して、次の実行を1本目からにする */
function rcdResetProgress() {
  PropertiesService.getScriptProperties().deleteProperty('rcd:progress');
  Logger.log('進捗の記録を消しました。次の実行は1本目から始まります。');
}


/* ------------------------------------------------------------------ */

/** 対象スプレッドシートを列挙する */
function rcdBooks(log) {
  if (RCD.SPREADSHEET_IDS.length) {
    return RCD.SPREADSHEET_IDS.map(function (id) { return { id: id, name: id }; });
  }
  const out = [];
  RCD.FOLDER_IDS.forEach(function (fid) {
    const it = DriveApp.getFolderById(fid).getFilesByType(MimeType.GOOGLE_SHEETS);
    while (it.hasNext()) {
      const f = it.next();
      if (f.getName().indexOf('給与一覧') !== 0) continue;   // 給与一覧以外は触らない
      out.push({ id: f.getId(), name: f.getName() });
    }
  });
  out.sort(function (a, b) { return a.name < b.name ? -1 : 1; });
  log.push('対象スプレッドシート ' + out.length + '本');
  return out;
}

/**
 * 1行目の見出しから氏名を取る。出勤簿タブでなければ空文字。
 * 見出しは A1 か B1 のどちらかに入っているので両方つないでから調べる。
 */
function rcdStaffName(grid) {
  if (!grid.length) return '';
  const title = rcdText(grid[0][0]) + rcdText(grid[0][1]);
  const m = title.match(/出勤簿[\s　]*(.+?)（パート・アルバイト）/);
  return m ? m[1] : '';
}

/** 氏名の表記ゆれ（全角・半角スペース）を吸収する */
function rcdNorm(s) {
  return String(s).replace(/[\s　]+/g, '');
}

function rcdText(v) {
  return String(v === undefined || v === null ? '' : v).trim();
}
