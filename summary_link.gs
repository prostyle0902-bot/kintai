/**
 * 給与一覧タブの C〜F 列を本人の出勤簿につなぎ直す（実行するとすぐ書き込みます）
 *
 * 岡田 梨沙とコイララ・マダバは、既存の人の行をコピーして挿入したため
 * 相対参照が1行下にずれ、他人のタブを見ている。
 *   R41 岡田 梨沙      C='衣幡千明'!D38 … F='衣幡千明'!G49
 *   R64 コイララ・マダバ C='村田耕一'!D38 … F='村田耕一'!G47
 * 参照先が空欄や他人の金額なので、C〜F が正しく出ない。
 *
 * 正しい参照先は本人のタブの
 *   C 出勤日数 = 合計行の D
 *   D 実働(h)  = 合計行の G
 *   E 深夜(h)  = 合計行の H
 *   F 支給額   = 「差引支給額」の行の G
 * 行番号はブックごとに違う（明細の項目数が違う）ので、
 * そのつどタブの中を見て決める。書き込んだ式はログに出る。
 *
 * 使い方
 *   Apps Script に貼って relinkSummaryRows() を実行するだけ。
 *   確認用の空実行はありません。実行するとその場で書き換えます。
 *   結果は「実行ログ」（表示 → ログ）に出ます。
 */

// 22期の給与一覧スプレッドシート（2026/9/16〜 の12本）
const LINK_BOOKS = [
  '1Dxr27E2tGwiPKESxWWB8HCGDaX0i2jwtSqfL5Lyc9tw', // 2026/9/16-10/15
  '1SSRb1I2RD6cOkRmBqW0tcZBVIntf_4RHDzDpzqXHi4w', // 2026/10/16-11/15
  '1ARP6ApVaj7VzSKBuz0P5TcdnbH8BQJ5J5IzB_AQgmBw', // 2026/11/16-12/15
  '1MwAXLQfBF-tBQ0jMHyJqLaNlUSbImikp23IdvD9CDBs', // 2026/12/16-2027/1/15
  '1awRXj9U0C18D_gT0tM5u14AENE-zak4d6pafseX5gAM', // 2027/1/16-2/15
  '1rju65K08_cefudwVkKcfw4gJh4uPAWc8kDgpqmeFYC4', // 2027/2/16-3/15
  '18bugMUJZtAQucKRLLA8CAawEiw7IrndNGHhDBIfOpbs', // 2027/3/16-4/15
  '1QUPZqD496tJqtLUh04su1FmmzD4Q8GmcN0yTS9YhsD0', // 2027/4/16-5/15
  '1S5Oq5bHtxBM2JaQApzfkOk2BPu9_GN04MjVCGM5XDn8', // 2027/5/16-6/15
  '1K2q7smAXcqW0VF12FbSmo4vWSjOEwewYgAKfzX6dzWU', // 2027/6/16-7/15
  '1MrtFtRVKxG97mIHiRkX27fkVXxrQNYy9Y5-Q__oZE8Y', // 2027/7/16-8/15
  '1TrgIpopQedhDhzCyx4I0-uqg38ravAP0lw1HnBd3Y6o', // 2027/8/16-9/15
];

// つなぎ直す人。給与一覧のB列の氏名と、同名の出勤簿タブを対応させる
const LINK_TARGETS = ['岡田 梨沙', 'コイララ・マダバ'];


function relinkSummaryRows() {
  const log = [];
  let total = 0;

  for (let i = 0; i < LINK_BOOKS.length; i++) {
    const tag = '[' + (i + 1) + '/' + LINK_BOOKS.length + '] ';
    let ss;
    try {
      ss = SpreadsheetApp.openById(LINK_BOOKS[i]);
    } catch (e) {
      log.push(tag + '!! 開けません: ' + e.message);
      continue;
    }

    const sheets = ss.getSheets();
    const summary = linkFind(sheets, '給与一覧');
    if (!summary) { log.push(tag + ss.getName() + ' … 給与一覧タブなし'); continue; }

    // 給与一覧のB列（氏名）をまとめて1回で読む
    const names = summary.getRange(1, 2, Math.min(summary.getLastRow(), 120), 1).getValues();
    log.push(tag + ss.getName());

    LINK_TARGETS.forEach(function (target) {
      const row = linkRowOf(names, target);
      if (!row) { log.push('    · ' + target + ' … 給与一覧に行なし'); return; }

      const sh = linkFind(sheets, target);
      if (!sh) { log.push('    · ' + target + ' … 出勤簿タブなし（R' + row + ' は触りません）'); return; }

      // 本人のタブから「合計」行と「差引支給額」行を探す
      const items = sh.getRange(1, 2, Math.min(sh.getLastRow(), 60), 1).getValues();
      let totalRow = 0, netRow = 0, netLabel = '';
      for (let k = 0; k < items.length; k++) {
        const v = String(items[k][0] === null ? '' : items[k][0]).trim();
        if (!totalRow && v === '合計') totalRow = k + 1;
        if (!netRow && v.indexOf('差引支給額') === 0) { netRow = k + 1; netLabel = v; }
      }
      if (!totalRow || !netRow) {
        log.push('    · ' + target + ' … タブの「合計」か「差引支給額」が見つからず中止');
        return;
      }

      const q = "'" + sh.getName().replace(/'/g, "''") + "'!";
      const want = [
        '=' + q + 'D' + totalRow,   // C 出勤日数
        '=' + q + 'G' + totalRow,   // D 実働(h)
        '=' + q + 'H' + totalRow,   // E 深夜(h)
        '=' + q + 'G' + netRow,     // F 差引支給額
      ];
      const before = summary.getRange(row, 3, 1, 4).getFormulas()[0];
      summary.getRange(row, 3, 1, 4).setFormulas([want]);
      total += 1;

      log.push('    · R' + row + ' ' + target + '  ' + (before[0] || '(値)') + ' → ' + want[0]);
      log.push('                F: ' + (before[3] || '(値)') + ' → ' + want[3]
        + '（' + netLabel + '）');
    });

    SpreadsheetApp.flush();
  }

  log.push('');
  log.push('合計 ' + total + '行をつなぎ直しました。');
  Logger.log(log.join('\n'));
}


/** タブ名で探す。スペースの有無・全角半角の違いを吸収する */
function linkFind(sheets, name) {
  const want = String(name).replace(/[\s　]+/g, '');
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().replace(/[\s　]+/g, '') === want) return sheets[i];
  }
  return null;
}

/** 給与一覧のB列から氏名の行番号を探す */
function linkRowOf(names, name) {
  const want = String(name).replace(/[\s　]+/g, '');
  for (let i = 0; i < names.length; i++) {
    const v = String(names[i][0] === null ? '' : names[i][0]).replace(/[\s　]+/g, '');
    if (v === want) return i + 1;
  }
  return 0;
}
