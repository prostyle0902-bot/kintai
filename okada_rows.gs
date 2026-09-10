/**
 * 岡田 梨沙タブの余分な2行を消す（実行するとすぐ書き込みます）
 *
 * 岡田 梨沙のタブは衣幡千明のコピーで作られたため、衣幡さんにしかない
 * 「海事・横河分」「合算支給額」の2行が残っている。
 *   R49 海事・横河分 = ='衣幡千明_海事横河'!G46   ← 衣幡さんの28,080円を拾っている
 *   R50 合算支給額   = =G48+G49
 * このままだと岡田さんの支給額に他人の金額が乗る。
 *
 * 使い方
 *   Apps Script に貼って deleteOkadaExtraRows() を実行するだけ。
 *   確認用の空実行はありません。実行するとその場で消します。
 *   結果は「実行ログ」（表示 → ログ）に出ます。
 *
 * 行番号について
 *   多くのブックでは49行目と50行目ですが、明細の項目数が違うブックでは
 *   48行目と49行目になります（例：2027/2/16〜3/15 は「修正給与」の行がない）。
 *   番号を決め打ちすると取り違えるので、B列の項目名で探します。
 *   実際に消した行番号はログに出ます。
 */

// 22期の給与一覧スプレッドシート（2026/9/16〜 の12本）
const OKADA_BOOKS = [
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

// 消す対象のタブと項目名
const OKADA_TAB = '岡田 梨沙';
const OKADA_LABELS = ['海事・横河分', '合算支給額'];


function deleteOkadaExtraRows() {
  const log = [];
  let total = 0;

  for (let i = 0; i < OKADA_BOOKS.length; i++) {
    let ss;
    try {
      ss = SpreadsheetApp.openById(OKADA_BOOKS[i]);
    } catch (e) {
      log.push('[' + (i + 1) + '/12] !! 開けません: ' + e.message);
      continue;
    }

    const sh = okadaFindTab(ss);
    if (!sh) {
      log.push('[' + (i + 1) + '/12] ' + ss.getName() + ' … 「' + OKADA_TAB + '」タブなし');
      continue;
    }

    // 明細のあたり（40〜55行）のB列とG列だけ読む
    const from = 40, rows = 16;
    const items = sh.getRange(from, 2, rows, 1).getValues();   // B列＝項目名
    const money = sh.getRange(from, 7, rows, 1).getValues();   // G列＝金額

    const hits = [];
    for (let k = 0; k < rows; k++) {
      const label = String(items[k][0] === null ? '' : items[k][0]).trim();
      if (OKADA_LABELS.indexOf(label) < 0) continue;
      const row = from + k;
      hits.push('R' + row + '「' + label + '」' + money[k][0]);
      sh.getRange(row, 2, 1, 8).clearContent();                // B〜I列を消す
      total += 1;
    }

    log.push('[' + (i + 1) + '/12] ' + ss.getName() + ' … '
      + (hits.length ? '消した: ' + hits.join(' / ') : '該当なし（すでに消えています）'));
    SpreadsheetApp.flush();
  }

  log.push('');
  log.push('合計 ' + total + '行を消しました。');
  Logger.log(log.join('\n'));
}


/** 「岡田 梨沙」タブを探す。スペースの有無・全角半角の違いを吸収する */
function okadaFindTab(ss) {
  const want = OKADA_TAB.replace(/[\s　]+/g, '');
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().replace(/[\s　]+/g, '') === want) return sheets[i];
  }
  return null;
}
