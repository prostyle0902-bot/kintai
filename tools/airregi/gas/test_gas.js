/**
 * Code.gs の動作確認（Node.js で実行）。
 *
 *     node tools/airregi/gas/test_gas.js
 *
 * Googleのサービス（DriveApp / Utilities / HtmlService など）を差し替えて、
 * Code.gs をそのまま読み込んで動かす。2026年7月の実績値を再現できるか検証する。
 *
 * ※ Shift-JISの読み取りは Python 版のテスト（tools/airregi/tests/test_report.py）で
 *    検証しているので、ここではUTF-8のCSVを使う。
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const assert = require('assert');

// --------------------------------------------------------- テスト用のCSV

const HEADER = [
  '集計日', '売上', '客数', '客単価', '組数', '組単価', '商品点数',
  '現金売上高', '現金売上金額(10%対象)', '現金売上金額(8%軽減)',
  'その他売上高', 'その他売上金額(10%対象)', '値引額', '売上金額(10%対象)',
  'クレジットカード(Airペイ) 売上金額',
];

// (日, 売上, 客数, 組数, 商品点数, 現金, その他)
const JULY = {
  'タコハイ7月': [
    [2, 12350, 7, 8, 20, 11700, 650], [3, 74380, 30, 38, 108, 46650, 27730],
    [6, 29450, 17, 17, 41, 19550, 9900], [7, 20450, 13, 13, 29, 12000, 8450],
    [9, 53780, 18, 21, 75, 17750, 36030], [10, 55280, 18, 28, 85, 45150, 10130],
    [11, 58420, 32, 37, 87, 35550, 22870], [12, 31200, 15, 20, 50, 20400, 10800],
    [13, 46680, 20, 22, 61, 10050, 36630], [14, 32440, 19, 19, 47, 15200, 17240],
    [16, 28140, 17, 19, 47, 11350, 16790], [17, 46640, 21, 26, 70, 10000, 36640],
    [18, 61660, 33, 41, 90, 29100, 32560], [19, 69480, 49, 52, 95, 32450, 37030],
    [20, 38000, 30, 31, 60, 13350, 24650], [21, 21340, 16, 16, 31, 6950, 14390],
    [23, 24980, 15, 21, 37, 21350, 3630], [24, 35920, 10, 20, 53, 26900, 9020],
    [25, 64610, 32, 36, 95, 28600, 36010], [26, 26670, 11, 18, 41, 15000, 11670],
    [27, 45940, 16, 27, 69, 39700, 6240], [28, 23460, 12, 14, 32, 11050, 12410],
    [30, 51700, 17, 21, 66, 30600, 21100], [31, 42140, 17, 29, 66, 21900, 20240],
  ],
};

// 7/4・7/5 も営業している店。これがあることで タコハイ の 7/4・7/5 が臨時休業として拾われる。
// 営業日は タコハイ と同じ日 ＋ 4日・5日（＝水曜だけが全店共通の定休日）。
JULY['となり店7月'] = JULY['タコハイ7月']
  .map(([d]) => [d, 30000, 10, 12, 30, 20000, 10000])
  .concat([[4, 50000, 20, 25, 60, 30000, 20000], [5, 45000, 18, 22, 55, 28000, 17000]])
  .sort((a, b) => a[0] - b[0]);

// タコハイの6月（合計 852,270円 / 485組）
const JUNE = {
  'タコハイ6月': [
    [1, 16250, 9, 10, 30, 10000, 6250], [2, 18050, 12, 14, 34, 12000, 6050],
    [4, 25980, 11, 12, 40, 15980, 10000], [5, 30640, 20, 22, 55, 20640, 10000],
    [6, 65580, 30, 33, 120, 40580, 25000], [7, 39000, 22, 25, 70, 24000, 15000],
    [8, 36070, 14, 16, 60, 21070, 15000], [9, 31440, 15, 17, 52, 19440, 12000],
    [11, 32280, 17, 19, 55, 20280, 12000], [12, 39970, 20, 22, 66, 24970, 15000],
    [13, 37920, 21, 24, 64, 22920, 15000], [14, 33830, 20, 22, 58, 20830, 13000],
    [15, 27200, 19, 22, 48, 17200, 10000], [16, 36140, 14, 16, 58, 22140, 14000],
    [18, 37890, 21, 24, 62, 23890, 14000], [19, 40520, 24, 27, 68, 25520, 15000],
    [20, 46220, 24, 27, 78, 28220, 18000], [21, 40730, 15, 17, 66, 25730, 15000],
    [22, 19630, 8, 9, 32, 12630, 7000], [23, 23660, 11, 13, 40, 14660, 9000],
    [25, 18320, 10, 11, 30, 11320, 7000], [26, 32090, 15, 17, 54, 20090, 12000],
    [27, 32530, 18, 20, 55, 20530, 12000], [28, 29580, 13, 15, 48, 18580, 11000],
    [29, 22400, 8, 9, 36, 14400, 8000], [30, 38350, 20, 22, 62, 23350, 15000],
  ],
};

function csvText(month, rows) {
  const lines = [HEADER.join(',')];
  rows.forEach(([d, sales, guests, groups, items, cash, other]) => {
    lines.push([
      `2026${String(month).padStart(2, '0')}${String(d).padStart(2, '0')}`,
      sales, guests, Math.round(sales / guests), groups, Math.round(sales / groups),
      items, cash, cash, 0, other, other, 0, sales, other,
    ].join(','));
  });
  return lines.join('\n');
}

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'airregi-'));
Object.keys(JULY).forEach((n) => fs.writeFileSync(path.join(dir, n + '.csv'), csvText(7, JULY[n])));
Object.keys(JUNE).forEach((n) => fs.writeFileSync(path.join(dir, n + '.csv'), csvText(6, JUNE[n])));

// --------------------------------------------------------- Googleのサービスを差し替え

function iter(arr) {
  let i = 0;
  return { hasNext: () => i < arr.length, next: () => arr[i++] };
}
global.DriveApp = {
  getFoldersByName: () => iter([{
    getFiles: () => iter(fs.readdirSync(dir).map((f) => ({
      getName: () => f,
      getBlob: () => ({
        getDataAsString: (charset) => {
          const buf = fs.readFileSync(path.join(dir, f));
          const enc = String(charset || 'UTF-8').toLowerCase().replace('_', '-');
          return new TextDecoder(enc === 'shift-jis' ? 'shift_jis' : 'utf-8').decode(buf);
        },
      }),
    }))),
  }]),
  getFolderById: () => { throw new Error('未使用'); },
};
global.Utilities = {
  parseCsv: (text) => {
    const rows = []; let row = [], cell = '', quoted = false;
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (quoted) {
        if (c === '"') { if (text[i + 1] === '"') { cell += '"'; i++; } else quoted = false; }
        else cell += c;
      } else if (c === '"') quoted = true;
      else if (c === ',') { row.push(cell); cell = ''; }
      else if (c === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; }
      else if (c !== '\r') cell += c;
    }
    if (cell || row.length) { row.push(cell); rows.push(row); }
    return rows;
  },
};
global.Logger = { log: () => {} };
global.HtmlService = {
  XFrameOptionsMode: { ALLOWALL: 1 },
  createHtmlOutput: (h) => ({
    setTitle() { return this; }, addMetaTag() { return this; },
    setXFrameOptionsMode() { return this; }, getContent: () => h,
  }),
};
global.ScriptApp = {
  getService: () => ({ getUrl: () => 'https://example.com/exec' }),
  getProjectTriggers: () => [],
};
global.Session = { getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' }) };
let sentMail = null;
global.MailApp = { sendEmail: (o) => { sentMail = o; } };

eval(fs.readFileSync(path.join(__dirname, 'Code.gs'), 'utf8'));

// --------------------------------------------------------- 検証

const tests = [];
function test(name, fn) { tests.push([name, fn]); }

test('CSVの日付から月を仕分けできる', () => {
  const all = loadAll_();
  assert.deepStrictEqual(all.months, ['2026-07', '2026-06']);
  assert.ok(all.stores['タコハイ'], '店舗名がファイル名から取れていない');
  assert.strictEqual(all.stores['タコハイ']['2026-07'].length, 24);
  assert.strictEqual(all.stores['タコハイ']['2026-06'].length, 26);
});

test('7月の実績値を再現できる', () => {
  const rep = buildReport_(loadAll_(), '2026-07');
  const taco = rep.stores.filter((s) => s.name === 'タコハイ')[0];
  assert.strictEqual(taco.sales, 995110);
  assert.strictEqual(taco.groups, 594);
  assert.strictEqual(taco.days, 24);
  assert.strictEqual(Math.round(taco.unit), 1675);
  assert.strictEqual(Math.round(taco.cashless * 1000), 465);       // 46.5%
  assert.strictEqual(Math.round(taco.guestInput * 100), 82);       // 人数入力率
});

test('曜日別の平均が正しい', () => {
  const rep = buildReport_(loadAll_(), '2026-07');
  const taco = rep.stores.filter((s) => s.name === 'タコハイ')[0];
  assert.strictEqual(Math.round(taco.weekday['土']), 61563);   // 7/11,18,25
  assert.strictEqual(taco.weekday['火'], 97690 / 4);           // 7/7,14,21,28
});

test('前月比と増減の分解が合う', () => {
  const rep = buildReport_(loadAll_(), '2026-07');
  const taco = rep.stores.filter((s) => s.name === 'タコハイ')[0];
  assert.strictEqual(taco.prev.sales, 852270);
  assert.ok(Math.abs(taco.mom - 1.1676) < 0.001, '前月比: ' + taco.mom);
  const diff = taco.sales - taco.prev.sales;
  assert.ok(Math.abs(taco.volumeEffect + taco.priceEffect - diff) < 1, '分解の合計が増減額と合わない');
  assert.ok(taco.volumeEffect > 0 && taco.priceEffect < 0);
});

test('他店が営業していた日だけ臨時休業として拾う', () => {
  const rep = buildReport_(loadAll_(), '2026-07');
  const taco = rep.stores.filter((s) => s.name === 'タコハイ')[0];
  const tonari = rep.stores.filter((s) => s.name === 'となり店')[0];
  assert.deepStrictEqual(taco.closed, [4, 5]);   // となり店だけが開けていた土日
  assert.strictEqual(tonari.closed.length, 0);   // 全店そろって休んだ日（水曜）は数えない
  assert.strictEqual(Math.round(taco.loss), 104013);
});

test('画面のHTMLが生成できる', () => {
  const out = doGet({ parameter: {} }).getContent();
  assert.ok(out.indexOf('¥995,110') >= 0, '店舗の売上が出ていない');
  assert.ok(out.indexOf('2026年7月') >= 0);
  assert.ok(out.indexOf('undefined') < 0, 'undefined が混ざっている');
  assert.ok(out.indexOf('NaN') < 0, 'NaN が混ざっている');
  const other = doGet({ parameter: { month: '2026-06' } }).getContent();
  assert.ok(other.indexOf('2026年6月') >= 0, '月の切り替えができていない');
});

test('毎月のメールが作れる', () => {
  sendMonthlyMail();
  assert.ok(sentMail, 'メールが送られていない');
  assert.ok(sentMail.subject.indexOf('2026年7月') >= 0, sentMail.subject);
  assert.ok(sentMail.htmlBody.indexOf('タコハイ') >= 0);
});

let failed = 0;
tests.forEach(([name, fn]) => {
  try {
    fn();
    console.log('  ok   ' + name);
  } catch (err) {
    failed++;
    console.log('  NG   ' + name + '\n       ' + err.message);
  }
});
fs.rmSync(dir, { recursive: true, force: true });
console.log(failed ? `\n${failed} 件失敗` : `\n${tests.length} 件すべて成功`);
process.exit(failed ? 1 : 0);
