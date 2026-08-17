/**
 * Airレジ 月次売上レポート（Google Apps Script 版）
 *
 * Googleドライブの「Airレジ売上CSV」フォルダに入っているCSVを読んで、
 * スマホで見られる売上レポートを表示する。毎月1日にメールでも届く。
 *
 * 設置のしかたは同じフォルダの README.md を参照。
 */

// ---------------------------------------------------------------- 設定

/** CSVを入れてあるGoogleドライブのフォルダ名 */
var FOLDER_NAME = 'Airレジ売上CSV';

/** フォルダ名で見つからないとき用。フォルダのURLの「/folders/」の後ろの文字列を入れる */
var FOLDER_ID = '';

/** レポートに出したくない店舗があればファイル名の一部を書く（例: ['十三里屋']） */
var EXCLUDE = [];

/** 毎月のメールの宛先。空ならこのスクリプトの持ち主に送る */
var MAIL_TO = '';

var WEEK = ['日', '月', '火', '水', '木', '金', '土'];

// ---------------------------------------------------------------- 入口

function doGet(e) {
  try {
    var all = loadAll_();
    if (all.months.length === 0) {
      return html_('<div class="wrap"><p class="empty">CSVが見つかりませんでした。<br>'
        + 'Googleドライブの「' + esc_(FOLDER_NAME) + '」フォルダに、Airレジからダウンロードした'
        + 'CSVを入れてください。</p></div>');
    }
    var month = (e && e.parameter && e.parameter.month) ? e.parameter.month : all.months[0];
    if (all.months.indexOf(month) < 0) month = all.months[0];
    return html_(renderReport_(buildReport_(all, month), all.months, month));
  } catch (err) {
    return html_('<div class="wrap"><p class="empty">エラーが起きました。<br>'
      + esc_(String(err)) + '</p></div>');
  }
}

function html_(body) {
  return HtmlService.createHtmlOutput(PAGE_HEAD + body + '</body></html>')
    .setTitle('Airレジ 売上レポート')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------------------------------------------------------------- 読み込み

function folder_() {
  if (FOLDER_ID) return DriveApp.getFolderById(FOLDER_ID);
  var it = DriveApp.getFoldersByName(FOLDER_NAME);
  if (!it.hasNext()) throw new Error('フォルダ「' + FOLDER_NAME + '」が見つかりません');
  return it.next();
}

/** フォルダ内の全CSVを読み、店舗ごと・月ごとに仕分けする */
function loadAll_() {
  var files = folder_().getFiles();
  var stores = {};      // 店舗名 -> { '2026-07': [行, ...] }
  var monthSet = {};
  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();
    if (!/\.csv$/i.test(name)) continue;
    if (EXCLUDE.some(function (w) { return name.indexOf(w) >= 0; })) continue;
    var store = storeName_(name);
    var rows = parseCsv_(file);
    for (var i = 0; i < rows.length; i++) {
      var ym = rows[i].ym;
      monthSet[ym] = true;
      if (!stores[store]) stores[store] = {};
      if (!stores[store][ym]) stores[store][ym] = [];
      stores[store][ym].push(rows[i]);
    }
  }
  var months = Object.keys(monthSet).sort().reverse();   // 新しい月が先頭
  return { stores: stores, months: months };
}

/** ファイル名から店舗名を取る（「タコハイ７月.csv」→「タコハイ」） */
function storeName_(fileName) {
  var s = fileName.replace(/\.csv$/i, '');
  s = s.replace(/[０-９]/g, function (c) {
    return String.fromCharCode(c.charCodeAt(0) - 0xFEE0);
  });
  s = s.replace(/[_\s-]*\d{4}[-_年]?\d{1,2}月?$/, '');
  s = s.replace(/[_\s-]*\d{1,2}月$/, '');
  return s.trim() || fileName;
}

/**
 * CSVを読む。Airレジのエクスポートは店舗の設定で列数が変わるので、
 * 列の位置ではなくヘッダー名で必要な列を探す。文字コードはShift-JIS優先。
 */
function parseCsv_(file) {
  var blob = file.getBlob();
  var text = '';
  try {
    text = blob.getDataAsString('Shift_JIS');
  } catch (err) {
    text = '';
  }
  if (text.indexOf('売上') < 0) {
    try {
      text = blob.getDataAsString('UTF-8');
    } catch (err2) { /* そのまま */ }
  }
  if (!text) return [];

  var table = Utilities.parseCsv(text);
  if (table.length < 2) return [];
  var col = resolveColumns_(table[0]);
  if (col.date == null || col.sales == null) return [];

  var out = [];
  for (var r = 1; r < table.length; r++) {
    var row = table[r];
    if (!row || !row[col.date]) continue;
    var m = String(row[col.date]).replace(/\s/g, '').match(/^(\d{4})[-\/]?(\d{2})[-\/]?(\d{2})/);
    if (!m) continue;
    var y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
    var pick = function (key) { return col[key] == null ? 0 : num_(row[col[key]]); };
    var groups = pick('groups') || pick('guests');
    out.push({
      ym: y + '-' + pad_(mo),
      day: d,
      dow: WEEK[new Date(y, mo - 1, d).getDay()],
      sales: pick('sales'),
      guests: pick('guests'),
      groups: groups,
      items: pick('items'),
      cash: pick('cash'),
      other: pick('other')
    });
  }
  out.sort(function (a, b) { return a.day - b.day; });
  return out;
}

var COLUMN_KEYS = [
  ['date', ['集計日', '日付']],
  ['sales', ['売上']],
  ['guests', ['客数']],
  ['groups', ['組数']],
  ['items', ['商品点数', '点数']],
  ['cash', ['現金売上高', '現金売上']],
  ['other', ['その他売上高', 'その他売上']]
];

function resolveColumns_(header) {
  var col = {};
  for (var i = 0; i < header.length; i++) {
    var n = String(header[i]).replace(/[　\s]/g, '').trim();
    for (var k = 0; k < COLUMN_KEYS.length; k++) {
      var key = COLUMN_KEYS[k][0], cands = COLUMN_KEYS[k][1];
      if (col[key] != null) continue;
      for (var c = 0; c < cands.length; c++) {
        if (n === cands[c] || n.indexOf(cands[c] + '(') === 0 || n.indexOf(cands[c] + '（') === 0) {
          col[key] = i;
          break;
        }
      }
    }
  }
  return col;
}

function num_(v) {
  var s = String(v == null ? '' : v).replace(/[,¥\s"]/g, '');
  if (!s || s === '-') return 0;
  var n = Number(s);
  return isNaN(n) ? 0 : Math.round(n);
}

// ---------------------------------------------------------------- 集計

function summarize_(rows) {
  var s = { sales: 0, guests: 0, groups: 0, items: 0, cash: 0, other: 0, days: rows.length };
  for (var i = 0; i < rows.length; i++) {
    s.sales += rows[i].sales; s.guests += rows[i].guests; s.groups += rows[i].groups;
    s.items += rows[i].items; s.cash += rows[i].cash; s.other += rows[i].other;
  }
  s.daily = s.days ? s.sales / s.days : 0;
  s.unit = s.groups ? s.sales / s.groups : 0;             // 組単価（会計1件あたり）
  s.perGroup = s.groups ? s.items / s.groups : 0;
  s.cashless = (s.cash + s.other) ? s.other / (s.cash + s.other) : 0;
  s.guestInput = s.groups ? s.guests / s.groups : 0;      // 人数入力率
  return s;
}

function byWeekday_(rows) {
  var acc = {};
  for (var i = 0; i < rows.length; i++) {
    var w = rows[i].dow;
    if (!acc[w]) acc[w] = { sum: 0, n: 0 };
    acc[w].sum += rows[i].sales; acc[w].n++;
  }
  var avg = {};
  Object.keys(acc).forEach(function (w) { avg[w] = acc[w].sum / acc[w].n; });
  return avg;
}

function prevMonth_(ym) {
  var y = Number(ym.slice(0, 4)), m = Number(ym.slice(5));
  m--; if (m === 0) { m = 12; y--; }
  return y + '-' + pad_(m);
}

function buildReport_(all, month) {
  var prev = prevMonth_(month);
  var names = Object.keys(all.stores).filter(function (n) { return all.stores[n][month]; });

  // 全店が休んだ日＝定休日。他店が開いていた日に休んだ日だけ臨時休業として拾う。
  var businessDays = {};
  names.forEach(function (n) {
    all.stores[n][month].forEach(function (r) { businessDays[r.day] = r.dow; });
  });

  var stores = names.map(function (n) {
    var rows = all.stores[n][month];
    var s = summarize_(rows);
    s.name = n;
    s.rows = rows;
    s.weekday = byWeekday_(rows);

    var openDays = {};
    rows.forEach(function (r) { openDays[r.day] = true; });
    s.closed = Object.keys(businessDays).map(Number).filter(function (d) { return !openDays[d]; })
      .sort(function (a, b) { return a - b; });
    s.loss = s.closed.reduce(function (a, d) {
      return a + (s.weekday[businessDays[d]] || 0);
    }, 0);

    var pr = all.stores[n][prev];
    if (pr && pr.length) {
      var p = summarize_(pr);
      s.prev = p;
      s.mom = p.sales ? s.sales / p.sales : 0;
      s.momDaily = p.daily ? s.daily / p.daily : 0;
      s.momUnit = p.unit ? s.unit / p.unit : 0;
      s.volumeEffect = (s.groups - p.groups) * p.unit;   // 組数が増えたぶん
      s.priceEffect = (s.unit - p.unit) * s.groups;      // 単価が動いたぶん
    }
    return s;
  }).sort(function (a, b) { return b.sales - a.sales; });

  var allRows = [];
  stores.forEach(function (s) { allRows = allRows.concat(s.rows); });
  var total = summarize_(allRows);
  total.weekday = byWeekday_(allRows);

  var rep = { month: month, stores: stores, total: total, closedDays: businessDays };

  var prevRows = [];
  names.forEach(function (n) {
    if (all.stores[n][prev]) prevRows = prevRows.concat(all.stores[n][prev]);
  });
  if (prevRows.length) {
    var pt = summarize_(prevRows);
    rep.prevTotal = pt;
    rep.mom = pt.sales ? total.sales / pt.sales : 0;
    rep.volumeEffect = (total.groups - pt.groups) * pt.unit;
    rep.priceEffect = (total.unit - pt.unit) * total.groups;
  }
  return rep;
}

// ---------------------------------------------------------------- 表示

function pad_(n) { return (n < 10 ? '0' : '') + n; }
function esc_(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}
function fmt_(n) {
  var v = Math.round(Math.abs(n)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return (n < 0 ? '−¥' : '¥') + v;
}
function pct_(v) { return Math.round(v * 100) + '%'; }

function chip_(ratio) {
  if (!ratio) return '<span class="chip flat">前月データなし</span>';
  var cls = ratio > 1.005 ? 'up' : (ratio < 0.995 ? 'down' : 'flat');
  var mark = cls === 'up' ? '▲' : (cls === 'down' ? '▼' : '');
  return '<span class="chip ' + cls + '">' + mark + pct_(ratio) + '</span>';
}

function bars_(rows) {
  var max = 0;
  rows.forEach(function (r) { if (r.sales > max) max = r.sales; });
  if (!max) return '';
  return '<div class="cols">' + rows.map(function (r) {
    var we = (r.dow === '土' || r.dow === '日') ? ' we' : '';
    return '<div class="col' + we + '" title="' + r.day + '日（' + r.dow + '） ' + fmt_(r.sales) + '">'
      + '<i style="height:' + (r.sales / max * 100).toFixed(1) + '%"></i>'
      + '<span>' + r.day + '</span></div>';
  }).join('') + '</div>';
}

function renderReport_(rep, months, current) {
  var y = rep.month.slice(0, 4), m = Number(rep.month.slice(5));
  var t = rep.total;
  var h = [];

  h.push('<div class="wrap">');
  h.push('<header><div class="eyebrow">Airレジ 売上レポート</div>'
    + '<h1>' + y + '年' + m + '月</h1>'
    + '<div class="months">' + months.slice(0, 12).map(function (ym) {
      var mm = Number(ym.slice(5));
      var on = (ym === current) ? ' on' : '';
      return '<a class="m' + on + '" href="?month=' + ym + '">' + ym.slice(2, 4) + '/' + pad_(mm) + '</a>';
    }).join('') + '</div></header>');

  // 全体
  h.push('<div class="kpis">');
  h.push('<div class="kpi big"><span class="k">売上合計（' + rep.stores.length + '店）</span>'
    + '<span class="v">' + fmt_(t.sales) + '</span><span class="s">'
    + (rep.prevTotal ? '前月 ' + fmt_(rep.prevTotal.sales) + '　' + chip_(rep.mom) : '&nbsp;')
    + '</span></div>');
  h.push('<div class="kpi"><span class="k">組数（会計件数）</span><span class="v">'
    + t.groups.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '</span>'
    + '<span class="s">1日あたり ' + (t.days ? (t.groups / t.days).toFixed(1) : 0) + '組</span></div>');
  h.push('<div class="kpi"><span class="k">組単価</span><span class="v">' + fmt_(t.unit) + '</span>'
    + '<span class="s">' + (rep.prevTotal ? '前月 ' + fmt_(rep.prevTotal.unit) + '　'
      + chip_(rep.prevTotal.unit ? t.unit / rep.prevTotal.unit : 0) : '&nbsp;') + '</span></div>');
  h.push('<div class="kpi big"><span class="k">キャッシュレス比率</span><span class="v">'
    + (t.cashless * 100).toFixed(1) + '%</span>'
    + '<span class="s">現金 ' + fmt_(t.cash) + '</span></div>');
  h.push('</div>');

  // 増減の内訳
  if (rep.prevTotal) {
    var diff = t.sales - rep.prevTotal.sales;
    h.push('<section><h2>前月からの増減</h2><div class="card">');
    h.push('<div class="wf"><span>差引</span><b class="' + (diff >= 0 ? 'good' : 'bad') + '">'
      + (diff >= 0 ? '+' : '') + fmt_(diff) + '</b></div>');
    h.push('<div class="wf sub"><span>組数が増えたぶん</span><b class="'
      + (rep.volumeEffect >= 0 ? 'good' : 'bad') + '">'
      + (rep.volumeEffect >= 0 ? '+' : '') + fmt_(rep.volumeEffect) + '</b></div>');
    h.push('<div class="wf sub"><span>単価が動いたぶん</span><b class="'
      + (rep.priceEffect >= 0 ? 'good' : 'bad') + '">'
      + (rep.priceEffect >= 0 ? '+' : '') + fmt_(rep.priceEffect) + '</b></div>');
    h.push('</div><p class="note">売上の増減を「お客さんの数が変わったぶん」と「1組あたりの単価が変わったぶん」に分けたもの。</p></section>');
  }

  // 店舗別
  h.push('<section><h2>店舗別</h2>');
  rep.stores.forEach(function (s) {
    h.push('<div class="card store">');
    h.push('<div class="srow"><h3>' + esc_(s.name) + '</h3><div class="stotal">' + fmt_(s.sales)
      + ' ' + (s.mom ? chip_(s.mom) : '') + '</div></div>');
    h.push('<div class="grid">'
      + cell_('日商', fmt_(s.daily))
      + cell_('営業日', s.days + '日')
      + cell_('組数', String(s.groups))
      + cell_('組単価', fmt_(s.unit))
      + cell_('1組の品数', s.perGroup.toFixed(1) + '点')
      + cell_('キャッシュレス', (s.cashless * 100).toFixed(0) + '%')
      + '</div>');
    h.push(bars_(s.rows));
    if (s.closed.length) {
      h.push('<p class="warn">臨時休業 ' + s.closed.length + '日（'
        + s.closed.map(function (d) { return d + '日(' + rep.closedDays[d] + ')'; }).join('、')
        + '）… 同じ曜日の平均で計算すると約' + fmt_(s.loss) + '</p>');
    }
    if (s.guestInput > 0 && s.guestInput < 0.99) {
      h.push('<p class="note small">人数入力 ' + Math.round(s.guestInput * 100)
        + '%（会計時の人数入力が漏れているぶん、1人あたり単価は出せません）</p>');
    }
    h.push('</div>');
  });
  h.push('</section>');

  // 曜日別
  var wkOrder = ['月', '火', '水', '木', '金', '土', '日'];
  var wkMax = 0;
  wkOrder.forEach(function (w) { if (t.weekday[w] > wkMax) wkMax = t.weekday[w]; });
  h.push('<section><h2>曜日別（全店・1営業日あたり）</h2><div class="card"><div class="dow">');
  wkOrder.forEach(function (w) {
    var v = t.weekday[w] || 0;
    h.push('<div' + (v ? '' : ' class="off"') + '><span class="d">' + w + '</span>'
      + '<span class="track"><i style="width:' + (wkMax ? v / wkMax * 100 : 0) + '%"></i></span>'
      + '<span class="mv">' + (v ? fmt_(v) : '休') + '</span></div>');
  });
  h.push('</div></div></section>');

  var totalLoss = rep.stores.reduce(function (a, s) { return a + s.loss; }, 0);
  if (totalLoss > 0) {
    h.push('<section><h2>臨時休業のぶん</h2><div class="card">'
      + '<p class="big-loss">' + fmt_(totalLoss) + '</p>'
      + '<p class="note">他のお店が営業していた日に休んだぶんを、その店の同じ曜日の平均で計算した目安です。'
      + '全店そろって休んだ日（定休日）は入っていません。</p></div></section>');
  }

  h.push('<footer><p>Googleドライブの「' + esc_(FOLDER_NAME) + '」フォルダのCSVから自動作成。'
    + '金額は税込・値引後。</p>'
    + '<p>「組数」は会計の件数です。Airレジの「客数」は会計時の手入力なので、'
    + '入力もれがあると少なく出ます。</p></footer>');
  h.push('</div>');
  return h.join('');
}

function cell_(label, value) {
  return '<div class="c"><span>' + label + '</span><b>' + value + '</b></div>';
}

// ---------------------------------------------------------------- 毎月のメール

/** 毎月1日の朝にレポートのお知らせメールを送る（トリガーから呼ばれる） */
function sendMonthlyMail() {
  var all = loadAll_();
  if (!all.months.length) return;
  var month = all.months[0];
  var rep = buildReport_(all, month);
  var y = month.slice(0, 4), m = Number(month.slice(5));
  var url = '';
  try { url = ScriptApp.getService().getUrl(); } catch (e) { url = ''; }

  var lines = ['<p>' + y + '年' + m + '月の売上がまとまりました。</p>'];
  lines.push('<p style="font-size:20px;margin:8px 0"><b>' + fmt_(rep.total.sales) + '</b>'
    + (rep.mom ? '（前月比 ' + pct_(rep.mom) + '）' : '') + '</p>');
  lines.push('<table cellpadding="6" style="border-collapse:collapse;font-size:14px">');
  rep.stores.forEach(function (s) {
    lines.push('<tr><td style="border-bottom:1px solid #ddd">' + esc_(s.name) + '</td>'
      + '<td style="border-bottom:1px solid #ddd;text-align:right">' + fmt_(s.sales) + '</td>'
      + '<td style="border-bottom:1px solid #ddd;text-align:right">'
      + (s.mom ? pct_(s.mom) : '-') + '</td></tr>');
  });
  lines.push('</table>');
  if (url) lines.push('<p><a href="' + url + '">くわしいレポートを見る</a></p>');

  MailApp.sendEmail({
    to: MAIL_TO || Session.getEffectiveUser().getEmail(),
    subject: '【売上レポート】' + y + '年' + m + '月　' + fmt_(rep.total.sales),
    htmlBody: lines.join('')
  });
}

/** 毎月1日の朝9時にメールが届くようにする（1回だけ実行すればOK） */
function setupMonthlyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendMonthlyMail') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendMonthlyMail').timeBased().onMonthDay(1).atHour(9).create();
  Logger.log('毎月1日の朝9時にメールが届くようにしました。');
}

/** 動作確認用。実行するとログに集計結果が出る */
function testRun() {
  var all = loadAll_();
  Logger.log('見つかった月: ' + all.months.join(', '));
  Logger.log('見つかった店舗: ' + Object.keys(all.stores).join(', '));
  if (!all.months.length) return;
  var rep = buildReport_(all, all.months[0]);
  Logger.log(all.months[0] + ' 合計 ' + fmt_(rep.total.sales)
    + ' / ' + rep.total.groups + '組 / 組単価 ' + fmt_(rep.total.unit));
  rep.stores.forEach(function (s) {
    Logger.log('  ' + s.name + ' ' + fmt_(s.sales) + ' (' + s.days + '日, '
      + s.groups + '組, 組単価 ' + fmt_(s.unit) + ', CL ' + (s.cashless * 100).toFixed(0) + '%)');
  });
}

// ---------------------------------------------------------------- 見た目

var PAGE_HEAD = '<!DOCTYPE html><html lang="ja"><head><meta charset="utf-8">'
  + '<style>'
  + ':root{--ground:#eceeed;--surface:#fbfcfb;--ink:#131a19;--ink-2:#3d4948;--muted:#65726f;'
  + '--line:rgba(19,26,25,.14);--line-soft:rgba(19,26,25,.08);--accent:#2b5f6b;--accent-soft:#8fb9c0;'
  + '--good:#2e7355;--bad:#a8483c;--flat:#7b8683;'
  + '--serif:"Hiragino Mincho ProN","Yu Mincho","Noto Serif JP",serif;'
  + '--sans:-apple-system,"Hiragino Sans","Noto Sans JP","Yu Gothic",sans-serif}'
  + '@media (prefers-color-scheme:dark){:root{--ground:#0d1214;--surface:#151d1f;--ink:#e6ecea;'
  + '--ink-2:#c2cdcb;--muted:#8d9b98;--line:rgba(230,236,234,.16);--line-soft:rgba(230,236,234,.09);'
  + '--accent:#78b8c4;--accent-soft:#3d6d77;--good:#5fb68d;--bad:#dc8577;--flat:#8d9b98}}'
  + '*{box-sizing:border-box}'
  + 'body{margin:0;background:var(--ground);color:var(--ink);font-family:var(--sans);'
  + 'font-size:15px;line-height:1.7;-webkit-text-size-adjust:100%}'
  + '.wrap{max-width:760px;margin:0 auto;padding:1.25rem 1rem 3rem;display:flex;'
  + 'flex-direction:column;gap:1.75rem}'
  + '.eyebrow{font-size:11px;letter-spacing:.16em;color:var(--muted);font-weight:600}'
  + 'h1{font-family:var(--serif);font-size:2rem;margin:.2rem 0 .6rem;font-weight:600}'
  + 'h2{font-family:var(--serif);font-size:1.15rem;margin:0 0 .6rem;font-weight:600;'
  + 'border-bottom:1px solid var(--line);padding-bottom:.4rem}'
  + 'h3{font-size:15px;margin:0;font-weight:600}'
  + '.months{display:flex;flex-wrap:wrap;gap:.35rem}'
  + '.months .m{font-size:12px;padding:.2rem .55rem;border:1px solid var(--line);border-radius:999px;'
  + 'color:var(--muted);text-decoration:none;font-variant-numeric:tabular-nums}'
  + '.months .m.on{background:var(--accent);border-color:var(--accent);color:#fff}'
  + '.kpis{display:grid;grid-template-columns:1fr 1fr;gap:.6rem}'
  + '.kpi{background:var(--surface);border:1px solid var(--line-soft);border-radius:6px;'
  + 'padding:.8rem .9rem;display:flex;flex-direction:column;gap:.1rem}'
  + '.kpi.big{grid-column:1/-1}'
  + '.kpi .k{font-size:11.5px;color:var(--muted);font-weight:600}'
  + '.kpi .v{font-family:var(--serif);font-size:1.6rem;line-height:1.25;font-variant-numeric:tabular-nums}'
  + '.kpi.big .v{font-size:2.1rem}.kpi.big+.kpi.big .v{font-size:1.6rem}'
  + '.kpi .s{font-size:12px;color:var(--ink-2);font-variant-numeric:tabular-nums}'
  + '.card{background:var(--surface);border:1px solid var(--line-soft);border-radius:6px;'
  + 'padding:.9rem 1rem;display:flex;flex-direction:column;gap:.6rem;margin-bottom:.6rem}'
  + '.chip{display:inline-block;font-size:12px;font-weight:600;padding:.05rem .4rem;border-radius:3px;'
  + 'font-variant-numeric:tabular-nums}'
  + '.up{color:var(--good);background:rgba(46,115,85,.13)}'
  + '.down{color:var(--bad);background:rgba(168,72,60,.13)}'
  + '.flat{color:var(--flat);background:rgba(123,134,131,.13)}'
  + '.srow{display:flex;justify-content:space-between;align-items:baseline;gap:.5rem;flex-wrap:wrap}'
  + '.stotal{font-family:var(--serif);font-size:1.25rem;font-variant-numeric:tabular-nums}'
  + '.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:.5rem .4rem}'
  + '.c{display:flex;flex-direction:column;line-height:1.4}'
  + '.c span{font-size:10.5px;color:var(--muted)}'
  + '.c b{font-size:14px;font-weight:600;font-variant-numeric:tabular-nums}'
  + '.cols{display:flex;align-items:flex-end;gap:2px;height:76px}'
  + '.col{flex:1;display:flex;flex-direction:column;justify-content:flex-end;gap:3px;height:100%}'
  + '.col i{display:block;background:var(--accent-soft);border-radius:2px 2px 0 0}'
  + '.col.we i{background:var(--accent)}'
  + '.col span{font-size:8px;color:var(--muted);text-align:center}'
  + '.wf{display:flex;justify-content:space-between;align-items:baseline;font-size:14px}'
  + '.wf b{font-family:var(--serif);font-size:1.25rem;font-variant-numeric:tabular-nums}'
  + '.wf.sub{font-size:13px;color:var(--ink-2);padding-left:1rem}'
  + '.wf.sub b{font-size:1rem;font-family:var(--sans)}'
  + '.good{color:var(--good)} .bad{color:var(--bad)}'
  + '.dow{display:flex;flex-direction:column;gap:.35rem}'
  + '.dow>div{display:grid;grid-template-columns:1.5rem 1fr 5.5rem;align-items:center;gap:.5rem}'
  + '.dow .d{font-family:var(--serif)}'
  + '.dow .track{height:10px;background:var(--line-soft);border-radius:2px;overflow:hidden}'
  + '.dow .track i{display:block;height:100%;background:var(--accent);border-radius:2px}'
  + '.dow .mv{text-align:right;font-size:13px;font-variant-numeric:tabular-nums}'
  + '.dow .off{opacity:.45}'
  + '.big-loss{font-family:var(--serif);font-size:1.7rem;margin:0;color:var(--bad);'
  + 'font-variant-numeric:tabular-nums}'
  + '.note{font-size:12.5px;color:var(--muted);margin:0}'
  + '.note.small{font-size:11.5px}'
  + '.warn{font-size:12.5px;color:var(--bad);margin:0}'
  + '.empty{padding:3rem 1rem;text-align:center;color:var(--muted)}'
  + 'footer{border-top:1px solid var(--line);padding-top:.8rem;font-size:12px;color:var(--muted)}'
  + 'footer p{margin:.2rem 0}'
  + '</style></head><body>';
