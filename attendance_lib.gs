/**
 * 出勤簿タブの共通ユーティリティ
 *
 * holiday_color.gs と sheet_cleanup.gs から使う。
 * どちらかを使うときは、このファイルも同じ Apps Script プロジェクトに入れること。
 *
 * 出勤簿タブの構成
 *   1行目  出勤簿　氏名（パート・アルバイト）　期間
 *   3行目  B=時給／日給ラベル  K=単価  M=通勤手当日額
 *   4行目  集計期間：YYYY年M月D日〜YYYY年M月D日
 *   5行目  日付 / 曜 / 出勤 / 退勤 / 休憩(分) / 実働(h) / 深夜(h) / 備考   = B〜I列
 *   6行目〜 日次データ（31日分の枠。30日以下の月は末尾が余る）
 *   合計行  D=出勤日数  G=実働合計  H=深夜合計
 *   明細    基本賃金 … 差引支給額
 */

const WEEKDAYS = ['日', '月', '火', '水', '木', '金', '土'];

/** 出勤簿タブの行位置。出勤簿でなければ null */
function readAttendanceRows(sh) {
  const last = sh.getLastRow();
  if (last < 10) return null;
  const labels = sh.getRange(1, 2, last, 1).getValues().map(function (r) {
    return String(r[0]).trim();
  });
  if (labels[0].indexOf('出勤簿') < 0) return null;

  const rateLabel = labels[2];
  if (rateLabel.indexOf('時給') < 0 && rateLabel.indexOf('日給') < 0) return null;

  const headerRow = labels.indexOf('日付') + 1;
  const totalRow = labels.indexOf('合計') + 1;
  if (!headerRow || !totalRow || totalRow <= headerRow) return null;

  let periodRow = 0, detailStart = 0, basicRow = 0, netRow = 0, netLabel = '';
  for (let i = 0; i < headerRow; i++) {
    if (labels[i].indexOf('集計期間') === 0) periodRow = i + 1;
  }
  for (let i = totalRow; i < labels.length; i++) {
    if (!detailStart && labels[i] === '給与明細') detailStart = i + 1;
    if (!basicRow && labels[i].indexOf('基本賃金') === 0) basicRow = i + 1;
    if (!netRow && labels[i].indexOf('差引支給額') === 0) { netRow = i + 1; netLabel = labels[i]; }
  }
  if (!basicRow || !netRow) return null;

  return {
    isDailyWage: rateLabel.indexOf('日給') >= 0,
    periodRow: periodRow,
    detailStart: detailStart,
    headerRow: headerRow,
    dataStart: headerRow + 1,
    dataEnd: totalRow - 1,
    totalRow: totalRow,
    basicRow: basicRow,
    netRow: netRow,
    netLabel: netLabel,
  };
}

/**
 * 集計期間の見出しから期間を読む。{ start, end, days } を返す。
 *
 * 日付を知りたいときに B 列の値を使ってはいけない。
 * 事前作成された 22期のスプレッドシートでは、2027年になるはずの日付が
 * 2026年で入っている個体がある（表示書式が m/d なので画面では気づけない）。
 * 曜（C列）はこの集計期間と一致しているので、集計期間を正とする。
 */
function readPeriod(sh, at) {
  if (!at || !at.periodRow) return null;
  const text = String(sh.getRange(at.periodRow, 2).getValue());
  const m = text.match(/(\d{4})年(\d{1,2})月(\d{1,2})日\s*[〜～~-]\s*(\d{4})年(\d{1,2})月(\d{1,2})日/);
  if (!m) return null;

  const start = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const end = new Date(Number(m[4]), Number(m[5]) - 1, Number(m[6]));
  const days = Math.round((end - start) / 86400000) + 1;
  if (days < 27 || days > 32) return null;          // 1か月分でなければ扱わない
  if (days > at.dataEnd - at.dataStart + 1) return null; // 枠に収まらなければ扱わない

  return { start: start, end: end, days: days };
}

/**
 * そのスプレッドシート自体の対象期間を、給与一覧タブの見出しから読む。
 * 「給与一覧表（パート・アルバイト）　2026/12/16〜2027/1/15」の形。
 */
function readBookPeriod(ss) {
  const sh = getSummarySheet(ss);
  if (!sh) return null;

  const text = String(sh.getRange(1, 1).getValue()) + ' ' + String(sh.getRange(1, 2).getValue());
  const m = text.match(/(\d{4})\/(\d{1,2})\/(\d{1,2})\s*[〜～~-]\s*(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (!m) return null;

  const start = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const end = new Date(Number(m[4]), Number(m[5]) - 1, Number(m[6]));
  return { start: start, end: end, days: Math.round((end - start) / 86400000) + 1 };
}

/**
 * 出勤簿タブの期間を返す。手を付けてはいけないタブなら null。
 *
 * 事前作成された22期のスプレッドシートには、前のテンプレートのまま
 * 作り直されていないタブが混ざっている。たとえば 2026/12/16〜2027/1/15 の
 * ブックに「集計期間：2026年3月16日〜2026年4月15日」のタブが16枚ある。
 * ブック自身の期間と食い違うタブは、日付も曜も別の月のものなので触らない。
 */
function tabPeriod(sh, at, bookPeriod, log) {
  const p = readPeriod(sh, at);
  if (!p) {
    log.push('  ? 集計期間を読めません。手を付けません: ' + sh.getName());
    return null;
  }
  if (bookPeriod && dateKey(p.start) !== dateKey(bookPeriod.start)) {
    log.push('  !! 期間がブックと違います。手を付けません: ' + sh.getName()
      + '（タブ ' + fmtDate(p.start) + '〜' + fmtDate(p.end) + '）');
    return null;
  }
  // 曜（C列）は全行を突き合わせる。1行目だけ正しくて途中から前のテンプレートの
  // まま、というタブが実在するため（例：2026/12/16〜2027/1/15 ブックの衣幡千明は
  // R6〜R7 だけ正しく、R8 以降が 3/18〜4/15 のまま）。
  const wd = sh.getRange(at.dataStart, 3, p.days, 1).getValues();
  let ng = 0;
  for (let i = 0; i < p.days; i++) {
    const c = String(wd[i][0]).trim();
    if (c && c !== WEEKDAYS[addDays(p.start, i).getDay()]) ng += 1;
  }
  if (ng) {
    log.push('  !! 曜が集計期間と合いません（' + ng + '行）。手を付けません: ' + sh.getName());
    return null;
  }
  return p;
}

function addDays(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}

/** Date を 'yyyy-MM-dd' に。日付でなければ null */
function dateKey(v) {
  if (!(v instanceof Date) || isNaN(v.getTime())) return null;
  return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function sameColor(a, b) {
  return String(a).toLowerCase() === String(b).toLowerCase();
}

function fmtDate(d) {
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/M/d');
}

/** 氏名の表記ゆれ（全角・半角スペース）を吸収する */
function normalizeName(s) {
  return String(s).replace(/[\s\u3000]+/g, '');
}

/**
 * 出勤簿タブを氏名で引けるようにする。
 * タブ名ではなく1行目の見出しから氏名を取る。給与一覧の「衣幡千明（海事・横河）」に
 * 対してタブ名は「衣幡千明_海事横河」というように、両者は一致しないため。
 */
function attendanceIndex(ss) {
  const map = {};
  ss.getSheets().forEach(function (sh) {
    const at = readAttendanceRows(sh);
    if (!at) return;
    const title = String(sh.getRange(1, 1).getValue()) + String(sh.getRange(1, 2).getValue());
    const m = title.match(/出勤簿[\s\u3000]*(.+?)（パート・アルバイト）/);
    if (!m) return;
    map[normalizeName(m[1])] = { sheet: sh, at: at };
  });
  return map;
}

/** 給与一覧タブのメンバー行（A列が番号、B列が氏名の行）を返す */
function readSummaryMembers(sh) {
  const last = sh.getLastRow();
  const a = sh.getRange(1, 1, last, 1).getValues();
  const b = sh.getRange(1, 2, last, 1).getValues();
  const out = [];
  for (let i = 0; i < last; i++) {
    const av = String(a[i][0]).trim();
    const bv = String(b[i][0]).trim();
    if (av.indexOf('店舗別合計') >= 0) break;   // ここから下は集計ブロック
    if (av === '' || bv === '') continue;
    if (av.indexOf('▶') === 0) continue;        // グループ見出し
    if (isNaN(Number(av))) continue;
    out.push({ row: i + 1, no: Number(av), name: bv });
  }
  return out;
}

/** 給与一覧タブ */
function getSummarySheet(ss) {
  return ss.getSheetByName('給与一覧') || ss.getSheets().filter(function (s) {
    return String(s.getRange(1, 1).getValue()).indexOf('給与一覧表') >= 0
        || String(s.getRange(1, 2).getValue()).indexOf('給与一覧表') >= 0;
  })[0] || null;
}

/** フォルダ内の給与一覧スプレッドシートを列挙する */
function listPayrollBooks(folderIds, log) {
  const out = [];
  folderIds.forEach(function (fid) {
    const it = DriveApp.getFolderById(fid).getFilesByType(MimeType.GOOGLE_SHEETS);
    while (it.hasNext()) {
      const f = it.next();
      if (f.getName().indexOf('給与一覧') !== 0) continue; // 給与一覧以外は触らない
      out.push({ id: f.getId(), name: f.getName() });
    }
  });
  out.sort(function (a, b) { return a.name < b.name ? -1 : 1; });
  if (log) log.push('対象スプレッドシート ' + out.length + '本');
  return out;
}
