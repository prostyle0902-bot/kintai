/**
 * 出勤簿の祝日を赤字にするスクリプト
 *
 * 対象：22期_PA勤怠フォルダの給与一覧スプレッドシート（2026/9/16〜2027/9/15 の12本）
 *
 * 各スプレッドシートの出勤簿タブについて、日付欄が祝日（振替休日・国民の休日を含む）
 * にあたる行の B〜H 列を、日曜と同じ赤字 #cc0000 にする。
 *
 * 既存の書式は getFontColors() で読んでから該当行だけ差し替えるため、
 * 日曜の赤・土曜の青・備考欄の色はそのまま残る。
 *
 * 使い方
 *   1. Apps Script プロジェクトにこのファイルを追加
 *   2. HOL.DRY_RUN = true のまま colorHolidays() を実行し、ログで対象日を確認
 *   3. 問題なければ HOL.DRY_RUN = false にして再実行
 *      （初回は日本の祝日カレンダーへのアクセス許可を求められる）
 */

const HOL = {
  // true の間は一切書き込まず、ログ出力だけ行う
  DRY_RUN: true,

  // 対象フォルダ。期が変わったら追加する
  FOLDER_IDS: [
    '1JQPlAe-jOMhCxg2jIbajcIyG0w7zXa6F', // 22期_PA勤怠
  ],

  // 個別に対象を指定したい場合はこちらに ID を入れる（FOLDER_IDS より優先）
  SPREADSHEET_IDS: [],

  // 日曜と同じ赤
  COLOR: '#cc0000',

  // 色を塗る列（日付〜深夜）。備考(I)より右は対象外
  FIRST_COL: 2, // B
  LAST_COL: 8,  // H

  // 日本の祝日カレンダー
  CALENDAR_ID: 'ja.japanese#holiday@group.v.calendar.google.com',
};

/**
 * 日本の祝日カレンダーが読めなかったときのフォールバック。
 * 2026/9/16〜2027/9/15 の祝日（Google カレンダー「日本の祝日」で確認済み）。
 * 「七五三」「クリスマス」などの祭日・行事は祝日ではないので含めない。
 */
const HOLIDAY_FALLBACK = [
  '2026-09-21', // 敬老の日
  '2026-09-22', // 国民の休日
  '2026-09-23', // 秋分の日
  '2026-10-12', // スポーツの日
  '2026-11-03', // 文化の日
  '2026-11-23', // 勤労感謝の日
  '2027-01-01', // 元日
  '2027-01-11', // 成人の日
  '2027-02-11', // 建国記念の日
  '2027-02-23', // 天皇誕生日
  '2027-03-21', // 春分の日
  '2027-03-22', // 春分の日 振替休日
  '2027-04-29', // 昭和の日
  '2027-05-03', // 憲法記念日
  '2027-05-04', // みどりの日
  '2027-05-05', // こどもの日
  '2027-07-19', // 海の日
  '2027-08-11', // 山の日
];


function colorHolidays() {
  const log = [];
  const holidays = loadHolidays(log);

  targetSpreadsheets(log).forEach(function (file) {
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
      colorOneSpreadsheet(ss, holidays, log);
    } catch (e) {
      log.push('!! 中断: ' + e.message);
    }
  });

  log.push('');
  log.push(HOL.DRY_RUN
    ? '*** DRY_RUN です。書き込みは行っていません。***'
    : '*** 書き込みを実行しました。***');
  Logger.log(log.join('\n'));
}


function colorOneSpreadsheet(ss, holidays, log) {
  let done = 0, skipped = 0;
  const hit = {};

  ss.getSheets().forEach(function (sh) {
    const at = readAttendanceRows(sh);
    if (!at) { skipped += 1; return; }

    const rows = at.dataEnd - at.dataStart + 1;
    const dates = sh.getRange(at.dataStart, HOL.FIRST_COL, rows, 1).getValues();
    const cols = HOL.LAST_COL - HOL.FIRST_COL + 1;
    const block = sh.getRange(at.dataStart, HOL.FIRST_COL, rows, cols);
    const colors = block.getFontColors();

    let changed = 0;
    for (let i = 0; i < rows; i++) {
      const key = dateKey(dates[i][0]);
      if (!key || !holidays[key]) continue;
      hit[key] = holidays[key];
      if (colors[i].every(function (c) { return sameColor(c, HOL.COLOR); })) continue;
      for (let k = 0; k < cols; k++) colors[i][k] = HOL.COLOR;
      changed += 1;
    }

    if (!changed) return;
    done += 1;
    if (!HOL.DRY_RUN) block.setFontColors(colors);
  });

  const days = Object.keys(hit).sort().map(function (k) { return k + '(' + hit[k] + ')'; });
  log.push('- 祝日 ' + days.length + '日: ' + (days.join('、') || 'なし'));
  log.push('- 出勤簿 ' + done + 'タブを赤字に変更（対象外タブ ' + skipped + '）');
}


/* ------------------------------------------------------------------ *
 * 祝日の取得
 * ------------------------------------------------------------------ */

/** { 'yyyy-MM-dd': '敬老の日', ... } を返す */
function loadHolidays(log) {
  const map = {};
  const from = new Date(2026, 0, 1);
  const to = new Date(2030, 0, 1);
  try {
    const cal = CalendarApp.getCalendarById(HOL.CALENDAR_ID);
    if (!cal) throw new Error('カレンダーを取得できません');
    cal.getEvents(from, to).forEach(function (ev) {
      // このカレンダーは「祝日」と「祭日（七五三・クリスマス等）」が混在する。
      // 説明の1行目が「祝日」のものだけを採用する。
      const head = String(ev.getDescription() || '').split('\n')[0].trim();
      if (head !== '祝日') return;
      map[dateKey(ev.getAllDayStartDate() || ev.getStartTime())] = ev.getTitle();
    });
    if (!Object.keys(map).length) throw new Error('祝日を取得できません');
    log.push('祝日カレンダーから ' + Object.keys(map).length + '日を取得');
    return map;
  } catch (e) {
    log.push('! 祝日カレンダーを読めないため、内蔵リストを使用: ' + e.message);
    HOLIDAY_FALLBACK.forEach(function (d) { map[d] = '祝日'; });
    return map;
  }
}

/** Date を 'yyyy-MM-dd' に。日付でなければ null */
function dateKey(v) {
  if (!(v instanceof Date) || isNaN(v.getTime())) return null;
  return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
}

function sameColor(a, b) {
  return String(a).toLowerCase() === String(b).toLowerCase();
}


/* ------------------------------------------------------------------ *
 * 対象スプレッドシートの列挙
 * ------------------------------------------------------------------ */
function targetSpreadsheets(log) {
  if (HOL.SPREADSHEET_IDS.length) {
    return HOL.SPREADSHEET_IDS.map(function (id) { return { id: id, name: id }; });
  }
  const out = [];
  HOL.FOLDER_IDS.forEach(function (fid) {
    const folder = DriveApp.getFolderById(fid);
    const it = folder.getFilesByType(MimeType.GOOGLE_SHEETS);
    while (it.hasNext()) {
      const f = it.next();
      if (f.getName().indexOf('給与一覧') !== 0) continue; // 給与一覧以外は触らない
      out.push({ id: f.getId(), name: f.getName() });
    }
  });
  out.sort(function (a, b) { return a.name < b.name ? -1 : 1; });
  log.push('対象スプレッドシート ' + out.length + '本');
  return out;
}


/* ------------------------------------------------------------------ *
 * 出勤簿タブの構造。payroll_fix.gs と同じ判定だが、同じプロジェクトに
 * 両方を入れても衝突しないよう独立した名前にしてある
 * ------------------------------------------------------------------ */
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

  let basicRow = 0, netRow = 0, netLabel = '';
  for (let i = totalRow; i < labels.length; i++) {
    if (!basicRow && labels[i].indexOf('基本賃金') === 0) basicRow = i + 1;
    if (!netRow && labels[i].indexOf('差引支給額') === 0) { netRow = i + 1; netLabel = labels[i]; }
  }
  if (!basicRow || !netRow) return null;

  return {
    isDailyWage: rateLabel.indexOf('日給') >= 0,
    headerRow: headerRow,
    dataStart: headerRow + 1,
    dataEnd: totalRow - 1,
    totalRow: totalRow,
    basicRow: basicRow,
    netRow: netRow,
    netLabel: netLabel,
  };
}
