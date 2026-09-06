/**
 * 出勤簿の祝日を赤字にするスクリプト
 *
 * 対象：22期_PA勤怠フォルダの給与一覧スプレッドシート（2026/9/16〜2027/9/15 の12本）
 * 必要ファイル：attendance_lib.gs
 *
 * 各スプレッドシートの出勤簿タブについて、祝日（振替休日・国民の休日を含む）に
 * あたる行の B〜H 列を、日曜と同じ赤字 #cc0000 にする。
 *
 * 既存の書式は getFontColors() で読んでから該当行だけ差し替えるため、
 * 日曜の赤・土曜の青・備考欄の色はそのまま残る。
 *
 * どの行が何月何日かは、B列の日付ではなく4行目の集計期間から決める。
 * 事前作成された22期のスプレッドシートは、2027年になるはずの日付が2026年で
 * 入っている個体があり、B列をそのまま使うと1年ずれた祝日を塗ってしまうため。
 * sheet_cleanup.gs の fixDateColumn() を先に流しておくのが望ましい。
 *
 * 使い方
 *   1. attendance_lib.gs と一緒に Apps Script プロジェクトに追加
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

  const books = HOL.SPREADSHEET_IDS.length
    ? HOL.SPREADSHEET_IDS.map(function (id) { return { id: id, name: id }; })
    : listPayrollBooks(HOL.FOLDER_IDS, log);

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
  const bookPeriod = readBookPeriod(ss);
  if (!bookPeriod) log.push('  ? 給与一覧タブから期間を読めません。タブ側の集計期間だけで判断します');

  ss.getSheets().forEach(function (sh) {
    const at = readAttendanceRows(sh);
    if (!at) { skipped += 1; return; }

    const p = tabPeriod(sh, at, bookPeriod, log);
    if (!p) return;

    const cols = HOL.LAST_COL - HOL.FIRST_COL + 1;
    const block = sh.getRange(at.dataStart, HOL.FIRST_COL, p.days, cols);
    const colors = block.getFontColors();

    let changed = 0;
    for (let i = 0; i < p.days; i++) {
      const key = dateKey(addDays(p.start, i));
      if (!holidays[key]) continue;
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
