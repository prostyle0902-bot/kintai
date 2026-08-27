/**
 * 給与一覧 一括補正スクリプト
 *
 * 対象：2026/8/16〜9/15 以降の給与一覧スプレッドシート（計13本）
 *
 * 補正内容
 *   1. 日給制スタッフ（三浦達弥）の基本賃金を「日給×実働時間」→「日給×出勤日数」に修正
 *   2. 給与一覧タブの「海事・横河」を 桜庭京子（退職）→ 衣幡千明（海事・横河）に差し替え
 *   3. 給与一覧タブの衣幡千明（トータルステイ）の参照ズレを修正
 *   4. 「神宮・亀甲堂・ゆう屋」を「ゆう屋」「神宮・亀甲堂」に分割し、鎌形房枝・宮内眞知子を追加
 *   5. 店舗別合計・総合計を再計算
 *
 * 使い方
 *   1. Apps Script プロジェクトにこのファイルを追加
 *   2. CONFIG.DRY_RUN = true のまま run() を実行し、実行ログで変更内容を確認
 *   3. 問題なければ CONFIG.DRY_RUN = false にして再実行
 *
 * 注意
 *   - 新規2名（CONFIG.NEW_STAFF）は rate が null の間はスキップされます。
 *     時給／日給と通勤手当を入れてから実行してください。
 */

const CONFIG = {
  // true の間は一切書き込まず、ログ出力だけ行う
  DRY_RUN: true,

  // 補正対象。2026/8/16〜9/15 以降の全スプレッドシート
  TARGET_IDS: [
    '1kknpC_KPyaHrldOpXTxC3rQke_sbRXRgYGq2IXlayAw', // 2026_8_16-2026_9_15
    '1Dxr27E2tGwiPKESxWWB8HCGDaX0i2jwtSqfL5Lyc9tw', // 2026_9_16-2026_10_15
    '1SSRb1I2RD6cOkRmBqW0tcZBVIntf_4RHDzDpzqXHi4w', // 2026_10_16-2026_11_15
    '1ARP6ApVaj7VzSKBuz0P5TcdnbH8BQJ5J5IzB_AQgmBw', // 2026_11_16-2026_12_15
    '1MwAXLQfBF-tBQ0jMHyJqLaNlUSbImikp23IdvD9CDBs', // 2026_12_16-2027_1_15
    '1awRXj9U0C18D_gT0tM5u14AENE-zak4d6pafseX5gAM', // 2027_1_16-2027_2_15
    '1rju65K08_cefudwVkKcfw4gJh4uPAWc8kDgpqmeFYC4', // 2027_2_16-2027_3_15
    '18bugMUJZtAQucKRLLA8CAawEiw7IrndNGHhDBIfOpbs', // 2027_3_16-2027_4_15
    '1QUPZqD496tJqtLUh04su1FmmzD4Q8GmcN0yTS9YhsD0', // 2027_4_16-2027_5_15
    '1S5Oq5bHtxBM2JaQApzfkOk2BPu9_GN04MjVCGM5XDn8', // 2027_5_16-2027_6_15
    '1K2q7smAXcqW0VF12FbSmo4vWSjOEwewYgAKfzX6dzWU', // 2027_6_16-2027_7_15
    '1MrtFtRVKxG97mIHiRkX27fkVXxrQNYy9Y5-Q__oZE8Y', // 2027_7_16-2027_8_15
    '1TrgIpopQedhDhzCyx4I0-uqg38ravAP0lw1HnBd3Y6o', // 2027_8_16-2027_9_15
  ],

  // 「海事・横河」の枠に入れる人（出勤簿タブの見出しで探す）
  KAIJI: { title: '衣幡千明（海事・横河）', name: '衣幡千明' },

  // 退職につき給与一覧・出勤簿から削除する人
  REMOVE_STAFF: ['桜庭京子'],

  // 「神宮・亀甲堂・ゆう屋」を分割して追加する2名
  // rate / commute を埋めるまでは追加処理をスキップします
  SPLIT_GROUP: '神宮・亀甲堂・ゆう屋',
  NEW_STAFF: [
    { name: '鎌形房枝',   store: 'ゆう屋',       rateType: '時給', rate: null, commute: null },
    { name: '宮内眞知子', store: '神宮・亀甲堂', rateType: '時給', rate: null, commute: null },
  ],

  // 新規出勤簿タブの雛形にする既存タブ（時給制・1店舗のみのシンプルな人）
  TEMPLATE_STAFF: '小倉孔子',
};

// 給与一覧タブの列
const COL = { NO: 1, NAME: 2, DAYS: 3, HOURS: 4, NIGHT: 5, PAY: 6, NOTE: 7 };

// 出勤簿タブの列
const AT = { LABEL: 2, DATE: 2, WORKED: 7, NIGHT: 8, AMOUNT: 7, RATE: 11, COMMUTE: 13 };


function run() {
  const log = [];
  CONFIG.TARGET_IDS.forEach(function (id) {
    let ss;
    try {
      ss = SpreadsheetApp.openById(id);
    } catch (e) {
      log.push('!! 開けません ' + id + ' : ' + e.message);
      return;
    }
    log.push('');
    log.push('========== ' + ss.getName() + ' ==========');
    try {
      fixSpreadsheet(ss, log);
    } catch (e) {
      log.push('!! 中断: ' + e.message);
    }
  });
  log.push('');
  log.push(CONFIG.DRY_RUN
    ? '*** DRY_RUN です。書き込みは行っていません。***'
    : '*** 書き込みを実行しました。***');
  Logger.log(log.join('\n'));
}


function fixSpreadsheet(ss, log) {
  fixDailyWageStaff(ss, log);

  const summary = getSummarySheet(ss);
  if (!summary) {
    log.push('!! 給与一覧タブが見つかりません');
    return;
  }
  let layout = readSummaryLayout(summary);

  fixKaijiRow(ss, summary, layout, log);
  fixIhataRow(ss, summary, layout, log);
  repointAllPayCells(ss, summary, layout, log);

  if (newStaffReady()) {
    splitGroupAndAddStaff(ss, summary, log);
    layout = readSummaryLayout(summary); // 行が増えたので読み直し
    renumber(summary, layout, log);
  } else {
    log.push('- 新規2名: 時給・通勤手当が未設定のためスキップ');
  }

  rebuildTotals(summary, readSummaryLayout(summary), log);
  removeRetiredSheets(ss, log);
}


/** 退職者の出勤簿タブが残っていれば削除する */
function removeRetiredSheets(ss, log) {
  CONFIG.REMOVE_STAFF.forEach(function (name) {
    const sh = findAttendanceSheet(ss, '出勤簿　' + name);
    if (!sh) return;
    log.push('- 出勤簿タブを削除: ' + sh.getName() + '（' + name + ' 退職）');
    if (!CONFIG.DRY_RUN) ss.deleteSheet(sh);
  });
}


/* ------------------------------------------------------------------ *
 * 1. 日給制スタッフの基本賃金を「日給×出勤日数」に直す
 * ------------------------------------------------------------------ */
function fixDailyWageStaff(ss, log) {
  ss.getSheets().forEach(function (sh) {
    const at = readAttendanceLayout(sh);
    if (!at || !at.isDailyWage) return;

    const formula = '=K3*COUNTIF(G' + at.dataStart + ':G' + at.dataEnd + ',">0")';
    log.push('- ' + sh.getName() + ' 基本賃金 R' + at.basicRow + ': ' + formula);
    if (CONFIG.DRY_RUN) return;

    sh.getRange(at.basicRow, AT.LABEL).setValue('基本賃金（日給×出勤日数）');
    sh.getRange(at.basicRow, AT.AMOUNT).setFormula(formula);
    sh.getRange(at.basicRow, 9).setValue('K3(日給)×出勤日数');
  });
}


/* ------------------------------------------------------------------ *
 * 2. 「海事・横河」を桜庭京子 → 衣幡千明（海事・横河）に差し替え
 * ------------------------------------------------------------------ */
function fixKaijiRow(ss, summary, layout, log) {
  const src = findAttendanceSheet(ss, CONFIG.KAIJI.title);
  if (!src) {
    log.push('!! 出勤簿「' + CONFIG.KAIJI.title + '」が見つかりません');
    return;
  }
  const at = readAttendanceLayout(src);
  if (!at) {
    log.push('!! ' + src.getName() + ' の書式を認識できません');
    return;
  }

  const member = findMember(layout, CONFIG.REMOVE_STAFF[0]);
  if (!member) {
    log.push('- 海事・横河: ' + CONFIG.REMOVE_STAFF[0] + ' の行なし（対応済みとみなします）');
    return;
  }

  const q = "'" + src.getName().replace(/'/g, "''") + "'!";
  log.push('- 給与一覧 R' + member.row + ': ' + CONFIG.REMOVE_STAFF[0] + ' → ' + CONFIG.KAIJI.name
    + ' (' + src.getName() + ')');
  if (CONFIG.DRY_RUN) return;

  summary.getRange(member.row, COL.NAME).setValue(CONFIG.KAIJI.name);
  summary.getRange(member.row, COL.DAYS).setFormula('=' + q + 'D' + at.totalRow);
  summary.getRange(member.row, COL.HOURS).setFormula('=' + q + 'G' + at.totalRow);
  summary.getRange(member.row, COL.NIGHT).setFormula('=' + q + 'H' + at.totalRow);
  summary.getRange(member.row, COL.PAY).setFormula('=' + q + 'G' + at.netRow);
}


/* ------------------------------------------------------------------ *
 * 3. 衣幡千明（トータルステイ）の参照ズレを直す
 *    彼女のタブだけ土曜／日曜加算・海事横河分・合算の行が増えており、
 *    「差引支給額（トータルステイ）」の行位置が他の人とずれている
 * ------------------------------------------------------------------ */
function fixIhataRow(ss, summary, layout, log) {
  const sh = findAttendanceSheet(ss, '出勤簿　' + CONFIG.KAIJI.name + '（パート');
  if (!sh) {
    log.push('!! 出勤簿「' + CONFIG.KAIJI.name + '」（トータルステイ）が見つかりません');
    return;
  }
  const at = readAttendanceLayout(sh);
  if (!at) return;

  // 海事・横河の行は fixKaijiRow で書き換え済みなので、同名でも別タブを参照している
  const member = findMemberInGroupOtherThan(layout, CONFIG.KAIJI.name, '海事・横河');
  if (!member) {
    log.push('!! 給与一覧に ' + CONFIG.KAIJI.name + '（トータルステイ）の行がありません');
    return;
  }

  const q = "'" + sh.getName().replace(/'/g, "''") + "'!";
  log.push('- 給与一覧 R' + member.row + ': ' + CONFIG.KAIJI.name
    + ' 支給額を ' + sh.getName() + '!G' + at.netRow + ' (' + at.netLabel + ') に再接続');
  if (CONFIG.DRY_RUN) return;

  summary.getRange(member.row, COL.DAYS).setFormula('=' + q + 'D' + at.totalRow);
  summary.getRange(member.row, COL.HOURS).setFormula('=' + q + 'G' + at.totalRow);
  summary.getRange(member.row, COL.NIGHT).setFormula('=' + q + 'H' + at.totalRow);
  summary.getRange(member.row, COL.PAY).setFormula('=' + q + 'G' + at.netRow);
}


/* ------------------------------------------------------------------ *
 * 給与一覧の全行を、対応する出勤簿タブの「差引支給額」行に貼り直す
 * （明細行数が人によって違うため、行位置のズレをまとめて解消する）
 * ------------------------------------------------------------------ */
function repointAllPayCells(ss, summary, layout, log) {
  layout.groups.forEach(function (g) {
    g.members.forEach(function (m) {
      if (m.name === CONFIG.KAIJI.name) return;              // 上で個別対応済み
      if (CONFIG.REMOVE_STAFF.indexOf(m.name) >= 0) return;  // 差し替え済みの退職者
      const sh = findAttendanceSheet(ss, '出勤簿　' + m.name + '（パート');
      if (!sh) {
        log.push('  ? 出勤簿なし: ' + m.name);
        return;
      }
      const at = readAttendanceLayout(sh);
      if (!at) return;

      const q = "'" + sh.getName().replace(/'/g, "''") + "'!";
      const want = '=' + q + 'G' + at.netRow;
      const now = summary.getRange(m.row, COL.PAY).getFormula();
      if (now === want) return;

      log.push('  · R' + m.row + ' ' + m.name + ': ' + (now || '(値)') + ' → ' + want);
      if (CONFIG.DRY_RUN) return;
      summary.getRange(m.row, COL.PAY).setFormula(want);
    });
  });
}


/* ------------------------------------------------------------------ *
 * 4. グループ分割と新規2名の追加
 * ------------------------------------------------------------------ */
function newStaffReady() {
  return CONFIG.NEW_STAFF.every(function (s) {
    return typeof s.rate === 'number' && typeof s.commute === 'number';
  });
}

function splitGroupAndAddStaff(ss, summary, log) {
  const layout = readSummaryLayout(summary);
  const target = layout.groups.filter(function (g) { return g.name === CONFIG.SPLIT_GROUP; })[0];
  if (!target) {
    log.push('- グループ「' + CONFIG.SPLIT_GROUP + '」は分割済み');
    return;
  }
  if (target.members.length) {
    log.push('!! 「' + CONFIG.SPLIT_GROUP + '」に既存メンバーがいます。手動で確認してください: '
      + target.members.map(function (m) { return m.name; }).join('、'));
    return;
  }

  log.push('- グループ分割: ' + CONFIG.SPLIT_GROUP + ' → '
    + CONFIG.NEW_STAFF.map(function (s) { return s.store; }).join(' / '));

  // 書式のコピー元。どちらも分割位置より上にあるので、以降の行挿入で位置がずれない
  const headStyleRow = target.row;
  const memberStyleRow = firstMemberRow(layout);

  // 既存の見出し行を1店舗目に流用し、その下へ「メンバー行／見出し行」を挿入していく
  let at = target.row;
  CONFIG.NEW_STAFF.forEach(function (staff, i) {
    const sh = ensureAttendanceSheet(ss, staff, log);

    if (i === 0) {
      log.push('  · R' + at + ' 見出しを ▶ ' + staff.store + ' に変更');
      if (!CONFIG.DRY_RUN) summary.getRange(at, COL.NO).setValue('▶ ' + staff.store);
    } else {
      at += 1;
      log.push('  · R' + at + ' 見出し行を挿入 ▶ ' + staff.store);
      if (!CONFIG.DRY_RUN) {
        summary.insertRowAfter(at - 1);
        copyFormat(summary, headStyleRow, at);
        summary.getRange(at, COL.NO).setValue('▶ ' + staff.store);
      }
    }

    at += 1;
    log.push('  · R' + at + ' メンバー行を挿入 ' + staff.name);
    if (CONFIG.DRY_RUN) return;

    summary.insertRowAfter(at - 1);
    if (memberStyleRow) copyFormat(summary, memberStyleRow, at);
    summary.getRange(at, COL.NO).setValue(0); // No は renumber() で振り直す
    summary.getRange(at, COL.NAME).setValue(staff.name);
    summary.getRange(at, COL.NOTE).clearContent();

    const a = sh ? readAttendanceLayout(sh) : null;
    if (!a) {
      log.push('  !! ' + staff.name + ' の出勤簿を認識できず、集計式は空のままです');
      summary.getRange(at, COL.DAYS, 1, 4).clearContent();
      return;
    }
    const q = "'" + sh.getName().replace(/'/g, "''") + "'!";
    summary.getRange(at, COL.DAYS).setFormula('=' + q + 'D' + a.totalRow);
    summary.getRange(at, COL.HOURS).setFormula('=' + q + 'G' + a.totalRow);
    summary.getRange(at, COL.NIGHT).setFormula('=' + q + 'H' + a.totalRow);
    summary.getRange(at, COL.PAY).setFormula('=' + q + 'G' + a.netRow);
  });

  splitTotalsRow(summary, log);
}

/** 店舗別合計ブロックも同じように1行→2行に分割する */
function splitTotalsRow(summary, log) {
  const layout = readSummaryLayout(summary);
  const t = layout.totalRows.filter(function (x) { return x.name === CONFIG.SPLIT_GROUP; })[0];
  if (!t) {
    log.push('  · 店舗別合計「' + CONFIG.SPLIT_GROUP + '」は分割済み');
    return;
  }

  let at = t.row;
  CONFIG.NEW_STAFF.forEach(function (staff, i) {
    if (i === 0) {
      log.push('  · 店舗別合計 R' + at + ' → ' + staff.store);
      if (!CONFIG.DRY_RUN) summary.getRange(at, COL.NO).setValue(staff.store);
      return;
    }
    at += 1;
    log.push('  · 店舗別合計 R' + at + ' に ' + staff.store + ' を挿入');
    if (CONFIG.DRY_RUN) return;
    summary.insertRowAfter(at - 1);
    copyFormat(summary, t.row, at);
    summary.getRange(at, COL.NO).setValue(staff.store);
  });
}

/**
 * 新規スタッフの出勤簿タブを、雛形タブの複製として用意する。
 * 日々の打刻セルが雛形の氏名を参照している場合は、その氏名を差し替える。
 */
function ensureAttendanceSheet(ss, staff, log) {
  const existing = findAttendanceSheet(ss, '出勤簿　' + staff.name);
  if (existing) {
    log.push('  · 出勤簿タブ既存: ' + existing.getName());
    return existing;
  }

  const tpl = findAttendanceSheet(ss, '出勤簿　' + CONFIG.TEMPLATE_STAFF);
  if (!tpl) {
    log.push('  !! 雛形タブ「' + CONFIG.TEMPLATE_STAFF + '」が見つからず、出勤簿を作成できません');
    return null;
  }
  log.push('  · 出勤簿タブを作成: ' + staff.name + '（雛形 ' + tpl.getName() + '）');
  if (CONFIG.DRY_RUN) return null;

  const sh = tpl.copyTo(ss).setName(staff.name);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(ss.getSheets().length);

  const a = readAttendanceLayout(sh);
  sh.getRange(1, AT.LABEL).setValue(
    String(tpl.getRange(1, AT.LABEL).getValue()).replace(CONFIG.TEMPLATE_STAFF, staff.name));
  sh.getRange(3, AT.LABEL).setValue(staff.rateType + '：' + staff.rate + '円');
  sh.getRange(3, 6).setValue('通勤手当：' + staff.commute + '円×出勤日数　深夜割増：22:00〜翌5:00（法定）');
  sh.getRange(3, AT.RATE).setValue(staff.rate);
  sh.getRange(3, AT.COMMUTE).setValue(staff.commute);

  // 打刻取り込み式が雛形の氏名を持つ場合の差し替え
  const rng = sh.getRange(a.dataStart, 1, a.dataEnd - a.dataStart + 1, sh.getLastColumn());
  const fs = rng.getFormulas();
  let touched = false;
  for (let r = 0; r < fs.length; r++) {
    for (let c = 0; c < fs[r].length; c++) {
      if (fs[r][c] && fs[r][c].indexOf(CONFIG.TEMPLATE_STAFF) >= 0) {
        fs[r][c] = fs[r][c].split(CONFIG.TEMPLATE_STAFF).join(staff.name);
        touched = true;
      }
    }
  }
  if (touched) rng.setFormulas(fs);
  return sh;
}


/* ------------------------------------------------------------------ *
 * No 列の振り直し
 * ------------------------------------------------------------------ */
function renumber(summary, layout, log) {
  let n = 0;
  layout.groups.forEach(function (g) {
    g.members.forEach(function (m) {
      n += 1;
      if (m.no === n) return;
      log.push('  · R' + m.row + ' No ' + m.no + ' → ' + n + ' (' + m.name + ')');
      if (!CONFIG.DRY_RUN) summary.getRange(m.row, COL.NO).setValue(n);
    });
  });
}


/* ------------------------------------------------------------------ *
 * 5. 店舗別合計・総合計の再構築
 * ------------------------------------------------------------------ */
function rebuildTotals(summary, layout, log) {
  if (!layout.totalRows.length) {
    log.push('!! 店舗別合計ブロックが見つかりません');
    return;
  }

  const byStore = {};
  layout.groups.forEach(function (g) { byStore[g.name] = g; });

  layout.totalRows.forEach(function (t) {
    const g = byStore[t.name];
    if (!g) {
      log.push('  ? 店舗別合計「' + t.name + '」に対応するグループがありません');
      return;
    }
    const want = g.members.length
      ? '=SUM(F' + g.members[0].row + ':F' + g.members[g.members.length - 1].row + ')'
      : '=0';
    if (summary.getRange(t.row, COL.PAY).getFormula() === want) return;
    log.push('  · 店舗別合計 R' + t.row + ' ' + t.name + ': ' + want);
    if (!CONFIG.DRY_RUN) summary.getRange(t.row, COL.PAY).setFormula(want);
  });

  // 店舗別合計に無い店舗（分割で増えた分）を警告
  layout.groups.forEach(function (g) {
    const hit = layout.totalRows.filter(function (t) { return t.name === g.name; })[0];
    if (!hit) log.push('  !! 店舗別合計に「' + g.name + '」の行がありません。1行追加してください');
  });

  if (layout.grandTotalRow) {
    const rows = layout.totalRows.map(function (t) { return 'F' + t.row; });
    const want = '=' + rows.join('+');
    if (summary.getRange(layout.grandTotalRow, COL.PAY).getFormula() !== want) {
      log.push('  · 総合計 R' + layout.grandTotalRow + ': ' + want);
      if (!CONFIG.DRY_RUN) summary.getRange(layout.grandTotalRow, COL.PAY).setFormula(want);
    }
  }
}


/* ------------------------------------------------------------------ *
 * 構造の読み取り
 * ------------------------------------------------------------------ */

function getSummarySheet(ss) {
  const byName = ss.getSheetByName('給与一覧');
  if (byName) return byName;
  return ss.getSheets().filter(function (sh) {
    return String(sh.getRange(1, 1).getValue()).indexOf('給与一覧表') >= 0;
  })[0] || null;
}

function findAttendanceSheet(ss, needle) {
  return ss.getSheets().filter(function (sh) {
    const a = String(sh.getRange(1, 1).getValue());
    const b = String(sh.getRange(1, AT.LABEL).getValue());
    return (a + b).indexOf(needle) >= 0;
  })[0] || null;
}

/**
 * 出勤簿タブの構造。認識できなければ null。
 *   3行目      K=単価 / M=通勤手当日額 / B=「時給：〜」または「日給：〜」
 *   5行目      日付ヘッダ
 *   6行目〜    日次データ（G=実働 h, H=深夜 h）
 *   合計行     D=出勤日数 / G=実働合計 / H=深夜合計
 *   明細       B が「基本賃金」「差引支給額」で始まる行
 */
function readAttendanceLayout(sh) {
  const last = sh.getLastRow();
  if (last < 10) return null;
  const labels = sh.getRange(1, AT.LABEL, last, 1).getValues().map(function (r) {
    return String(r[0]).trim();
  });
  if (labels[0].indexOf('出勤簿') !== 0 && labels[0].indexOf('出勤簿') < 0) return null;

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

/**
 * 給与一覧タブの構造。
 *   A列が「▶ 店舗名」の行がグループ見出し、A列が数値の行がメンバー。
 *   「【 店舗別合計 】」以降が集計ブロック。
 */
function readSummaryLayout(sh) {
  const last = sh.getLastRow();
  const a = sh.getRange(1, COL.NO, last, 1).getValues().map(function (r) { return String(r[0]).trim(); });
  const b = sh.getRange(1, COL.NAME, last, 1).getValues().map(function (r) { return String(r[0]).trim(); });

  const groups = [];
  const totalRows = [];
  let grandTotalRow = 0;
  let inTotals = false;
  let cur = null;

  for (let i = 0; i < last; i++) {
    const row = i + 1;
    if (a[i].indexOf('店舗別合計') >= 0) { inTotals = true; cur = null; continue; }

    if (!inTotals) {
      if (a[i].indexOf('▶') === 0) {
        cur = { name: a[i].replace('▶', '').trim(), row: row, members: [] };
        groups.push(cur);
      } else if (cur && a[i] !== '' && !isNaN(Number(a[i])) && b[i] !== '') {
        cur.members.push({ row: row, no: Number(a[i]), name: b[i] });
      }
    } else {
      if (a[i].indexOf('総合計') >= 0) { grandTotalRow = row; }
      else if (a[i] !== '') { totalRows.push({ row: row, name: a[i].trim() }); }
    }
  }
  return { groups: groups, totalRows: totalRows, grandTotalRow: grandTotalRow };
}

function findMember(layout, name) {
  for (let i = 0; i < layout.groups.length; i++) {
    const hit = layout.groups[i].members.filter(function (m) { return m.name === name; })[0];
    if (hit) return hit;
  }
  return null;
}

function findMemberInGroupOtherThan(layout, name, excludeGroup) {
  for (let i = 0; i < layout.groups.length; i++) {
    if (layout.groups[i].name === excludeGroup) continue;
    const hit = layout.groups[i].members.filter(function (m) { return m.name === name; })[0];
    if (hit) return hit;
  }
  return null;
}

/** 罫線・結合・表示形式だけを複製する（数式は持ち込まない） */
function copyFormat(sh, fromRow, toRow) {
  sh.getRange(fromRow, COL.NO, 1, COL.NOTE)
    .copyTo(sh.getRange(toRow, COL.NO), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
}

function firstMemberRow(layout) {
  for (let i = 0; i < layout.groups.length; i++) {
    if (layout.groups[i].members.length) return layout.groups[i].members[0].row;
  }
  return 0;
}
