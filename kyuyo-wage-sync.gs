/* =============================================================
   出勤簿GAS：時給・交通費の変更を、名簿から出勤簿へ反映する

   ※ 先に kyuyo-roster-patch.gs（名簿パッチ）を入れてください。
      このファイルは、その中の関数を使います。

   入社登録アプリで時給や交通費を直したら、各人の出勤簿シートの
     K3（単価）／ M3（通勤手当）／ B3・F3 の表示
     給与明細の「通勤手当」の行（日額×出勤日数／月固定／なし）
     給与明細の「基本賃金」の備考
   を名簿の値に合わせます。金額の式はシートの中で完結しているので、
   ここを直せば支給額まで自動で直ります。

   ―――― 大事な約束 ――――
   ・すでに終わった期間（締めたぶん）には、いっさい触りません。
     今の期間から先だけを直します。過去の給与を書き換えないためです。
   ・変わったところだけ書き換え、「誰の何が いくら→いくら」をログに出します。
   ・土日祝加算や役職手当など、その人ならではの行には触りません。

   ―――― 使い方 ――――
   1. 出勤簿GASに新しいファイルを作って、この中身を貼り付ける
   2. previewWages() を実行して、何が変わるかログで確かめる
   3. applyWages() を実行する（完了と出るまで繰り返す）

   毎朝ひとりでに反映させたいときは、このファイルの installAllTrigger() を
   1回実行してください（名簿パッチの installRosterTrigger() の代わりです）。
   人の追加と時給の変更を、続けて反映します。
   ============================================================= */


/* 毎朝これ1つで、名簿の追加も時給・交通費の変更も反映する */
function syncAll() {
  syncRoster();
  applyWages();
}

function installAllTrigger() {
  removeAllTrigger();
  if (typeof removeRosterTrigger === 'function') removeRosterTrigger();
  ScriptApp.newTrigger('syncAll').timeBased().everyDays(1).atHour(5).create();
  Logger.log('毎朝5時ごろに、名簿の内容を出勤簿へ反映するようにしました。');
  Logger.log('やめるときは removeAllTrigger() を実行してください。');
}

function removeAllTrigger() {
  var ts = ScriptApp.getProjectTriggers();
  var n = 0;
  for (var i = 0; i < ts.length; i++) {
    if (ts[i].getHandlerFunction() === 'syncAll') { ScriptApp.deleteTrigger(ts[i]); n++; }
  }
  if (n) Logger.log('自動実行を ' + n + '件やめました。');
}


// ① 何が変わるかを見るだけ（書き換えなし）
function previewWages() { runWages_(true); }

// ② 実際に直す。完了と出るまで繰り返し実行する
function applyWages() { runWages_(false); }

// 最初からやり直したいとき
function resetWageCheckpoint() {
  PropertiesService.getScriptProperties().deleteProperty('wageSyncIdx');
  Logger.log('次回は最初から処理します。');
}


function runWages_(dryRun) {
  var t0 = new Date().getTime();
  var props = PropertiesService.getScriptProperties();
  var lastIdx = dryRun ? 0 : parseInt(props.getProperty('wageSyncIdx') || '0');

  clearRosterCache();
  var recs = rosterRecords_();
  if (!recs.length) { Logger.log('【要対応】名簿を1人も読めていません'); return; }

  var books = booksFromNow_();
  if (!books.length) { Logger.log('今の期間から先の給与一覧がありません'); return; }
  Logger.log((dryRun ? '【確認だけ】' : '【反映します】')
    + ' 名簿 ' + recs.length + '人 / 対象 ' + books.length + '冊（今の期間から先）');

  var k = lastIdx;
  for (; k < books.length; k++) {
    if (!dryRun && new Date().getTime() - t0 > 4 * 60 * 1000) {
      Logger.log('時間切れです。もう一度 applyWages() を実行して続きをやってください。');
      break;
    }
    var t = books[k];
    var changes = [];
    try {
      var ss = SpreadsheetApp.openById(t.id);
      for (var i = 0; i < recs.length; i++) {
        var line = syncOneWage_(ss, recs[i], dryRun);
        if (line) changes.push(line);
      }
    } catch (e) {
      Logger.log('エラー: ' + t.label + ' / ' + e.message);
      continue;
    }
    Logger.log(t.label + '：' + (changes.length ? '\n    ' + changes.join('\n    ') : '変更なし'));
    Utilities.sleep(600);
    if (dryRun && k >= 2) {   // 確認は先頭3冊で足りる
      Logger.log('（確認は先頭3冊まで。残りも同じように直ります）');
      break;
    }
  }

  if (dryRun) { Logger.log('※確認のみ。直すのは applyWages() です。'); return; }

  props.setProperty('wageSyncIdx', String(k));
  if (k >= books.length) {
    Logger.log('■全' + books.length + '冊 完了！');
    props.deleteProperty('wageSyncIdx');
  } else {
    Logger.log('■' + k + '/' + books.length + '冊まで完了。'
               + 'もう一度 applyWages() を実行してください。');
  }
}


/* ひとりぶんを、名簿の値に合わせる。
   変えたところがあれば説明の文字列を返す。無ければ null。 */
function syncOneWage_(ss, rec, dryRun) {
  var nm = sheetNameOf_(rec);
  var sh = ss.getSheetByName(nm);
  if (!sh) return null;   // シートが無い人は syncRoster() の担当

  var wantRate = Number(rec.payRate) || 0;
  var wantKind = rec.commuteKind === 'monthly' ? 'monthly'
               : rec.commuteKind === 'daily'   ? 'daily' : 'none';
  var wantCom  = wantKind === 'none' ? 0 : (Number(rec.commute) || 0);
  var payLabel = rec.payKind === 'daily' ? '日給' : '時給';

  var nowRate = Number(sh.getRange(3, 11).getValue()) || 0;
  var nowCom  = Number(sh.getRange(3, 13).getValue()) || 0;

  // いまの明細が、通勤手当をどう出しているか
  var rows = wageRows_(sh);
  var nowKind = rows.commuteKind;

  var msg = [];
  if (nowRate !== wantRate) msg.push(payLabel + ' ' + yen_(nowRate) + '→' + yen_(wantRate));
  if (nowCom !== wantCom || nowKind !== wantKind) {
    msg.push('通勤手当 ' + comLabel_(nowKind, nowCom) + '→' + comLabel_(wantKind, wantCom));
  }
  if (!msg.length) return null;
  if (dryRun) return nm + '：' + msg.join(' / ');

  // --- ここから書き換え ---
  sh.getRange(3, 11).setValue(wantRate);
  sh.getRange(3, 13).setValue(wantCom);
  sh.getRange(3, 2).setValue(payLabel + '：' + wantRate + '円');
  sh.getRange(3, 6).setValue(comText_(wantKind, wantCom)
    + '　深夜割増：22:00〜翌5:00（法定）');

  // 基本賃金の備考
  if (rows.kihon > 0) {
    sh.getRange(rows.kihon, 9).setValue(payLabel + wantRate + '円×実働合計');
  }

  // 通勤手当の行
  if (rows.commute > 0) {
    var r = rows.commute;
    if (wantKind === 'daily' && rows.total > 0) {
      sh.getRange(r, 2).setValue('通勤手当（日額×出勤日数）');
      sh.getRange(r, 7).setFormula('=M3*VALUE(SUBSTITUTE(D' + rows.total + ',"日出勤",""))');
      sh.getRange(r, 9).setValue('M3(' + wantCom + '円)×出勤日数');
    } else if (wantKind === 'monthly') {
      sh.getRange(r, 2).setValue('通勤手当（月固定）');
      sh.getRange(r, 7).setValue(wantCom);
      sh.getRange(r, 9).setValue(wantCom + '円/月');
    } else {
      sh.getRange(r, 2).setValue('通勤手当（なし）');
      sh.getRange(r, 7).setValue(0);
      sh.getRange(r, 9).setValue('');
    }
    sh.getRange(r, 7).setNumberFormat('#,##0');
  } else if (wantCom > 0) {
    Logger.log('【要対応】' + nm + ' の明細に通勤手当の行がありません。手で足してください。');
  }

  return nm + '：' + msg.join(' / ');
}


/* 明細のどこに何があるかを調べる。
   合計行より後ろだけを見る（日付ブロックの文字に引っかからないため）。 */
function wageRows_(sh) {
  var last = Math.max(sh.getLastRow(), 1);
  var vals = sh.getRange(1, 2, last, 1).getValues();
  var out = { total: -1, kihon: -1, commute: -1, commuteKind: 'none' };

  for (var r = 0; r < vals.length; r++) {
    var b = String(vals[r][0]).trim();
    if (b === '合計' && out.total < 0) { out.total = r + 1; continue; }
    if (out.total < 0) continue;
    if (out.kihon < 0 && b.indexOf('基本賃金') === 0) out.kihon = r + 1;
    if (out.commute < 0 && b.indexOf('通勤手当') === 0) {
      out.commute = r + 1;
      out.commuteKind = b.indexOf('月固定') >= 0 ? 'monthly'
                      : b.indexOf('日額') >= 0   ? 'daily' : 'none';
    }
  }
  return out;
}

function comLabel_(kind, amt) {
  if (kind === 'daily')   return amt + '円×出勤日数';
  if (kind === 'monthly') return amt + '円（月固定）';
  return 'なし';
}

function comText_(kind, amt) {
  if (kind === 'daily')   return '通勤手当：' + amt + '円×出勤日数';
  if (kind === 'monthly') return '通勤手当：' + amt + '円（月固定）';
  return '通勤手当：なし';
}

function yen_(n) {
  return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '円';
}
