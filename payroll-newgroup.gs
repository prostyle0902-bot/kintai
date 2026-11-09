/* =========================================================
   給与一覧表に、新しい現場の見出しを作る
   （シフト自動作成アプリ用の Apps Script に、ファイルを足して使います）

   ■ なにをするもの？

   給与一覧表は、現場ごとに「▶ 焼きたて屋」のような見出しがあって、
   その下に人が並ぶ作りになっています。

     ▶ 焼きたて屋
      1  小林メグ    22日出勤  …
      2  荒宏信      23日出勤  …
     ▶ 横丁清掃
      9  橋本桂子    13日出勤  …

   新しく人を足すときは、同じ現場にすでにいる人の行をまるごと複製して、
   氏名だけ差し替えています。計算式を組み立て直さずに済むやり方です。

   ところが、まったく新しい現場の「1人目」だけは、これができません。
   見出しも無ければ、複製する見本の行も無いためです。
   「『○○』の見出し（▶）が給与一覧表に見つかりません」で止まります。

   このファイルは、その1人目のために
     ・「▶ ○○」の見出し
     ・その人の行（計算式つき）
     ・【 店舗別合計 】の行
   の3つを作ります。2人目からは、いつもの流れで自動で足せます。

   ■ 使い方

   1. 下の NEW_GROUP を、足したい現場と人に書き換える
   2. 関数の選択欄から previewNewGroup を選んで「実行」
      → 実行ログに「何をするつもりか」だけ出ます。まだ何も変えません
   3. 中身を確かめてから makeNewGroup を選んで「実行」
      → 実際に作ります

   デプロイは要りません。エディタから実行するだけで、保存した内容で動きます。

   何度実行しても大丈夫です。すでに見出しがある期間は飛ばします。
   ========================================================= */

var NEW_GROUP = {
  // 給与一覧表に作る見出しの名前（「▶ 」は要りません）
  group: '波崎DC仮設事務所',

  // その現場の1人目。入社登録の氏名と同じに
  name: '奥田京子',

  /* 見本にする人。この人の出勤簿シートと、給与一覧表の行を複製します。
     空にすると、自動で選びます。選ぶのは
       ・作る人と給与の種類（時給／日給）が同じ
       ・固定手当が無い
       ・修正給与（手入力）が入っていない
     方です。どれかが違う人を見本にすると、計算式や金額がそのまま
     付いてきて、新しい人の給与が狂います。 */
  model: '',

  // この日より前の期間は作りません（その人の入社日を入れてください）
  from: '2026-11-11',
};


/* ---------- 下ごしらえ ---------- */

// 給与一覧表で、いちばん下にある「人の行」を探す
function ngLastMemberRow_(vals) {
  for (var r = vals.length - 1; r >= 0; r--) {
    if (/^\d+$/.test(String(vals[r][0] || '').trim())) return r + 1;
  }
  return -1;
}

// 「▶」の付いた見出し行を1つ探す（見た目をまねるため）
function ngHeadRow_(vals) {
  for (var r = 0; r < vals.length; r++) {
    if (vals[r].join(' ').indexOf('▶') >= 0) return r + 1;
  }
  return -1;
}

/* 見本にする人を決める。
   ・出勤簿シートがあること
   ・給与一覧表に行があること
   ・固定手当が入っていないこと（複製すると手当までついてくる）
   この3つを満たす人を、下から探します。 */
function ngPickModel_(ss, list, vals, want, payKind) {
  if (want) {
    var row0 = -1;
    for (var i = 0; i < vals.length; i++) {
      if (!/^\d+$/.test(String(vals[i][0] || '').trim())) continue;
      if (payrollNameKey_(vals[i][1]) === payrollNameKey_(want)) { row0 = i + 1; break; }
    }
    if (row0 < 0) return { ok: false, why: '見本にする ' + want + ' さんが給与一覧表にいません' };
    var shw = payrollStaffSheet_(ss, want);
    if (!shw) return { ok: false, why: '見本にする ' + want + ' さんの出勤簿シートがありません' };
    var ex = ngHasExtra_(shw);
    if (ex) {
      return { ok: false, why: '見本にする ' + want + ' さんは' + ex
        + 'ため、複製すると新しい人にもその金額が付いてしまいます' };
    }
    var kw = ngPayKind_(shw);
    if (payKind && kw && kw !== payKind) {
      return { ok: false, why: '見本にする ' + want + ' さんは' + ngKindLabel_(kw)
        + 'の計算です。作る人は' + ngKindLabel_(payKind) + 'なので、給与の計算式が合いません' };
    }
    return { ok: true, name: String(vals[row0 - 1][1] || '').trim(), row: row0,
             kind: kw, why: '指定されたため' };
  }

  var otherKind = null;
  for (var r = vals.length - 1; r >= 0; r--) {
    if (!/^\d+$/.test(String(vals[r][0] || '').trim())) continue;
    var nm = String(vals[r][1] || '').trim();
    if (!nm) continue;
    var sh = payrollStaffSheet_(ss, nm);
    if (!sh) continue;
    if (ngHasExtra_(sh)) continue;
    var k = ngPayKind_(sh);
    /* 給与の種類が違う人は見本にできない。
       基本賃金の計算式そのものが変わるため（時給×実働 ／ 日給×出勤日数）。 */
    if (payKind && k && k !== payKind) { if (!otherKind) otherKind = nm + '（' + ngKindLabel_(k) + '）'; continue; }
    return { ok: true, name: nm, row: r + 1, kind: k || payKind,
             why: ngKindLabel_(k || payKind) + 'で、固定手当も修正給与も無いため' };
  }
  return { ok: false, why: '見本にできる人がいません。'
    + ngKindLabel_(payKind) + 'で、固定手当も修正給与も無いパートの方が、この期間の給与一覧表にいません'
    + (otherKind ? '（' + otherKind + ' はいますが、給与の種類が違います）' : '') };
}

/* その出勤簿が、時給・日給・月給のどれで計算しているかを見る。

   給与明細の「基本賃金（時給×実働時間）」「基本賃金（日給×出勤日数）」
   という項目名で見分けます。計算式そのものがここで変わるので、
   種類の違う人を見本にすると、給与が丸ごと間違います。
   （日給の人を見本にして時給の人を作ると、時給×出勤日数になる） */
function ngPayKind_(sheet) {
  var last = Math.min(sheet.getLastRow(), 70);
  if (last < 2) return '';
  var col = sheet.getRange(1, 2, last, 1).getValues();
  for (var i = 0; i < col.length; i++) {
    var t = String(col[i][0] || '');
    if (t.indexOf('基本賃金') < 0) continue;
    if (t.indexOf('時給') >= 0) return 'hourly';
    if (t.indexOf('日給') >= 0) return 'daily';
    if (t.indexOf('月給') >= 0) return 'monthly';
  }
  // 項目名で分からないときは、見出し（B3あたり）の書き出しで見る
  var head = String(sheet.getRange(3, 2).getValue() || '');
  if (head.indexOf('時給') === 0) return 'hourly';
  if (head.indexOf('日給') === 0) return 'daily';
  if (head.indexOf('月給') === 0) return 'monthly';
  return '';
}

function ngKindLabel_(k) {
  return k === 'daily' ? '日給' : k === 'monthly' ? '月給' : k === 'hourly' ? '時給' : '不明';
}

/* 見本にすると、その人のお金がそのまま付いてきてしまうものがあるか。

   1) 固定手当の行
      給与明細の項目名（B列）に「手当」が入っていて、通勤手当・深夜割増
      でないもの（例「シングル手当」）。複製すると手当ごと複製されます。

   2) 修正給与（手入力）に数字が入っている
      ここは計算式ではなく手で入れた金額なので、複製したときに消えません。
      前の人の調整額が、新しい人の給与に乗ってしまいます。

   どちらかがあれば、その人は見本にしません。 */
function ngHasExtra_(sheet) {
  var last = Math.min(sheet.getLastRow(), 70);
  if (last < 2) return '';
  var b = sheet.getRange(1, 2, last, 1).getValues();
  var g = sheet.getRange(1, 7, last, 1).getValues();   // G列＝給与明細の金額
  for (var i = 0; i < b.length; i++) {
    var t = String(b[i][0] || '');
    if (t.indexOf('修正給与') >= 0 && Number(g[i][0])) return '修正給与が入っている';
    if (t.indexOf('手当') < 0) continue;
    if (t.indexOf('通勤') >= 0 || t.indexOf('深夜') >= 0) continue;
    return t + 'がある';
  }
  return '';
}

// 【 店舗別合計 】の、いちばん下の現場の行を探す
function ngTotalRow_(list) {
  var vals = list.getRange(1, 1, Math.max(list.getLastRow(), 1), 6).getValues();
  var head = -1;
  for (var r = 0; r < vals.length; r++) {
    if (String(vals[r][0] || '').indexOf('店舗別合計') >= 0) { head = r + 1; break; }
  }
  if (head < 0) return { head: -1, last: -1 };

  var last = -1;
  for (var s = head; s < vals.length; s++) {
    var a = String(vals[s][0] || '').trim();
    if (!a) continue;
    if (a.indexOf('総合計') >= 0) break;
    last = s + 1;
  }
  return { head: head, last: last };
}


/* ---------- 中身を見るだけ（何も変えません） ---------- */

function previewNewGroup() { ngRun_(true); }

/* ---------- 実際に作る ---------- */

function makeNewGroup() { ngRun_(false); }


function ngRun_(dryRun) {
  var group = String(NEW_GROUP.group || '').trim();
  var name = String(NEW_GROUP.name || '').trim();
  var from = String(NEW_GROUP.from || '').trim();
  if (!group || !name) {
    Logger.log('NEW_GROUP の group と name を入れてください。');
    return;
  }

  Logger.log(dryRun ? '=== 下見（何も変えません）===' : '=== 作ります ===');
  Logger.log('現場：' + group + '　1人目：' + name + '　' + from + ' 以降の期間');

  // 名簿から、その人の単価や通勤手当を取る
  var rec = ngStaffRec_(name);
  if (!rec) {
    Logger.log('✗ 名簿に ' + name + ' さんがいません。入社登録で先に登録してください。');
    return;
  }
  Logger.log('名簿：' + (rec.payKind === 'daily' ? '日給' : '時給') + rec.payRate + '円'
    + '／通勤手当 ' + (rec.commute || 0) + '円'
    + (rec.commuteKind === 'monthly' ? '（月固定）' : rec.commuteKind === 'none' ? '（なし）' : '×出勤日数'));
  Logger.log('');

  var files = payrollFiles_().files;
  var did = 0, skipped = 0, failed = 0;

  for (var i = files.length - 1; i >= 0; i--) {     // 古い期間から順に
    var f = files[i];
    if (from && f.end < from) { continue; }         // その人が入る前の期間は触らない

    var r;
    try { r = ngOneFile_(f, group, name, rec, dryRun); }
    catch (e) { r = { status: 'error', message: String(e) }; }

    if (r.status === 'skip') { skipped++; Logger.log('－ ' + f.label + '：' + r.message); }
    else if (r.status === 'error') { failed++; Logger.log('✗ ' + f.label + '：' + r.message); }
    else { did++; Logger.log((dryRun ? '○ ' : '✓ ') + f.label + '：' + r.message); }
  }

  Logger.log('');
  Logger.log((dryRun ? '作る予定：' : '作った：') + did + '件／'
    + 'そのまま：' + skipped + '件／できなかった：' + failed + '件');
  if (dryRun && did) {
    Logger.log('');
    Logger.log('問題なければ、関数の選択欄で makeNewGroup を選んで実行してください。');
  }
}

// 名簿（staff シート）から1人ぶん取る
function ngStaffRec_(name) {
  var sh = getStaffSheet_();
  var head = staffHeader_(sh);
  var list = readStaff_(sh, head);
  var key = payrollNameKey_(name);
  for (var i = 0; i < list.length; i++) {
    var n = list[i].kyuyoName || list[i].name;
    if (payrollNameKey_(n) === key) return list[i];
  }
  return null;
}


function ngOneFile_(f, group, name, rec, dryRun) {
  var ss = SpreadsheetApp.openById(f.id);
  var list = payrollListSheet_(ss);
  if (!list) return { status: 'error', message: '給与一覧表のシートが見つかりません' };

  if (payrollGroup_(list, group)) {
    return { status: 'skip', message: '「' + group + '」の見出しはもうあります' };
  }

  var vals = payrollListVals_(list);
  var lastMember = ngLastMemberRow_(vals);
  if (lastMember < 0) return { status: 'error', message: '人の行が1行も見つかりません' };
  var headSample = ngHeadRow_(vals);
  if (headSample < 0) return { status: 'error', message: '「▶」の見出しが1つも見つかりません' };

  var payKind = (rec.payKind === 'daily') ? 'daily'
              : (rec.payKind === 'monthly') ? 'monthly' : 'hourly';
  var model = ngPickModel_(ss, list, vals, NEW_GROUP.model, payKind);
  if (!model.ok) return { status: 'error', message: model.why };

  var tot = ngTotalRow_(list);
  if (tot.last < 0) return { status: 'error', message: '【 店舗別合計 】が見つかりません' };

  if (dryRun) {
    return { status: 'ok', message: lastMember + '行目の下に「▶ ' + group + '」と '
      + name + ' さんの行を作り（見本：' + model.name + '／' + model.why + '）、'
      + '店舗別合計の ' + tot.last + '行目の上に合計の行を入れます'
      + (payrollStaffSheet_(ss, name) ? '。出勤簿シートはすでにあります' : '。出勤簿シートも作ります') };
  }

  var lastCol = list.getLastColumn();

  /* --- 1. 出勤簿シートを作る（すでにあれば触らない） --- */
  var madeSheet = false;
  var sheet = payrollStaffSheet_(ss, name);
  if (!sheet) {
    var tpl = payrollStaffSheet_(ss, model.name);
    if (!tpl) return { status: 'error', message: model.name + ' の出勤簿が見つかりません' };
    sheet = tpl.copyTo(ss).setName(name);
    if (sheet.isSheetHidden()) sheet.showSheet();
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(tpl.getIndex() + 1);
    payrollFillSheet_(sheet, model.name, name, rec, null);
    madeSheet = true;
  }

  /* --- 2. 見出しと、その人の行を作る ---
     見本の行の計算式は、複製すると参照が1行ずれるので、
     ずれる前の式を控えてから入れ直します（payrollEnsureRow_ と同じやり方）。 */
  var srcFs = list.getRange(model.row, 1, 1, lastCol).getFormulas()[0];

  list.insertRowsAfter(lastMember, 2);
  var headRow = lastMember + 1;
  var memberRow = lastMember + 2;

  // 見出し行：他の見出しから「見た目だけ」まねて、名前を入れる
  list.getRange(headSample, 1, 1, lastCol)
    .copyTo(list.getRange(headRow, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  list.getRange(headRow, 1).setValue('▶ ' + group);

  // 人の行：見本の行をまるごと複製してから、氏名と計算式を差し替える
  list.getRange(model.row, 1, 1, lastCol).copyTo(list.getRange(memberRow, 1));
  list.getRange(memberRow, 2).setValue(name);
  for (var c = 0; c < lastCol; c++) {
    if (!srcFs[c] || srcFs[c].indexOf('!') < 0) continue;
    list.getRange(memberRow, c + 1).setFormula(payrollRefix_(srcFs[c], model.name, name));
  }
  try { list.showRows(headRow, 2); } catch (e) { /* 隠れた行の下に足したとき用 */ }

  /* --- 3. 店舗別合計に行を足す ---
     上で2行入れたので、合計の位置も2行ずれています。探し直します。
     いちばん下の現場の「上」に入れるのは、総合計の =SUM(…) の範囲の
     内側に入れるためです。外側（下）に足すと合計に入りません。 */
  var tot2 = ngTotalRow_(list);
  if (tot2.last < 0) {
    return { status: 'ok', message: '見出しと行は作りましたが、店舗別合計が見つからず足せませんでした'
      + (madeSheet ? '（出勤簿シートは作りました）' : '') };
  }
  list.insertRowsBefore(tot2.last, 1);
  /* 入れたぶん、もとの行は1つ下がっている。
     そこから「見た目だけ」写して、中身はこちらで入れる。 */
  list.getRange(tot2.last + 1, 1, 1, 6)
    .copyTo(list.getRange(tot2.last, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  list.getRange(tot2.last, 1).setValue(group);
  list.getRange(tot2.last, 6).setFormula('=SUM(F' + memberRow + ')');

  payrollRenumber_(list);

  return { status: 'ok', message: '「▶ ' + group + '」と ' + name + ' さんの行を作りました'
    + '（見本：' + model.name + '）'
    + (madeSheet ? '。出勤簿シートも作りました' : '。出勤簿シートはすでにありました') };
}


/* =========================================================
   おまけ：壊れた合計を探す（読むだけ）

   人の行を消すと、店舗別合計の =SUM(…) が =SUM(#REF!) になり、
   そのせいで総合計まで #REF! になることがあります。
   どの期間のどこが壊れているかを、まとめて出します。
   ========================================================= */
function checkTotals() {
  var files = payrollFiles_().files;
  var bad = 0;

  for (var i = files.length - 1; i >= 0; i--) {
    var f = files[i];
    var hits = [];
    try {
      var ss = SpreadsheetApp.openById(f.id);
      var list = payrollListSheet_(ss);
      if (!list) { Logger.log('✗ ' + f.label + '：給与一覧表が見つかりません'); bad++; continue; }
      var tot = ngTotalRow_(list);
      if (tot.head < 0) continue;

      var n = Math.max(list.getLastRow(), 1);
      var names = list.getRange(1, 1, n, 1).getValues();
      var fs = list.getRange(1, 6, n, 1).getFormulas();
      var vs = list.getRange(1, 6, n, 1).getDisplayValues();
      for (var r = tot.head - 1; r < n; r++) {
        var label = String(names[r][0] || '').trim();
        if (!label) continue;
        if (String(fs[r][0]).indexOf('#REF!') >= 0 || String(vs[r][0]).indexOf('#REF!') >= 0) {
          hits.push('  ' + (r + 1) + '行目 ' + label + '：' + (fs[r][0] || vs[r][0]));
        }
      }
    } catch (e) {
      Logger.log('✗ ' + f.label + '：' + e); bad++; continue;
    }
    if (hits.length) {
      bad++;
      Logger.log('✗ ' + f.label);
      for (var h = 0; h < hits.length; h++) Logger.log(hits[h]);
    }
  }

  Logger.log('');
  Logger.log(bad ? '壊れている期間：' + bad + '件' : 'すべて正常です');
  if (bad) {
    Logger.log('直し方：その行の F列を選んで、=SUM(F○:F○) を入れ直してください。');
    Logger.log('○には、その現場の見出し（▶）の下に並ぶ人の行を入れます。');
    Logger.log('その現場に人が1人もいないときは、=0 と入れておけば総合計は直ります。');
  }
}
