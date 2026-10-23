/**
 * 打刻GAS（プロジェクト「勤怠」）に足すファイル
 * ── 社員のタブ名を、共通名簿からも引けるようにする
 *
 * ＜なぜ要るか＞
 * 打刻GASには、社員の「氏名 → 出勤簿のタブ名」の対応表（SHEET_NAMES）が
 * コードに直接書いてあります。名簿で姓が変わっても、ここは変わりません。
 * そのため打刻のたびに
 *     出勤を記録できませんでした／社員マスタにありません: 加瀬 幸枝
 * と出て、記録がどこにも残らなくなります
 * （伏田幸枝さん → 加瀬幸枝さん でこれが起きました）。
 * 新しく入った社員も、同じ理由で弾かれます。
 *
 * パートのほうは、同じことが起きたときに共通名簿を見にいくように直して
 * あります（paSheetName_）。社員のほうにも同じ逃げ道を作ります。
 *
 * ＜どうするか＞
 *   1. これまでどおり、まず SHEET_NAMES を見る（今まで通っている人はそのまま）
 *   2. 無ければ共通名簿（staff シート）の社員を見る
 *   3. それでも無ければ、これまでどおりエラーにする
 * 名簿が読めないときは 1 だけで動くので、止まりません。
 * 毎回読むと打刻が遅くなるので、5分だけ覚えておきます。
 *
 * ＜貼り方＞
 * 1. 打刻GAS（勤怠）に新しいファイルを作って、この中身を貼る
 * 2. 「コード.gs」の3か所を、下の【差し替え】のとおりに直す
 * 3. デプロイ → デプロイを管理 → ✏️ → 新バージョン → デプロイ
 *
 * ＜【差し替え】コード.gs の3か所＞
 *
 *   ① getStatus の中（139行目あたり）
 *      前： const sheet = getSheet(ss, SHEET_NAMES, data.name);
 *      後： const sheet = staffSheet_(ss, data.name);
 *
 *   ② そのすぐ下（152行目あたり）
 *      前： const prevSheet = getSheet(prevSs, SHEET_NAMES, data.name);
 *      後： const prevSheet = staffSheet_(prevSs, data.name);
 *
 *   ③ 打刻を書き込むところ（199行目あたり）
 *      前： const sheetName = SHEET_NAMES[name];
 *      後： const sheetName = staffSheetName_(name);
 *
 * SHEET_NAMES の表はそのままにしてください（消す必要はありません）。
 *
 * ＜貼ったあとの確認＞
 * 関数の選択欄から checkStaffNames を選んで実行すると、
 * 「誰が打刻できて、どのタブに入るか」が実行ログに出ます。読むだけです。
 */

/* 覚えておく名前。MASTER_SS_ID と MASTER_SHEET は
   打刻GASの「コード.gs」にすでにあるものを、そのまま使います。 */
var STAFF_CACHE = 'staffSheetNames_v1';

/* 氏名を突き合わせる形。「加瀬 幸枝」「加瀬　幸枝」「加瀬幸枝」を同じに扱う。
   出勤簿のタブ名は空白なし（加瀬幸枝）でそろえてあるため。 */
function staffKey_(s) {
  return String(s === null || s === undefined ? '' : s).replace(/[\s　]+/g, '');
}

/* 社員の出勤簿のタブ名。見つからなければ null */
function staffSheetName_(name) {
  if (SHEET_NAMES[name]) return SHEET_NAMES[name];     // これまでどおり
  var key = staffKey_(name);
  if (!key) return null;
  var m = masterStaffNames_();
  return m[key] || null;                               // 共通名簿から
}

/* 社員の出勤簿のシートそのもの */
function staffSheet_(ss, name) {
  var n = staffSheetName_(name);
  return n ? ss.getSheetByName(n) : null;
}

/* 共通名簿から「氏名 → タブ名」を作る。社員（kind=staff）で在籍の人だけ。 */
function masterStaffNames_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(STAFF_CACHE);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) { /* 壊れていたら読み直す */ }
  }

  var map = {};
  try {
    var sh = SpreadsheetApp.openById(MASTER_SS_ID).getSheetByName(MASTER_SHEET);
    if (sh) {
      var vals = sh.getDataRange().getValues();
      var head = vals[0].map(function (v) { return String(v).trim(); });
      var iName   = head.indexOf('name');
      var iKyuyo  = head.indexOf('kyuyoName');
      var iStatus = head.indexOf('status');
      var iKind   = head.indexOf('kind');

      for (var r = 1; r < vals.length; r++) {
        var nm = String(vals[r][iName] || '').trim();
        if (!nm) continue;
        if (iStatus >= 0 && String(vals[r][iStatus]).trim() === 'retired') continue;
        if (iKind < 0 || String(vals[r][iKind]).trim() !== 'staff') continue;   // 社員だけ
        var kn = (iKyuyo >= 0) ? String(vals[r][iKyuyo] || '').trim() : '';
        map[staffKey_(nm)] = staffKey_(kn || nm);
      }
    }
  } catch (e) {
    // 名簿が読めなくても、SHEET_NAMES だけで動かす（打刻を止めない）
  }

  cache.put(STAFF_CACHE, JSON.stringify(map), 300);   // 5分
  return map;
}

/* 名簿を直したのに反映されないときは、これを実行して覚えている内容を捨てる */
function clearStaffCache() {
  CacheService.getScriptCache().remove(STAFF_CACHE);
  Logger.log('覚えていた社員の名簿を捨てました。次の打刻から読み直します。');
}

/* 誰が打刻できて、どのタブに入るかを一覧で出す（実行して結果を見るだけ） */
function checkStaffNames() {
  var m = masterStaffNames_();
  var all = {};
  Object.keys(SHEET_NAMES).forEach(function (k) { all[k] = '表'; });
  Object.keys(m).forEach(function (k) { all[k] = all[k] ? '表＋名簿' : '名簿'; });

  var names = Object.keys(all).sort();
  var out = ['社員で打刻できる人：' + names.length + '名'];
  for (var i = 0; i < names.length; i++) {
    out.push('  ' + names[i] + '　→　タブ「' + staffSheetName_(names[i]) + '」（' + all[names[i]] + '）');
  }
  var msg = out.join('\n');
  Logger.log(msg);
  return msg;
}
