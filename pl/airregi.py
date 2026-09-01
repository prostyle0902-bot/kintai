#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""エアレジ（Airレジ）の会計明細 → 各店舗の売上

--- どこに置いてあるか ---------------------------------------------------
Dropbox  /※請求書※/会計明細/<期>/<YYMM月>/
    <店名>_20260801-20260831.csv        … エアレジの日別会計明細（CP932）
    <日時>_焼きたて屋税抜.xlsx           … 焼きたて屋だけFCの月間売上集計一覧表
    <日時>_焼きたて屋税込.xlsx
手元へは pl/uriage/<YYMM月>/ に落とす（.gitignore 済み。元データは入れない）。
★フォルダ名の月＝PLの月。買掛と同じ約束（利用者確認 2026-09-01）。

--- ★売上行は店によって税込・税抜が違う ---------------------------------
build2.py の 売上合計(1) は
    消費税行がある店 … 売上（税込） − 消費税
    消費税行がない店 … 売上 そのまま
なので さわら十三里屋だけ【税抜】を入れる。ほかの5店は【税込＋消費税】。

--- 消費税の出し方（利用者確認 2026-09-01「OKです」）--------------------
エアレジのCSVは税率別の【税込】金額しかくれない（売上（10%標準）／売上（8%軽減））。
なので日ごと・税率ごとに逆算して足している。レジは取引ごとに税を計算するので、
月まとめで一発逆算するより日ごとのほうが実額に近い。
★焼きたて屋だけは税抜と税込の両方をもらえるので、その差＝実額を使う。

--- ★ファイルは2種類ある（日別と月まとめ）------------------------------
    pl/uriage/<YYMM月>/<店名>_20260801-20260831.csv   … 日別（1行＝1日）
    pl/uriage/通期/<店名>_202509-202608.csv           … 月まとめ（1行＝1か月）

★月まとめからは【消費税を実額で出せない】。
  エアレジは取引ごとに税を計算しているので、月まとめの金額から逆算すると
  取引単位の切り捨てが再現できず、毎月100〜260円多く出る（実測）。
      りゅうちゃん9月  逆算168,343 / 実額168,109（差+234）
      タコとハイボール9月 逆算 74,542 / 実額 74,280（差+262）
  日別なら日単位まで刻めるので、ぐっと近くなる（りゅうちゃん8月で差9円）。
  なので月まとめの月は【売上（税込）だけ】入れて、消費税行には触らない。
  消費税と さわら十三里屋の税抜は、既存21期PLに入っているエアレジの実額を残す。
  ★2509月〜2607月の日別ファイルは【もらわないことにした】
    （利用者確認 2026-09-01「エアレジの日別ファイル（2509月〜2607月）はやらなくていいや」）。
    なのでこの11か月の消費税は既存21期PLの実額のまま。売上（税込）は月まとめから
    入っていて、既存PLと1円まで一致しているので中身は同じもの。
    さわら十三里屋も同じ理由で税抜が出せないため既存PLのまま。

★売上（税込）は4店とも9月〜7月が既存21期PLと【1円まで完全一致】した（2026-09-01 実測）。
  既存シートの売上はもともとエアレジそのものだったということ。
  なので数字は動かない。変わるのは「どの書類から来たか」だけ。

--- ★確認済みのこと（利用者確認 2026-09-01）-----------------------------
・★タコとハイボールのエアレジには【出前館ぶんが混ざっている】（利用者確認「混ざってますね」）。
  タコとハイボールには「出前館売上」の別行が無いので、これで二重にはならない。
  出前館の手数料は demaekan.py が「支払手数料（出前館）」「支払手数料（出前館返金）」に
  入れているが、あちらは費用なので売上と重ならない。
・焼きたて屋はFCの月間売上集計一覧表に出前館が入っていないので、
  「出前館売上（税込）」「出前館消費税」の別行で持つ（8月ぶんは未着）。
・神栖横丁にエアレジは無い。横丁の売上はテナントへの請求から組む（利用者確認「はい、OK」）
"""
import csv
import glob
import io
import os
import warnings

warnings.filterwarnings("ignore")   # openpyxl の「既定スタイルなし」警告

BASE = os.path.dirname(os.path.abspath(__file__))

# フォルダ名 -> PL列
MONTHS = {"2509月": "9月", "2510月": "10月", "2511月": "11月", "2512月": "12月",
          "2601月": "1月", "2602月": "2月", "2603月": "3月", "2604月": "4月",
          "2605月": "5月", "2606月": "6月", "2607月": "7月", "2608月": "8月"}

# ファイル名の店名 -> (タブ, 売上行, 消費税行 or None)
# ★消費税行が None の店は【税抜】を売上行に入れる（さわら十三里屋だけ）
STORES = {
    "りゅうちゃん": ("りゅうちゃん", "売上", "消費税"),
    "もも焼き": ("もも焼きJAPAN", "売上（税込）", "消費税"),
    "ハナ": ("韓国酒場ハナ", "売上（税込）", "消費税"),
    "タコハイ": ("タコとハイボール", "売上（税込）", "消費税"),
    "十三里屋": ("さわら十三里屋", "売上", None),
    "焼きたて屋": ("焼きたて屋", "売上（税込）", "消費税"),
}


def back(inc, rate):
    """税込→税抜。X + floor(X*rate/100) == 税込 を満たす X（円未満切り捨ての約束）"""
    x = inc * 100 // (100 + rate)
    while x + (x * rate) // 100 < inc:
        x += 1
    while x + (x * rate) // 100 > inc:
        x -= 1
    return x


def _csv(path):
    """エアレジの日別CSV → (税込合計, 消費税, 日数, 10%税込, 8%税込)"""
    rows = list(csv.DictReader(io.open(path, encoding="cp932").read().splitlines()))
    inc = sum(int(r["売上"]) for r in rows)
    t10 = sum(int(r["売上（10%標準）"]) for r in rows)
    t8 = sum(int(r["売上（8%軽減）"]) for r in rows)
    assert t10 + t8 == inc, f"{path}: 税率別の合計 {t10+t8:,} が売上 {inc:,} と違う"
    tax = 0
    for r in rows:                      # ★日ごと・税率ごとに逆算
        for col, rate in (("売上（10%標準）", 10), ("売上（8%軽減）", 8)):
            v = int(r[col])
            if v:
                tax += v - back(v, rate)
    return inc, tax, len(rows), t10, t8


def _yakitate(folder):
    """焼きたて屋のFC月間売上集計一覧表（税抜・税込の2ファイル）→ (税込, 消費税)"""
    import openpyxl
    got = {}
    for kind, pat in (("税込", "*税込*.xlsx"), ("税抜", "*税抜*.xlsx")):
        hit = glob.glob(os.path.join(folder, pat))
        assert len(hit) == 1, f"{folder}: 焼きたて屋の{kind}ファイルが{len(hit)}件"
        ws = openpyxl.load_workbook(hit[0], data_only=True)["ＦＣ店"]
        for r in ws.iter_rows(values_only=True):
            if r[0] and str(r[0]).replace("　", "") == "合計":
                got[kind] = int(str(r[2]).replace(",", ""))
                break
        assert kind in got, f"{hit[0]}: 「合計」の行が見つからない"
    return got["税込"], got["税込"] - got["税抜"]


def _folder(ym):
    return os.path.join(BASE, "uriage", ym)


def months_with_data():
    """日別ファイルがある月（売上も消費税も書類から出せる月）"""
    return [ym for ym in MONTHS if os.path.isdir(_folder(ym))]


def _tsuki():
    """月まとめファイル → {(店名, YYYYMM): 税込売上}。売上（税込）だけ使う。"""
    out = {}
    for path in glob.glob(os.path.join(BASE, "uriage", "通期", "*_*.csv")):
        name = os.path.basename(path).split("_")[0]
        if name not in STORES:
            continue
        for r in csv.DictReader(io.open(path, encoding="cp932").read().splitlines()):
            out[(name, r["集計期間"])] = int(r["売上"])
    return out


YM2 = {v: k.replace("月", "") for k, v in MONTHS.items()}      # "9月" -> "2509"
YM6 = {m: ("20" + YM2[m][:2] + YM2[m][2:]) for m in MONTHS.values()}   # "9月" -> "202509"


def rows():
    """(タブ, PL行, 月, 金額, 元ファイル, メモ) を列挙"""
    # --- ① 月まとめ（売上（税込）だけ。消費税行には触らない）---------------
    daily = {MONTHS[ym] for ym in months_with_data()}
    tsuki = _tsuki()
    for name, (tab, urirow, zeirow) in STORES.items():
        if zeirow is None:
            continue        # ★さわら十三里屋は税抜が要る。月まとめでは出せないので飛ばす
        for m in MONTHS.values():
            if m in daily:
                continue    # 日別があるほうを使う
            inc = tsuki.get((name, YM6[m]))
            if inc is None:
                continue
            yield tab, urirow, m, inc, \
                f"会計明細/21期/{name}_202509-202608.csv", \
                (f"エアレジの月まとめ会計明細。税込{inc:,}。"
                 "★月まとめからは消費税を実額で出せない（取引ごとの切り捨てが再現できない）ので、"
                 "消費税行は既存21期PLに入っているエアレジの実額を残している。"
                 "日別ファイルをもらえば書類ベースにできるが、利用者確認 2026-09-01"
                 "「エアレジの日別ファイル（2509月〜2607月）はやらなくていいや」で"
                 "このままにすると決めた")
    # --- ② 日別（売上も消費税も書類から）----------------------------------
    for ym in months_with_data():
        m = MONTHS[ym]
        folder = _folder(ym)
        for name, (tab, urirow, zeirow) in STORES.items():
            if name == "焼きたて屋":
                hit = glob.glob(os.path.join(folder, "*焼きたて屋税込*.xlsx")) \
                    or glob.glob(os.path.join(folder, "焼きたて屋_税込.xlsx"))
                if not hit:
                    continue
                inc, tax = _yakitate(folder)
                src = f"会計明細/21期/{ym}/（焼きたて屋 FC月間売上集計一覧表 税抜・税込）"
                note = (f"FCの月間売上集計一覧表。税込{inc:,}−税抜{inc-tax:,}＝消費税{tax:,}（実額）。"
                        "出前館は入っていないので別行（出前館売上（税込））で持つ")
            else:
                hit = glob.glob(os.path.join(folder, f"{name}_*.csv"))
                if not hit:
                    continue
                inc, tax, days, t10, t8 = _csv(hit[0])
                src = f"会計明細/21期/{ym}/{os.path.basename(hit[0])}"
                note = (f"エアレジの日別会計明細 {days}日分。税込{inc:,}"
                        f"（10%標準 {t10:,}／8%軽減 {t8:,}）。"
                        f"消費税{tax:,}は日ごと・税率ごとの逆算の合計")
                if tab == "タコとハイボール":
                    note += "。★出前館ぶんも含む（利用者確認 2026-09-01）"
            if zeirow:
                yield tab, urirow, m, inc, src, note + "。売上行は【税込】"
                yield tab, zeirow, m, tax, src, note + "。消費税行"
            else:
                # ★さわら十三里屋は消費税行が無いので税抜を入れる
                yield tab, urirow, m, inc - tax, src, \
                    note + "。★この店は消費税行が無いので【税抜】を入れている"


def check(wb=None):
    import build2
    for tab, plrow, _m, _v, _s, _n in rows():
        assert plrow in build2.RIDX[tab], f"{tab} に「{plrow}」行が無い"
    if wb is None:
        return
    for tab, plrow, m, v, _s, _n in rows():
        c = wb[tab][f"{build2.MCOL[m]}{build2.RIDX[tab][plrow]}"]
        assert not c.value, (f"{tab} {plrow} {m} に既に {c.value} が入っている。"
                             f"エアレジの {v:,} を足すと二重計上になる")


if __name__ == "__main__":
    check()
    got = list(rows())
    print(f"エアレジの売上セル {len(got)} 件 ／ {len(months_with_data())} か月ぶん\n")
    print(f"{'月':>4} {'タブ':<12}{'行':<12}{'金額':>12}")
    for tab, plrow, m, v, _s, _n in got:
        print(f"{m:>4} {tab:<12}{plrow:<12}{v:>12,}")
    uri = sum(v for _t, r, _m, v, _s, _n in got if r != "消費税")
    zei = sum(v for _t, r, _m, v, _s, _n in got if r == "消費税")
    print(f"\n売上行の合計 {uri:,}円 ／ 消費税行の合計 {zei:,}円")
