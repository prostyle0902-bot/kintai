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
    return [ym for ym in MONTHS if os.path.isdir(_folder(ym))]


def rows():
    """(タブ, PL行, 月, 金額, 元ファイル, メモ) を列挙"""
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
