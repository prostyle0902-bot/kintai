#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""決済サービスの手数料（Airペイ／AirペイQR／UberEats／SBペイメント）→ 各店舗タブ

--- これは何か ----------------------------------------------------------
店舗のキャッシュレス決済にかかる手数料。売上から天引きされて入金されるので、
銀行の明細には【手数料を引いたあとの入金額】しか出てこない。手数料そのものは
各決済会社の明細を見るしかない。

    Airペイ（カード・電子マネー）… 支払手数料（AirPay）
    AirペイQR（PayPay・d払い等） … 支払手数料（AirPayQR）
    UberEats                    … 支払手数料（UberEats）
    SBペイメント                 … 支払手数料（Softbank）※焼きたて屋のみ
    出前館                       … 支払手数料（出前館）※demaekan.py が担当

--- ★元データはGoogleドライブにある（2026-09-19 に判明）------------------
Dropboxではなく【Googleドライブ】。店長（brothertak83davab@gmail.com）が
店舗ごとのフォルダに毎月ダウンロードして置いている。

  店舗フォルダ                                     フォルダID
    2.焼きたて屋                                 167oDsyo8DWhnoF-ukLvhp2rUYz3UlcjN
    3.タコとハイボール                            16YDus2nw_4LuTZRrdywr5Huk0m4L4Zzi
    5.大衆酒場りゅうちゃん                         181dl3XWq_RmvC6m2AyDbNyWlbqV_VpAf
    6.韓国酒場ハナ                                1cz_0oONzzXzda6t90b-ZSzWyMDKUwgvJ
    もも焼きJAPAN                                174IWTHzfVKFSMPVOXtyWhHnfno4xIzof

  その下の「取引集計（〜）」フォルダ
    取引集計（Airペイ）      振込明細-<加盟店番号>-<YYYYMMDD>.csv
    取引集計（AirペイQR）    取引集計_<YYYYMMDDhhmmss>.pdf
    取引集計（UberEats）     <月> <年>_<店名>.pdf
    取引集計（SBペイメント）  【…】<YYYYMM>_収納明細書<MMDD-MMDD>_<番号>.pdf
    取引集計（出前館）

--- ★月の切り方が決済サービスごとに違う（重要）--------------------------
  AirペイQR   取引期間が明記されている（例「2026/08/01-2026/08/31」）＝その月
  Airペイ     入金日ベース。8月の明細は 8/5・8/17・8/25 入金＝7/20〜8/18 の利用分
  UberEats    「月次明細 August 2026」＝暦月（Aug 01-31）
  SBペイメント 半月ごとの収納明細書。8月＝0801-0815 と 0816-0831 の2枚を足す
社長が作っていた既存21期PLもこの切り方のまま入れていた（下の実測で確認済み）。
そろえ直すなら指示をください。**いまは既存PLと同じ切り方に合わせている。**

--- 読み取りかた（次の月を足すとき）------------------------------------
Claudeがドライブのコネクタでファイルを開き、下の DATA に【税抜の手数料】を
書き足す。PDFが手元に無い環境でも組み立てられるように数字を焼いてある
（yokocho_data.py・rikuji.py の PDF_DATA と同じ作り）。

  Airペイ CSV … 最終行が合計。「売上合計,手数料,調整,振込額」の2つ目
  AirペイQR PDF … 「合計手数料(税抜)」
  UberEats PDF … 統合月次概要の 総売り上げ − 正味合計
                 （＝Uber手数料の合計＋マーケティング費用の合計＋修正回数の合計）
  SBペイメント PDF … 合計行の手数料【税抜】。半月2枚を足す

--- 21期の9月〜7月について ----------------------------------------------
既存21期PLから入っている（exist_fill.py）。8月だけ既存PLを使わない約束
（利用者指示 2026-08-23）なので、このモジュールが受け持つ。
9月〜7月もドライブの明細で裏を取り直せるが、いまは触っていない。
"""

# (タブ, PL行, 月) -> (税抜, 元ファイル, 備考)
# ★金額は【税抜】。既存21期PLの同じ行も税抜で入っている。
DATA = {
    # ===== 2026-09-19 にドライブで裏を取って入れた8月分 =====
    ("もも焼きJAPAN", "支払手数料（AirPay）", "8月"): (
        11207, "取引集計（Airペイ）/振込明細-052616588018-20260904.csv",
        "CSV最終行の合計 388,400／手数料11,207／振込377,193。"
        "入金日 8/5・8/17・8/25（利用 7/20〜8/18）"),
    ("タコとハイボール", "支払手数料（AirPay）", "8月"): (
        3938, "取引集計（Airペイ）/振込明細-049409198047-20260904.csv",
        "CSV最終行の合計 147,100／手数料3,938／振込143,162。"
        "入金日 8/5・8/17・8/25（利用 7/20〜8/18）"),
    ("韓国酒場ハナ", "支払手数料（AirPay）", "8月"): (
        3773, "取引集計（Airペイ）/振込明細-055210538032-20260904.csv",
        "CSV最終行の合計 135,610／手数料3,773／振込131,837。"
        "入金日 8/5・8/17・8/25（利用 7/20〜8/18）"),
    ("タコとハイボール", "支払手数料（AirPayQR）", "8月"): (
        6854, "取引集計（AirペイQR）/取引集計_20260904161245.pdf",
        "取引期間 2026/08/01-08/31。合計手数料(税抜)6,854。"
        "内訳 PayPay 3,559／d払い1,810／楽天ペイ794／au PAY 634／Smart Code 57"),
    ("タコとハイボール", "支払手数料（UberEats）", "8月"): (
        53978, "取引集計（UberEats）/8 2026_タコとハイボール 神栖横丁店….pdf",
        "月次明細 August 2026（Aug 01-31・53注文）。総売り上げ100,930 −"
        "正味合計46,952 ＝ 53,978。内訳 Uber手数料32,139＋マーケティング費用"
        "21,834＋修正5。★売上100,930は別系統（エアレジ）なので、ここは手数料だけ"),
    ("焼きたて屋", "支払手数料（Softbank）", "8月"): (
        10085, "取引集計（SBペイメント）/202608_収納明細書0801-0815…pdf ＋ "
               "202609_収納明細書0816-0831…pdf",
        "SBペイメントサービス（登録番号T4010401058731）。半月2枚の合計行の"
        "手数料【税抜】3,985＋6,100＝10,085。"
        "前半は売上139,450・入金135,114／後半は売上144,785・入金138,111"
        "（後半には月額 共通決済ｻｰﾋﾞｽ利用料1,980を含む）"),
    # ★ドライブに8月分のファイルがまだ置かれていない。既存21期PLの値を採った。
    #   9月にダウンロードされたら差し替えて確かめること。
    ("韓国酒場ハナ", "支払手数料（AirPayQR）", "8月"): (
        2808, "（ドライブに8月分が未着。既存21期PLの値）",
        "★裏取りできていない1件。取引集計（AirペイQR）のいちばん新しいファイルが"
        "20260804（＝7月分）で、8月分がまだ置かれていない。"
        "ほかの5件はドライブの明細と1円まで一致したので、同じ画面から写した"
        "この値も確からしいとみて入れた。ファイルが届いたら照合すること"),
}

# ドライブの店舗フォルダID（次の月を取りにいくとき用）
DRIVE_FOLDERS = {
    "焼きたて屋":        "167oDsyo8DWhnoF-ukLvhp2rUYz3UlcjN",
    "タコとハイボール":   "16YDus2nw_4LuTZRrdywr5Huk0m4L4Zzi",
    "りゅうちゃん":      "181dl3XWq_RmvC6m2AyDbNyWlbqV_VpAf",
    "韓国酒場ハナ":      "1cz_0oONzzXzda6t90b-ZSzWyMDKUwgvJ",
    "もも焼きJAPAN":     "174IWTHzfVKFSMPVOXtyWhHnfno4xIzof",
}

SRC = "Googleドライブ 店舗フォルダ/取引集計（〜）"


def rows():
    """(タブ, PL行, 月, 税抜, 元ファイル, メモ) を列挙。"""
    for (tab, plrow, m), (val, f, note) in DATA.items():
        yield tab, plrow, m, val, f"{SRC}／{f}", note


def hold_rows():
    """裏が取れていないもの。"""
    for (tab, plrow, m), (val, f, note) in DATA.items():
        if note.startswith("★裏取りできていない"):
            yield (m, tab, f"{plrow} {val:,}", note)


def check(wb=None):
    """行があるか、書き込み先が空かを確かめる。"""
    import build2
    for tab, plrow, m, val, _src, _note in rows():
        assert tab in build2.RIDX, f"タブ「{tab}」が無い"
        assert plrow in build2.RIDX[tab], f"{tab} に「{plrow}」行が無い"
        assert val > 0, f"{tab} {plrow} {m}: 金額が {val}"
    if wb is None:
        return True
    for tab, plrow, m, val, _src, _note in rows():
        c = wb[tab][f"{build2.MCOL[m]}{build2.RIDX[tab][plrow]}"]
        assert not c.value, (f"{tab} {plrow} {m} に既に {c.value} が入っている。"
                             f"決済手数料 {val:,} を足すと二重計上になる")
    return True


if __name__ == "__main__":
    check()
    print("決済サービスの手数料（税抜）\n")
    for tab, plrow, m, val, src, note in sorted(rows()):
        print(f"{m:<4}{tab:<14}{plrow:<24}{val:>8,}")
        print(f"      {src}")
    print(f"\n計 {sum(x[3] for x in rows()):,}円（{len(DATA)}セル）")
    hr = list(hold_rows())
    if hr:
        print(f"\n★保留 {len(hr)}件")
        for m, tab, what, why in hr:
            print(f"  {m} {tab} {what}")
