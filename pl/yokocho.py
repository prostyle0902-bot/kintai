#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""神栖横丁 → 入居店舗への「合計請求書」（社内請求）

--- これは何か ----------------------------------------------------------
神栖横丁（Prostyle株式会社）が、横丁に入っている自社店舗あてに毎月出す請求書。
家賃・駐車場・電気・水道・共益費などをまとめて請求している。
発行元も宛先もProstyle株式会社なので、社内の部門間取引。

    店舗側 → 地代家賃・水道光熱費・広告宣伝費 などの【費用】
    神栖横丁側 → 【売上】

Dropbox: /※請求書※/買掛/21期/YYMM月/<店舗フォルダ>/2026年MM月請求分合計請求書.pdf
テキスト層あり（pdftotext -layout でそのまま読める）。

--- 請求書の構造 --------------------------------------------------------
「摘要／数量／単位／単価／金額」の表。数量0の行は請求なし。
末尾に スポンサー売掛相殺・その他値引き のマイナス行があり、
    小計 = 明細の合計 − 相殺・値引き
    合計（税込） = 小計 + 消費税10%

--- ★10か月ぶんに増やした（2026-08-22）------------------------------------
利用者指示「他の月もDropbox内にあると思うので確認してみて」。
最初は7月分4枚しか見つけられず手で起こしていたが、他の月はファイル名が
「横丁　<店舗><月>末.pdf」で「合計請求書」と違うため名前では探せなかった。
全文検索で「スポンサー売掛相殺」を引いたら21期40枚ぜんぶ出てきた。
    2509月〜2606月 … 横丁　<店舗><月>末.pdf（各店舗フォルダの中）
    2607月         … 2026年07月請求分合計請求書.pdf
いまは yokocho_parse.py がPDFを直に読み、yokocho_data.py に書き出している。
このモジュールはそれを読むだけ。手で起こした7月分の数字とは全部一致した。

★手元に無いのは2か月ぶん
    2510月（11月末請求）… 4店舗ともDropboxに無い。10月列が空くのはこのため
    2608月（9月末請求） … 21期の最終月。まだ発行されていない

--- 検算 -----------------------------------------------------------------
40枚とも「明細の合計＝小計」「小計の10%＝消費税（四捨五入）」
「小計＋消費税＝合計」を通している。

--- ★2511月の3枚だけ相殺が税込だった（税抜にそろえた）--------------------
ふつうは  明細の和（相殺こみ）＝小計（税抜）  なのだが、
2511月のもも焼き・タコハイ・ハナの3枚は
    プラス明細の和 ×1.1 − 相殺 ＝ 合計（税込）
という組み方で、請求書の小計には 合計÷1.1 が書いてあった。
PLに入れるのは税抜なので請求書の小計を正として相殺の側を直している。
    もも焼き ▲112,133（税込）→ ▲101,939（税抜相当）
    タコハイ  ▲63,097 → ▲57,361 ／ ハナ ▲151,350 → ▲137,591
りゅうちゃんの2511月は普通の組み方だったので触っていない。

--- ★神栖横丁側の売上は入れていない -------------------------------------
2026-08-22 に解決。board.py が売掛の請求書CSVから神栖横丁の売上を
9月〜7月とも入れている（7月は2,965,826）。テナント全店ぶんが入っているので、
ここで読んだ4店舗ぶんを足すと二重計上になる。

--- スポンサー売掛相殺・その他値引き（利用者判断 2026-08-20）------------
「そのままでいい。費用のマイナスで」。雑収入にはしない。
行は「その他経費」。既存スプシにもマイナス計上の前例がある
（韓国酒場ハナ 5月 ▲2,727）。
"""
TAB_OWNER = "神栖横丁"          # 請求を出す側（売上）

# 請求書フォルダ名 → PLタブ名
FOLDER2TAB = {
    "ハナ":       "韓国酒場ハナ",
    "タコハイ":    "タコとハイボール",
    "りゅうちゃん": "りゅうちゃん",
    "もも焼き":     "もも焼きJAPAN",
}

# 摘要 → 店舗側のPL行は yokocho_parse.ITEM2ROW / _row_of が持っている。
# 「共益費/〜」と「日常清掃（等）」はまとめて 地代家賃（共益費） へ。
from yokocho_parse import ITEM2ROW, _row_of        # noqa: E402,F401

# 月 -> 店舗 -> 各値。yokocho_parse.py が請求書PDFから起こしている。
# 中身をなおすときは yokocho_data.py ではなくPDFのほうを見ること。
from yokocho_data import DATA                      # noqa: E402

# 手元に請求書が無い月。PLの該当列はこのモジュールでは埋まらない。
MISSING = {
    "10月": "2510月（11月末請求）の4店舗ぶんがDropboxに無い",
    "8月": "2608月（9月末請求）は21期の最終月でまだ発行されていない",
}

# 共益費の内訳（明細ログ用。合計だけをPLに入れる）
KYOEKI_ITEMS = [
    "日常清掃等", "共通洗浄エリア管理費", "ゴミ運搬処理費", "ゴミ運搬処分費",
    "共有部の電気料金", "共有部のガス料金", "共有部の水道料金", "共有部の水光熱費",
    "共有部の清掃管理費", "清掃管理費", "共通HP管理費", "共通ＨＰ管理費",
    "通信回線＆機器使用料", "施設警備費", "館内BGM利用料",
    "LINE公式アカウント利用料", "各種申請手数料", "その他一般管理費等", "その他",
]


def rows():
    """(タブ, PL行, 月, 税抜金額, 元ファイル) を列挙。店舗側の費用だけ。"""
    for month, stores in DATA.items():
        for tab, d in stores.items():
            by_row = {}
            for item, amount in d["摘要別"].items():
                by_row.setdefault(_row_of(item), 0)
                by_row[_row_of(item)] += amount
            for plrow, amount in by_row.items():
                yield tab, plrow, month, amount, d["src"]


def detail(tab, plrow, month):
    """明細ログ用。その行に足しこんだ摘要の内訳を「名前 金額」で返す。"""
    d = DATA.get(month, {}).get(tab)
    if not d:
        return ""
    parts = [f"{k} {v:,}" for k, v in d["摘要別"].items() if _row_of(k) == plrow]
    return "／".join(parts)


def hold_rows():
    """保留: 請求書が手元に無い月だけ。"""
    for month, why in MISSING.items():
        yield (month, "神栖横丁", "社内請求（4店舗あて）",
               f"横丁から入居店舗への請求書が無いので、この月は"
               f"家賃・駐車場・共通販促費・電気・水道・共益費・相殺を入れていない。{why}")


def check(wb=None):
    """請求書の内部整合と、PLに書き込む先が空かどうか。"""
    for month, stores in DATA.items():
        for tab, d in stores.items():
            got = sum(d["摘要別"].values())
            assert got == d["小計"], \
                f"{month} {tab}: 明細 {got:,} ≠ 小計 {d['小計']:,}"
            assert abs(d["小計"] * 0.1 - d["消費税"]) < 1, \
                f"{month} {tab}: 小計の10% ≠ 消費税 {d['消費税']:,}"
            assert d["小計"] + d["消費税"] == d["税込"], \
                f"{month} {tab}: 小計＋消費税 ≠ 合計 {d['税込']:,}"
    # 請求書No は月を追うごとに増えるはず（枚数の取り違え検知）
    seq = [(m, t, d["No"]) for m, s2 in DATA.items() for t, d in s2.items()]
    assert len({x[2] for x in seq}) == len(seq), \
        f"請求書Noが重複している: {sorted(seq, key=lambda x: x[2])}"
    if wb is None:
        return True
    import build2
    for tab, plrow, month, amount, _src in rows():
        assert plrow in build2.RIDX[tab], f"{tab} に「{plrow}」行が無い"
    return True


if __name__ == "__main__":
    check()
    n = sum(len(v) for v in DATA.values())
    print(f"神栖横丁の社内請求書 {n}枚（{len(DATA)}か月 × 4店舗）検算すべて通過\n")
    import collections
    agg = collections.defaultdict(dict)
    for tab, plrow, month, amount, _src in rows():
        agg[(tab, plrow)][month] = amount
    months = list(DATA)
    print(f"{'店舗':<14}{'PL行':<22}" + "".join(f"{m:>10}" for m in months) + f"{'計':>11}")
    for (tab, plrow), v in sorted(agg.items()):
        print(f"{tab:<14}{plrow:<22}"
              + "".join(f"{v.get(m,0):>10,}" for m in months)
              + f"{sum(v.values()):>11,}")
    print("\n保留:")
    for m, tab, item, why in hold_rows():
        print(f"  [{m}] {item} … {why}")
