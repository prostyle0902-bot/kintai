#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""神栖横丁 → 入居している自社4店舗あての「合計請求書」（22期）

読み方・PL行の決め方は 21期の yokocho.py / yokocho_parse.py とまったく同じ。
    店舗側 → 地代家賃・水道光熱費・広告宣伝費 などの【費用】（ここで入れる）
    神栖横丁側 → 【売上】（board の飲食事業部が入れる。ここで足すと二重）
    スポンサー売掛相殺 → 損益に載せない（売掛金と買掛金の相殺）

PDFは yokocho_pdf/22_<YYMM>_<店舗>.pdf に置いて（.gitignore 済み）

    python3 yokocho22.py          … 読んで検算し、DATA22 に貼る形で表示する

DATA22 は手で貼る（21期の yokocho_data.py のような自動生成にはしていない。
1か月4枚ずつなので、読めたことを目で確かめてから貼るほうが安全）。

--- 2026-10-20 9月分（2609月）---------------------------------------------
4枚とも 買掛/22期/2609月/ の【店舗フォルダの外】に「横丁請求書<店舗>.pdf」で届いた。
★9月から「腸内細菌検査料」（700／1,400）が付くようになった。21期には無かった摘要。
  検便の費用なので店舗の「雑費」に置いた（ITEM22）。費目を変えるなら ITEM22 を直す。
"""
import collections
import glob
import os

import yokocho_parse as yp

DIR = yp.DIR
YM2MONTH22 = {"2609": "9月", "2610": "10月", "2611": "11月", "2612": "12月",
              "2701": "1月", "2702": "2月", "2703": "3月", "2704": "4月",
              "2705": "5月", "2706": "6月", "2707": "7月", "2708": "8月"}

# 22期で増えた摘要 → 店舗側のPL行（21期の ITEM2ROW に足す形）
ITEM22 = {
    "腸内細菌検査料": "雑費",
}

SRC22 = {
    "2609": "買掛/22期/2609月/横丁請求書{store}.pdf",
}
_SRCNAME = {"ハナ": "ハナ", "タコハイ": "タコハイ", "りゅうちゃん": "りゅうちゃん", "もも焼き": "もも焼"}


def _row(item):
    if item in ITEM22:
        return ITEM22[item]
    return yp._row_of(item)


# 月 -> PLタブ -> {"摘要別": {摘要: 税抜}, "小計", "消費税", "税込", "No", "src"}
DATA22 = {
    "9月": {
        "韓国酒場ハナ": {
            "摘要別": {"固定家賃": 100000, "共通販促費": 10000, "店舗駐車場": 5000,
                     "店舗電気料金": 35881, "店舗水道料金": 3100, "共益費（計）": 90752,
                     "スポンサー売掛相殺": -64090, "腸内細菌検査料": 700},
            "小計": 181343, "消費税": 18134, "税込": 199477, "No": 8399,
            "src": "買掛/22期/2609月/横丁請求書ハナ.pdf",
        },
        "りゅうちゃん": {
            "摘要別": {"固定家賃": 200000, "共通販促費": 10000, "店舗駐車場": 10000,
                     "店舗電気料金": 23197, "店舗水道料金": 16127, "共益費（計）": 124402,
                     "スポンサー売掛相殺": -1500, "腸内細菌検査料": 1400},
            "小計": 383626, "消費税": 38363, "税込": 421989, "No": 8403,
            "src": "買掛/22期/2609月/横丁請求書りゅうちゃん.pdf",
        },
        "もも焼きJAPAN": {
            "摘要別": {"固定家賃": 98000, "共通販促費": 10000, "店舗駐車場": 5000,
                     "店舗電気料金": 27417, "店舗水道料金": 1969, "共益費（計）": 107052,
                     "スポンサー売掛相殺": -75431, "腸内細菌検査料": 700},
            "小計": 174707, "消費税": 17471, "税込": 192178, "No": 8406,
            "src": "買掛/22期/2609月/横丁請求書もも焼.pdf",
        },
        "タコとハイボール": {
            "摘要別": {"固定家賃": 128000, "共通販促費": 10000, "店舗駐車場": 5000,
                     "店舗電気料金": 21848, "店舗水道料金": 3980, "共益費（計）": 97911,
                     "スポンサー売掛相殺": -3681, "腸内細菌検査料": 1400},
            "小計": 264458, "消費税": 26446, "税込": 290904, "No": 8402,
            "src": "買掛/22期/2609月/横丁請求書タコハイ.pdf",
        },
    },
}


def _row_of_data(item):
    if item == "共益費（計）":
        return yp.KYOEKI
    return _row(item)


def check():
    """DATA22 の検算（明細の和＝小計、小計＋消費税＝税込）"""
    for m, stores in DATA22.items():
        for tab, d in stores.items():
            assert sum(d["摘要別"].values()) == d["小計"], (m, tab, "明細の和≠小計")
            assert d["小計"] + d["消費税"] == d["税込"], (m, tab, "小計＋消費税≠税込")
            assert abs(d["小計"] / 10 - d["消費税"]) < 1, (m, tab, "消費税")
            for item in d["摘要別"]:
                assert _row_of_data(item) is not yp._UNKNOWN, (m, tab, item)


def rows22():
    """(タブ, PL行, 月, 税抜, 元ファイル, メモ)。店舗側の費用だけ。相殺は載せない。"""
    check()
    for m, stores in DATA22.items():
        for tab, d in stores.items():
            by = collections.defaultdict(int)
            for item, v in d["摘要別"].items():
                row = _row_of_data(item)
                if row is None:
                    continue
                by[row] += v
            for row, v in by.items():
                yield (tab, row, m, v, d["src"],
                       f"神栖横丁→{tab} 合計請求書No.{d['No']}（{m}度）。小計{d['小計']:,}＋税{d['消費税']:,}＝税込{d['税込']:,}。"
                       "スポンサー売掛相殺は損益に載せない")


if __name__ == "__main__":
    for path in sorted(glob.glob(os.path.join(DIR, "22_*.pdf"))):
        _p, ym, store = os.path.basename(path)[:-4].split("_", 2)
        saved = dict(yp.ITEM2ROW)
        yp.ITEM2ROW.update(ITEM22)
        try:
            d = yp._read(path)
        finally:
            yp.ITEM2ROW.clear(); yp.ITEM2ROW.update(saved)
        assert sum(d["items"].values()) == d["小計"], (path, "明細の和≠小計")
        kyo = sum(v for k, v in d["items"].items() if _row(k) == yp.KYOEKI)
        items = {k: v for k, v in d["items"].items() if _row(k) != yp.KYOEKI}
        items["共益費（計）"] = kyo
        print(YM2MONTH22[ym], yp.FILE2TAB[store], f"No.{d['No']}", items,
              d["小計"], d["消費税"], d["税込"])
    check()
    print("DATA22 検算OK")
