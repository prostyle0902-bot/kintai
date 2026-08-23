#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""神栖横丁 → 入居店舗あて請求書PDF の読み取り

--- 使い方 ---------------------------------------------------------------
    python3 yokocho_parse.py            … 読んだ結果を表示して検算
    python3 yokocho_parse.py --gen      … yokocho_data.py を書き出す

PDFは yokocho_pdf/<請求月>_<店舗>.pdf に置く（.gitignore 済み。容量が大きい）。
書き出した yokocho_data.py のほうをリポジトリに入れて、yokocho.py はそれを読む。
PDFが手元に無い環境でも動くようにするため。

--- Dropboxでの置き場所 --------------------------------------------------
2509月〜2606月  /※請求書※/買掛/21期/YYMM月/確認済/<店舗>/…/横丁　<店舗><月>末.pdf
2607月          /※請求書※/買掛/21期/2607月/<店舗>/2026年07月請求分合計請求書.pdf
★ファイル名がバラバラなので「横丁」では探しきれない。
  全文検索で「スポンサー売掛相殺」を引くと確実に全部出る。
★2510月（＝11月末請求）だけ4店舗ともDropboxに無い。
★2608月（＝9月末請求）は21期の最終月。まだ発行されていない。

--- 請求書の構造 ---------------------------------------------------------
「摘要／数量／単位（件）／単価／金額」の表。数量0の行は請求なし。
末尾に スポンサー売掛相殺・その他値引き のマイナス行があり、
    小計 ＝ 明細の合計（マイナス行を含む）
    合計（税込） ＝ 小計 ＋ 消費税10%
発行元も宛先もProstyle株式会社なので社内の部門間取引。
    店舗側 → 地代家賃・水道光熱費・広告宣伝費などの【費用】
    神栖横丁側 → 【売上】（board.py が売掛CSVから入れている）

--- ★摘要の表記ゆれ ------------------------------------------------------
同じものが月によって違う名前で出る。ITEM2ROW は全部の綴りを持っている。
    固定家賃 / 固定賃料
    店舗駐車場 / 店舗駐車場代
    店舗水道料金 / 店舗水道代
    共益費/ゴミ運搬処分費 / 共益費/ゴミ運搬処理費
    共益費/共通HP管理費 / 共益費/共通ＨＰ管理費（全角）
    日常清掃 / 日常清掃等 / 共益費/日常清掃等
    共益費/共有部の電気料金＋ガス料金＋水道料金 → 月によっては「共有部の水光熱費」1行
知らない摘要が金額つきで出てきたら止まる（ZERO_ONLY に無いもの）。
"""
import argparse
import collections
import glob
import os
import re
import subprocess

DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "yokocho_pdf")

# ファイル名の店舗 → PLタブ名
FILE2TAB = {"ハナ": "韓国酒場ハナ", "タコハイ": "タコとハイボール",
            "りゅうちゃん": "りゅうちゃん", "もも焼き": "もも焼きJAPAN"}

# 請求月（フォルダのYYMM）→ PL列。買掛フォルダの請求月＝PL列（他モジュールと同じ）
YM2MONTH = {"2509": "9月", "2510": "10月", "2511": "11月", "2512": "12月",
            "2601": "1月", "2602": "2月", "2603": "3月", "2604": "4月",
            "2605": "5月", "2606": "6月", "2607": "7月", "2608": "8月"}

# 摘要 → 店舗側のPL行
ITEM2ROW = {
    "固定家賃": "地代家賃（賃料）",
    "固定賃料": "地代家賃（賃料）",
    "歩合賃料": "地代家賃（賃料）",
    "店舗駐車場": "地代家賃（駐車場利用料）",
    "店舗駐車場代": "地代家賃（駐車場利用料）",
    "共通販促費": "広告宣伝費（共通宣伝費）",
    "店舗電気料金": "水道光熱費（電気料金）",
    "店舗水道料金": "水道光熱費（水道料金）",
    "店舗水道代": "水道光熱費（水道料金）",
    "備品購入費等（レジサーマル等）": "消耗品費",
    "備品購入費代（レジサーマル等）": "消耗品費",
    "備品購入費等": "消耗品費",
    "備品購入費代": "消耗品費",
    # ★スポンサー売掛相殺・その他値引きは【費用のマイナス】（利用者判断 2026-08-20
    #   「そのままでいい。費用のマイナスで」）。雑収入にはしない。
    #   既存21期PLにも前例がある（韓国酒場ハナ 5月 ▲2,727）。
    "スポンサー売掛相殺": "その他経費",
    "その他値引き": "その他経費",
}
# 「共益費/〜」と「日常清掃（等）」はまとめて 地代家賃（共益費） へ
KYOEKI = "地代家賃（共益費）"

# 金額が0でしか出てこない摘要。0なら無視、金額が付いたら止まる。
ZERO_ONLY = {"チャレンジショップ使用料", "テナント会雑費", "慶弔費", "備品購入費代"}

_LINE = re.compile(r"^(.+?)\s{2,}(-?[\d,]+)\s+件\s+(-?[\d,]+)\s+(-?[\d,]+)\s*$")
_NO = re.compile(r"No\.?\s+(\d{4})")
_SUM = re.compile(r"小計\s+(-?[\d,]+)")
_TAX = re.compile(r"消費税\s+(-?[\d,]+)\s*$")
_TOTAL = re.compile(r"合計\s+(-?[\d,]+)\s*$")


def _row_of(item):
    if item.startswith("共益費/") or item.startswith("日常清掃"):
        return KYOEKI
    return ITEM2ROW.get(item)


def _read(path):
    """PDF1枚 → {"items": {摘要: 金額}, "小計":…, "消費税":…, "税込":…, "No":…}"""
    txt = subprocess.run(["pdftotext", "-layout", path, "-"],
                         capture_output=True, text=True).stdout
    out = {"items": {}}
    for line in txt.splitlines():
        line = line.rstrip()
        m = _LINE.match(line)
        if m:
            item = m.group(1).strip()
            amount = int(m.group(4).replace(",", ""))
            if amount == 0:
                continue
            assert _row_of(item) or item in ZERO_ONLY, \
                f"{os.path.basename(path)}: 知らない摘要「{item}」に {amount:,}円 付いている"
            assert _row_of(item), \
                f"{os.path.basename(path)}: 「{item}」は0のはずなのに {amount:,}円 ある"
            out["items"][item] = out["items"].get(item, 0) + amount
            continue
        for key, pat in (("小計", _SUM), ("消費税", _TAX), ("税込", _TOTAL)):
            mm = pat.search(line)
            if mm and key not in out:
                out[key] = int(mm.group(1).replace(",", ""))
        mm = _NO.search(line)
        if mm and "No" not in out:
            out["No"] = int(mm.group(1))
    return out


def load():
    """{PL列: {タブ: {"摘要別": {...}, "小計":…, "消費税":…, "税込":…, "No":…,
                      "src": …}}}"""
    data = collections.defaultdict(dict)
    for path in sorted(glob.glob(os.path.join(DIR, "*.pdf"))):
        ym, store = os.path.basename(path)[:-4].split("_", 1)
        d = _read(path)
        # ★2511月（12月末請求）の3店舗だけ、相殺が【税込】で後引きされている。
        #   ふつうは  明細の和（相殺こみ）＝小計（税抜）  なのだが、この3枚は
        #       プラス明細の和 ×1.1 − 相殺 ＝ 合計（税込）
        #   という組み方で、請求書の小計は 合計÷1.1 を書いている。
        #   PLに入れるのは税抜なので、請求書の小計を正として、
        #   相殺の側を「小計 − プラス明細の和」に直す（税抜相当にそろえる）。
        #   例: 2511月もも焼き 相殺 ▲112,133（税込）→ ▲101,939（税抜相当）
        d["相殺の税抜補正"] = None
        got = sum(d["items"].values())
        if got != d["小計"]:
            plus = sum(v for v in d["items"].values() if v > 0)
            neg = {k: v for k, v in d["items"].items() if v < 0}
            assert neg, f"{os.path.basename(path)}: 明細の和 {got:,} ≠ 小計 {d['小計']:,}"
            want = d["小計"] - plus          # 税抜にそろえた相殺の合計（マイナス）
            raw = sum(neg.values())
            assert abs(plus * 11 / 10 + raw - d["税込"]) <= 1, \
                (f"{os.path.basename(path)}: 明細の和 {got:,} ≠ 小計 {d['小計']:,}。"
                 f"相殺を税込として後引きしても合計 {d['税込']:,} にならない")
            rest = want
            for i, k in enumerate(sorted(neg, key=lambda x: neg[x])):
                v = neg[k] * want // raw if i < len(neg) - 1 else rest
                d["items"][k] = v
                rest -= v
            d["相殺の税抜補正"] = f"請求書の相殺は税込。小計から逆算して税抜相当に直した" \
                                 f"（{raw:,} → {want:,}）"
            got = sum(d["items"].values())
        assert got == d["小計"], \
            f"{os.path.basename(path)}: 明細の合計 {got:,} ≠ 小計 {d['小計']:,}"
        # ★請求書の消費税は切り捨てではなく四捨五入。1円ずれる月がある
        #   （2509月のもも焼き 223,969×10% = 22,396.9 → 22,397）。
        #   PLに入れるのは税抜（小計）なので、ここは1円の幅を許す。
        assert abs(d["小計"] * 10 / 100 - d["消費税"]) < 1, \
            f"{os.path.basename(path)}: 小計の10% {d['小計']*10/100:,.1f} ≠ 消費税 {d['消費税']:,}"
        assert d["小計"] + d["消費税"] == d["税込"], \
            f"{os.path.basename(path)}: 小計＋消費税 ≠ 合計 {d['税込']:,}"
        d["src"] = f"買掛/21期/{ym}月/（横丁の請求書）{os.path.basename(path)}"
        data[YM2MONTH[ym]][FILE2TAB[store]] = d
    return dict(data)


def gen(out="yokocho_data.py"):
    """読み取り結果を Python のデータファイルに書き出す。"""
    data = load()
    months = [m for m in YM2MONTH.values() if m in data]
    with open(out, "w", encoding="utf-8") as f:
        f.write('#!/usr/bin/env python3\n# -*- coding: utf-8 -*-\n')
        f.write('"""神栖横丁の社内請求書（自動生成）。'
                '直接なおさないこと。\n\n'
                '    python3 yokocho_parse.py --gen\n\n'
                'で yokocho_pdf/*.pdf から作り直す。'
                '請求書の読み方は yokocho_parse.py の冒頭を参照。\n'
                f'21期で手元にあるのは {len(months)}か月ぶん'
                f'（{"・".join(months)}）×4店舗。\n'
                '10月分（2510月）と8月分（2608月）はDropboxに無い。\n"""\n')
        f.write("DATA = {\n")
        for m in months:
            f.write(f'    "{m}": {{\n')
            for tab, d in data[m].items():
                f.write(f'        "{tab}": {{\n')
                f.write('            "摘要別": {\n')
                for k, v in d["items"].items():
                    f.write(f'                "{k}": {v},\n')
                f.write('            },\n')
                for k in ("小計", "消費税", "税込", "No", "相殺の税抜補正", "src"):
                    v = d[k]
                    f.write(f'            "{k}": {v!r},\n')
                f.write("        },\n")
            f.write("        },\n")
        f.write("}\n")
    print(f"{out} を書き出した（{len(months)}か月 × "
          f"{sum(len(v) for v in data.values())//len(months)}店舗）")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--gen", action="store_true")
    a = ap.parse_args()
    if a.gen:
        gen()
    else:
        data = load()
        for m, stores in data.items():
            print(f"── {m}")
            for tab, d in stores.items():
                by = collections.Counter()
                for item, amt in d["items"].items():
                    by[_row_of(item)] += amt
                print(f"   {tab:<14} No.{d['No']}  小計{d['小計']:>9,}  "
                      f"税込{d['税込']:>9,}")
                for row, v in by.items():
                    print(f"        {row:<24}{v:>10,}")
        print("\n検算すべて通過")
