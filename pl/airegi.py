#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Airレジの集計CSV → 各店の月別売上（税込）

Dropbox `※請求書※/会計明細/` に置かれた Airレジの集計CSVを読んで、
`sales.py` の 売上 系列（9月〜8月の12個）を作る。

    python3 airegi.py            # 取り込んである全店を sales.py と突き合わせる
    python3 airegi.py --emit     # sales.py に貼る形で出力する

## CSVの置き方（利用者の運用・2026-09-01）

- **21期は年間で1本**。`会計明細/21期/<店舗>_YYYYMM-YYYYMM.csv`
  → 期の直下。1行=1か月。遡って埋めるため。
- **22期からは月ごと**。`会計明細/22期/<YYMM月>/<店舗>_YYYYMMDD-YYYYMMDD.csv`
  → 月フォルダの中。1行=1日。

どちらも1列目が `集計期間`（年間なら `YYYYMM`、月次なら `YYYYMMDD`）で、
2列目が `売上`。この2列だけ見れば月合計は出せる。月次CSVは日別を合算する。

**文字コードは cp932（Shift-JIS）。** utf-8 で開くと落ちる。

## 消費税は入っていない

CSVにあるのは 売上 とその税率別内訳（`売上（10%標準）` `売上（8%軽減）` …）だけで、
**消費税額の列は無い**。月合計から 10/110 で逆算しても既存PLの値には合わない。
りゅうちゃんの21期9月〜7月で比べると毎月100〜240円ほど逆算が多く、その差は
会計数にきれいに比例する（差 ÷ 会計数 ≒ 0.47円）。POSが会計1件ごとに切り捨てて
いるため。README「既存PLの消費税＝かめや合計精算書のPOS売上明細表の日別消費税の
合計」とも整合する。

→ **売上はCSVで入れ替えてよいが、消費税はCSVからは作れない。** 既存PLの値を残すか、
   POSの精算書から取る。逆算で埋めない（README「推測で埋めない」）。
"""
import csv
import glob
import os
import re
import sys

import sales

# CSVのファイル名の店舗名 → PLのタブ名
STORE_TO_TAB = {
    "りゅうちゃん": "りゅうちゃん",
    "もも焼き": "もも焼きJAPAN",
    "ハナ": "韓国酒場ハナ",
    "十三里屋": "さわら十三里屋",
    "タコハイ": "タコとハイボール",
    "焼きたて屋": "焼きたて屋",
}

# PLの列の並び。21期なら 202509〜202608
MONTHS = ["9月", "10月", "11月", "12月", "1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月"]
PERIOD_START = {"21期": 202509, "22期": 202609}

# タブごとの「売上」行の名前（sales.SALES_ROWS の1本目）
def uriage_row(tab):
    return sales.SALES_ROWS[tab][0][0]


def ym_list(period):
    """その期の 202509,202510,... を9月〜8月の順で12個"""
    y, m = divmod(PERIOD_START[period] - 1, 100)
    out = []
    for _ in range(12):
        m += 1
        if m > 12:
            y, m = y + 1, 1
        out.append(y * 100 + m)
    return out


def read_csv(path):
    """集計期間 → 売上 の辞書。月次CSV（日別）は YYYYMM に丸めて合算する"""
    with open(path, encoding="cp932", newline="") as f:
        rows = list(csv.reader(f))
    head = rows[0]
    i_ki, i_uri = head.index("集計期間"), head.index("売上")
    out = {}
    for r in rows[1:]:
        if not r or not r[i_ki].strip():
            continue
        ki = r[i_ki].strip()
        ym = int(ki[:6])          # YYYYMMDD でも YYYYMM でも先頭6桁が年月
        out[ym] = out.get(ym, 0) + int(r[i_uri] or 0)
    return out


def collect(period="21期"):
    """{タブ名: {YYYYMM: 売上}}。年間CSVと月次CSVの両方を拾う"""
    base = os.path.join(os.path.dirname(os.path.abspath(__file__)), "airegi")
    found = {}
    for path in sorted(glob.glob(os.path.join(base, "**", "*.csv"), recursive=True)):
        name = os.path.basename(path)
        m = re.match(r"(.+?)_\d{6}", name)
        if not m:
            print(f"  ?? ファイル名が読めない: {name}", file=sys.stderr)
            continue
        store = m.group(1)
        tab = STORE_TO_TAB.get(store)
        if tab is None:
            print(f"  ?? 店舗名が対応表に無い: {store} ({name})", file=sys.stderr)
            continue
        for ym, v in read_csv(path).items():
            found.setdefault(tab, {})[ym] = v
    return found


def series(period, tab, found):
    """9月〜8月の12個。CSVに無い月は None"""
    return [found.get(tab, {}).get(ym) for ym in ym_list(period)]


def compare(period="21期"):
    """CSVと sales.py を突き合わせて差を出す。戻り値は差のあった件数"""
    found = collect(period)
    if not found:
        print("airegi/ にCSVがまだ無い")
        return 0
    ngs = 0
    for tab in sales.SALES_ROWS:
        if tab not in found:
            continue
        row = uriage_row(tab)
        cur = sales.SALES.get(tab, {}).get(row, [None] * 12)
        new = series(period, tab, found)
        print(f"\n=== {tab}（{row}）")
        print(f"{'月':>4} {'CSV':>11} {'sales.py':>11} {'差':>10}")
        for m, a, b in zip(MONTHS, new, cur):
            if a is None:
                print(f"{m:>4} {'-':>11} {fmt(b):>11}   CSVに無い")
                continue
            d = a - b if b is not None else None
            mark = "" if d == 0 else ("  ★新規" if b is None else "  ★差あり")
            if d != 0:
                ngs += 1
            print(f"{m:>4} {a:>11,} {fmt(b):>11} {fmt(d):>10}{mark}")
    return ngs


def fmt(v):
    return "-" if v is None else f"{v:,}"


def emit(period="21期"):
    """sales.py の SALES に貼れる形で出す"""
    found = collect(period)
    for tab in sales.SALES_ROWS:
        if tab not in found:
            continue
        new = series(period, tab, found)
        vals = ",".join("N" if v is None else str(v) for v in new)
        print(f' "{tab}": {{\n   "{uriage_row(tab)}": [{vals}],\n }},')


if __name__ == "__main__":
    if "--emit" in sys.argv:
        emit()
    else:
        n = compare()
        print(f"\n差のあったセル: {n}")
