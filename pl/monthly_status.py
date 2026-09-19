#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dropboxの「★毎月ここに入れる」に置く受け入れ状況メモを作る

    python3 monthly_status.py            # 本文を標準出力へ
    python3 monthly_status.py --out x.txt

定期チェック（Routine）から呼び、出てきた本文を Claude が Dropbox の
    /※請求書※/★毎月ここに入れる/今月の受け入れ状況.txt
へ書く（create_file は同名があれば上書きになる）。

--- 何を見ているか -------------------------------------------------------
pl/ の中に元データのファイルが実際に置いてあるかを見るだけ。
Dropbox に入れてもらったものは定期チェックで pl/ に落とすので、
「pl/ にある＝取り込み済み」でよい。

まだ足りないものは status8.EMPTY8_WAITING と kessai.hold_rows() から拾う。
どちらも「8月が空いている理由」を持っている表。
"""
import argparse
import datetime
import glob
import os

BASE = os.path.dirname(os.path.abspath(__file__))

# 21期＝2025年9月〜2026年8月。いまの対象月（PL列）とそのファイル年月
TARGET = "8月"
YM = {
    "銀行":       ("bank/小見川支店_普通_*_202608.csv", "202608", 9),
    "PayPay銀行": ("bank/NBG_2026b.csv",                "2026年8月", 1),
    "freeeカード": ("csv/statement-2026-09.csv",         "2026-09（＝8月利用分）", 1),
    "JCB":        ("cards/202609meisai.csv",            "202609（＝8月ぶん）", 1),
    "三井住友":    ("cards/202609.csv",                  "202609（＝8月ぶん）", 1),
    "board売上":   ("cards/8月売上*.csv",                "8月売上", 2),
    "エアレジ売上": ("uriage/2608月/*.csv",              "2608月", 5),
    "給料一覧表":  ("kyuyo/202608.pdf",                  "202608", 1),
}

# ★ファイルではなくモジュールの中に数字を焼いてあるもの。
#   （Claudeが Dropbox / ドライブ の書類を読んで書き込む作り）
BAKED = {
    "かめや":      ("kameya",  "DATA",      "2608月の3点セット"),
    "横丁の請求書": ("yokocho_data", "DATA", "2608月の合計請求書4枚"),
}


def rows():
    """(区分, 期待するもの, 何件あるか, 足りているか)"""
    for kind, (pattern, want, need) in YM.items():
        hit = glob.glob(os.path.join(BASE, pattern))
        yield kind, want, len(hit), len(hit) >= need
    for kind, (mod, attr, want) in BAKED.items():
        try:
            data = getattr(__import__(mod), attr)
            have = TARGET in data
        except Exception:
            have = False
        yield kind, want, 1 if have else 0, have


def waiting():
    """まだ届いていないもの。(どこ, 何)"""
    import status8
    for tab, plrow, v7, why in status8.EMPTY8_WAITING:
        yield f"{tab} {plrow}", why
    import kessai
    for m, tab, what, why in kessai.hold_rows():
        yield f"{tab} {what}", why


def text():
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    out = []
    out.append("■ 今月の受け入れ状況（21期 8月ぶん）")
    out.append(f"   {now} 時点。自動チェックのたびに書き換えます")
    out.append("")
    ok = [r for r in rows() if r[3]]
    ng = [r for r in rows() if not r[3]]
    out.append(f"● 届いているもの（{len(ok)}件）")
    for kind, want, n, _ in ok:
        out.append(f"    ✓ {kind}　{want}　{n}件")
    if ng:
        out.append("")
        out.append(f"● ★まだ届いていないもの（{len(ng)}件）")
        for kind, want, n, _ in ng:
            out.append(f"    × {kind}　{want} を入れてください")
    w = list(waiting())
    if w:
        out.append("")
        out.append(f"● 1件ずつ待っているもの（{len(w)}件）")
        for what, why in w:
            out.append(f"    ・{what}")
            out.append(f"        {why}")
    out.append("")
    out.append("─" * 30)
    out.append("このフォルダが空＝入れてもらったものは全部取り込み済みです。")
    out.append("ファイルが残っているときは、まだ取り込めていないのでご連絡します。")
    return "\n".join(out) + "\n"


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out")
    a = ap.parse_args()
    t = text()
    if a.out:
        open(a.out, "w", encoding="utf-8").write(t)
        print(f"{a.out} に書き出した")
    else:
        print(t)
