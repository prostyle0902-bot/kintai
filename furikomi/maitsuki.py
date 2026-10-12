#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""先の月まで請求書が届いているものを、その月が来たら振込一覧へ入れる

--- 利用者指示 2026-10-12 ---------------------------------------------------
「神栖横丁の10月〜4月まで、タカギの請求書がすでに転記されてるが、これ、毎月入れるように変更」
→ 損益計算書も振込一覧も、その月が来るまで入れない。
  振込一覧の 22期 11月〜4月タブの F11/G11（タカギ 90,000）は 2026-10-12 に消した。
  損益計算書のほうは pl/inv22.py の MAITSUKI。

    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 maitsuki.py --dry-run   # 見るだけ
    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 maitsuki.py             # 入れる
    ... maitsuki.py --today 2026-11-01 --dry-run                     # 試すとき

定期チェック（Routine）が毎回呼ぶ。その月のタブに入っていなければ入れる。
入っていれば transcribe.py が「転記済み」と言うだけなので、何度呼んでもよい。
★書き込みは transcribe.py と同じ決まり（空のセルにだけ書く・上書きしない）。
"""
import datetime
import json
import os
import sys
import tempfile

import transcribe

_TAKAGI = "22期/{f}/横丁/{n}神栖横丁さまタカギPR請求書（{s}様） 横丁{d}月末.pdf"
# (期, タブの月, 店舗フォルダ, 振込先, 税込, セクション, 元ファイル)
SCHEDULE = [
    ("22期", m, "横丁", "タカギ", 90000, "月末",
     "/※請求書※/買掛/" + _TAKAGI.format(f=f, n=n, s=s, d=m % 12 + 1))
    for m, f, n, s in [
        (10, "2610月", 3, "炭焼きさぶちゃん"),
        (11, "2611月", 4, "韓国酒場ハナ"),
        (12, "2612月", 5, "藁と薪"),
        (1, "2701月", 6, "PONO"),
        (2, "2702月", 7, "もも焼きJAPAN"),
        (3, "2703月", 8, "大衆酒場りゅうちゃん"),
        (4, "2704月", 9, "タコとハイボール"),
    ]
]


def _year(period, m):
    start = {"21期": 2025, "22期": 2026}[period]
    return start if m >= 9 else start + 1


def due(today):
    return [e for e in SCHEDULE if (today.year, today.month) >= (_year(e[0], e[1]), e[1])]


def main():
    args = sys.argv[1:]
    dry = "--dry-run" in args
    today = datetime.date.today()
    if "--today" in args:
        today = datetime.date.fromisoformat(args[args.index("--today") + 1])
    got = due(today)
    later = [e for e in SCHEDULE if e not in got]
    print(f"{today}: 月が来ているもの {len(got)} 件 ／ これから {len(later)} 件"
          + (f"（次は {later[0][1]}月 {later[0][3]}）" if later else ""), file=sys.stderr)
    if not got:
        return
    entries = [{"period": p, "month": m, "store": st, "vendor": v, "amount": a,
                "section": sec, "source": src} for p, m, st, v, a, sec, src in got]
    fd, path = tempfile.mkstemp(suffix=".json")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(entries, f, ensure_ascii=False)
        sys.argv = ["transcribe.py", path] + (["--dry-run"] if dry else [])
        transcribe.main()
    finally:
        os.remove(path)


if __name__ == "__main__":
    main()
