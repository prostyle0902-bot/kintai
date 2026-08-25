#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""既に入っている金額を訂正する（transcribe.py は上書きを拒否するため別建て）

    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 correct.py corrections.json --dry-run
    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 correct.py corrections.json

corrections.json の形式（配列）:
    [{"period": "21期", "month": 7, "store": "ハナ", "vendor": "日本食研",
      "from": 31515,        # いま入っている値。一致しなければ中止する
      "to": 34036,
      "reason": "税抜→税込"}]

**必ず `from` を書くこと。** 実際の値と違ったら1件も書き込まずに止まる。
別のセッションが先に直していた／行がずれていた、を事故にしないため。

このスクリプトは利用者が明示的に訂正を指示したときだけ使う。
自動転記（transcribe.py）からは呼ばない。
"""
import json
import sys

import transcribe as t


def main():
    args = [a for a in sys.argv[1:]]
    dry = "--dry-run" in args
    args = [a for a in args if a != "--dry-run"]
    if not args:
        sys.exit(__doc__)

    items = json.load(open(args[0], encoding="utf-8"))
    api = t.svc()

    tabs = {}
    plans = []
    errors = []
    for it in items:
        key = (it["period"], int(it["month"]))
        if key not in tabs:
            tabs[key] = t.Tab(api, t.SHEETS[it["period"]], int(it["month"]))
        tab = tabs[key]

        col_name = t.STORE_TO_COLUMN.get(it["store"], it["store"])
        if col_name not in tab.store_cols:
            errors.append(f"{it['store']}/{it['vendor']}: 列『{col_name}』が無い")
            continue
        vcol, acol = tab.store_cols[col_name]

        found = []
        for k in tab.sections:
            for r in tab.section_rows(k):
                if tab.is_total_row(r):
                    continue
                if t.vendor_match(tab.cell(r, vcol), it["vendor"]):
                    found.append(r)
        if len(found) != 1:
            errors.append(
                f"{it['store']}/{it['vendor']}: 行が{len(found)}件見つかった"
                f"（1件でないと訂正しない）: "
                + ", ".join(f"行{r+1}" for r in found))
            continue

        r = found[0]
        cur = t.amount_of(tab.cell(r, acol))
        if cur != int(it["from"]):
            errors.append(
                f"{it['store']}/{it['vendor']} {t.col_letter(acol)}{r+1}: "
                f"いまの値 {cur} が from={it['from']} と違う。中止")
            continue

        plans.append((tab, r, acol, int(it["to"]), it))

    for tab, r, c, to, it in plans:
        print(f"  {it['store']:8} {it['vendor']:14} "
              f"{t.col_letter(c)}{r+1}: {it['from']:,} → {to:,}   {it.get('reason','')}")
    for e in errors:
        print("  NG:", e, file=sys.stderr)

    if errors:
        sys.exit("\nエラーがあるので1件も書き込まずに終了した。")
    if dry:
        print("\n（dry-run: 書き込みなし）")
        return

    by_sheet = {}
    for tab, r, c, to, it in plans:
        by_sheet.setdefault(it["period"], []).append(
            {"range": f"'{tab.title}'!{t.col_letter(c)}{r + 1}", "values": [[to]]})
    for period, data in by_sheet.items():
        api.spreadsheets().values().batchUpdate(
            spreadsheetId=t.SHEETS[period],
            body={"valueInputOption": "USER_ENTERED", "data": data},
        ).execute()
    print(f"\n{len(plans)}件を訂正した。")


if __name__ == "__main__":
    main()
