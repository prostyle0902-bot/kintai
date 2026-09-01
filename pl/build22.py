#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""22期（2026.9〜2027.8）の空シートを作る

21期テスト版とまったく同じ11タブ・同じ行構成・同じ数式で、データは空。
2026年9月から本番運用するための器。

    python3 build22.py

行を増やしたくなったら build2.py の EXTRA_COGS / EXTRA_SGA を直す。
21期と22期で同じ定義を共有しているので、両方に反映される。
"""
import build2

OUT = "損益計算書_22期.xlsx"

if __name__ == "__main__":
    wb = build2.new_wb(period="22期")
    wb.save(OUT)
    print(f"{OUT} を作成")
    for t in build2.TABS:
        print(f"  {t:<14} {len(build2.LAYOUTS[t]):>3}行")
    print(f"  タブ: {wb.sheetnames}")
