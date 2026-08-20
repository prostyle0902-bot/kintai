#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""board（請求管理システム）の請求一覧CSV → 業務課・鳥害対策課・神栖横丁の売上

店舗（りゅうちゃん等）の売上はPOSなのでここには来ない。
boardに載るのは請求書を発行する側＝売掛だけ。

読み方（2026-08-20 に7月分で検証）:
    「請求日」で月を切る
    「グループ」列で部門を分ける  ← これが決め手だった
    「請求金額（JPY・税抜）」を使う

検証: 鳥害対策課 6,217,800 が既存PLスプシの7月と1円まで一致。

グループ → PLタブ・行
    業務課      → 業務課「売上」
    鳥害対策課   → 鳥害対策課「売上」
    飲食事業部   → 神栖横丁「その他売上」

★神栖横丁について（利用者確認済 2026-08-20）
    既存スプシの7月には キッチンカー25,000／スポンサー（横丁）407,000／
    ビアガーデン800,000／横丁加盟金400,000 が入っているが、ビアガーデンの請求は
    6月計上なので7月ではない。boardを基準とするため、これら4行の7月値は使わない。
    boardの9件（481,245）を「その他売上」にまとめて入れ、内訳は明細ログに残す。
    ※どの請求がスポンサー／キッチンカー／加盟金に当たるかはboardのデータからは
      判別できないため、行レベルの割り当てはしていない。売上合計(1)は変わらない。
"""

import csv, os

DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cards")
FILES = [("invoices.csv", "7月")]          # (ファイル, PL列)

GROUP2PL = {
    "業務課":    ("業務課", "売上"),
    "鳥害対策課": ("鳥害対策課", "売上"),
    "飲食事業部": ("神栖横丁", "その他売上"),
}

# 既存スプシの売上のうち、boardを正として置き換える（＝転記しない）もの
SUPPRESS = [("神栖横丁", "キッチンカー", "7月"),
            ("神栖横丁", "スポンサー（横丁）", "7月"),
            ("神栖横丁", "ビアガーデン", "7月"),
            ("神栖横丁", "横丁加盟金", "7月")]

# 検算値（既存PLスプシの7月）
CHECK = {"鳥害対策課": 6217800}


def _read(path):
    for enc in ("utf-8-sig", "cp932"):
        try:
            with open(path, encoding=enc) as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            continue
    raise RuntimeError(path)


def rows():
    """(タブ, PL行, 月, 税抜, 消費税, 件数, 元ファイル, 内訳) を列挙。"""
    for fname, month in FILES:
        det = _read(os.path.join(DIR, fname))
        by = {}
        for r in det:
            g = r["グループ"]
            if g not in GROUP2PL:
                continue
            tab, plrow = GROUP2PL[g]
            k = (tab, plrow)
            ex = int(float(r["請求金額（JPY・税抜）"]))
            tax = int(float(r["消費税"]))
            by.setdefault(k, {"ex": 0, "tax": 0, "n": 0, "detail": []})
            by[k]["ex"] += ex; by[k]["tax"] += tax; by[k]["n"] += 1
            by[k]["detail"].append((r["請求日"], r["顧客名"], ex))
        for tab, v in CHECK.items():
            got = by[(tab, "売上")]["ex"]
            assert got == v, f"{tab}: board {got:,} ≠ 既存スプシ {v:,}"
        for (tab, plrow), v in by.items():
            yield (tab, plrow, month, v["ex"], v["tax"], v["n"],
                   f"freeeカード明細/21期/{fname}", sorted(v["detail"], key=lambda d: -d[2]))
