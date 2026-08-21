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

import csv, glob, os

DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cards")
# ★エクスポートを取り直すたびに絞り込みが変わることがある。1本を正としない。
#   2026-08-20版: No.7737〜8175（58件）。テナントの合計請求書8210〜8219が無い
#   2026-08-21版: No.7737〜8219（55件）。↑は入ったが、Seaplus 1,100,000 など
#                 業務課・鳥害対策課の13件が落ちていた
#   両方を請求書Noで和集合にすると68件になり、鳥害対策課7月が 6,217,800 と
#   既存PLに一致する（CHECK）。だから cards/invoices*.csv を全部読んで束ねる。
FILES = [("invoices*.csv", "7月")]          # (ファイル名パターン, PL列)

GROUP2PL = {
    "業務課":    ("業務課", "売上"),
    "鳥害対策課": ("鳥害対策課", "売上"),
    # ★2026-08-21 変更: 「その他売上」→「売上」
    #   8/20版は9件481,245で、スポンサー等の細かいものだけだった。だから
    #   その他売上に入れていた。8/21版でテナントの合計請求書10件が加わり、
    #   19件2,965,826（＝横丁の家賃収入そのもの）になったので「売上」に移す。
    #   既存シートの神栖横丁「売上」も9月〜6月は毎月193万〜252万の家賃収入。
    "飲食事業部": ("神栖横丁", "売上"),
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


def _union(pattern):
    """cards/invoices*.csv を全部読んで、合計請求書Noで重複を落とす。
    後から読んだファイルで上書きする（新しいエクスポートを優先）。"""
    merged = {}
    for path in sorted(glob.glob(os.path.join(DIR, pattern))):
        for r in _read(path):
            merged[r["合計請求書No"]] = r
    return list(merged.values())


def rows():
    """(タブ, PL行, 月, 税抜, 消費税, 件数, 元ファイル, 内訳) を列挙。"""
    for pattern, month in FILES:
        det = _union(pattern)
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
                   "freeeカード明細/21期/invoices*.csv（8/20版＋8/21版の和集合）", sorted(v["detail"], key=lambda d: -d[2]))
