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
# (ファイル名パターン, PL列, 検算値, 元ファイルの表示名, 重複除去キー)
# ★エクスポートの粒度が月によって違う。重複除去のキーを間違えると行が消える。
#   7月ぶん（invoices_2026082*.csv）… **合計請求書**単位。1行=1合計請求書。請求書No列が無い
#   8月ぶん（8月売上invoices.csv）  … **請求書**単位。1合計請求書に複数行ぶら下がる
#   8月を合計請求書Noで束ねると90行→67行になり、業務課が 7,334,577→2,673,423 と
#   激減してしまう（2026-09-10 に気づいた）。だからキーはファイルごとに持つ。
# ★2026-09-10 追加: 8月ぶんは1本のエクスポート（8月売上invoices.csv・90件）で来た。
#   請求日は全件 2026-08 なので月で切る必要が無い。7月のような取り直しの和集合も不要。
#   検算値は7月のときだけ（既存PLスプシと突き合わせた値）。8月は既存PLを使わない方針
#   なので突き合わせ先が無い。
FILES = [
    ("invoices*.csv", "7月", {"鳥害対策課": 6217800},
     "freeeカード明細/21期/invoices*.csv（8/20版＋8/21版の和集合）", "合計請求書No"),
    ("8月売上invoices.csv", "8月", {},
     "freeeカード明細/21期/8月売上invoices.csv", "請求書No"),
]

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

# ★グループが空の請求は、どの部門か決められないので計上せず保留に出す。
#   board側で部門が付いていない＝データからは判別できない（推測で埋めない）。
NOGROUP = []          # rows() が埋める。(月, 請求書No, 請求日, 顧客名, 案件名, 税抜)


def _read(path):
    for enc in ("utf-8-sig", "cp932"):
        try:
            with open(path, encoding=enc) as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            continue
    raise RuntimeError(path)


def _union(pattern, key):
    """パターンに合うCSVを全部読んで、key の列で重複を落とす。
    後から読んだファイルで上書きする（新しいエクスポートを優先）。
    ★key はファイルごと（FILES の5番目）。粒度を取り違えると行が消える。"""
    merged = {}
    for path in sorted(glob.glob(os.path.join(DIR, pattern))):
        for r in _read(path):
            assert key in r, f"{path} に「{key}」列が無い"
            merged[r[key]] = r
    return list(merged.values())


def rows():
    """(タブ, PL行, 月, 税抜, 消費税, 件数, 元ファイル, 内訳) を列挙。"""
    NOGROUP.clear()
    for pattern, month, check, srclabel, key in FILES:
        det = _union(pattern, key)
        by = {}
        for r in det:
            g = r["グループ"].strip()
            ex = int(float(r["請求金額（JPY・税抜）"]))
            if g not in GROUP2PL:
                if not g:
                    NOGROUP.append((month, r["請求書No"], r["請求日"],
                                    r["顧客名"], r.get("案件名", ""), ex))
                continue
            tab, plrow = GROUP2PL[g]
            k = (tab, plrow)
            tax = int(float(r["消費税"]))
            by.setdefault(k, {"ex": 0, "tax": 0, "n": 0, "detail": []})
            by[k]["ex"] += ex; by[k]["tax"] += tax; by[k]["n"] += 1
            by[k]["detail"].append((r["請求日"], r["顧客名"], ex))
        for tab, v in check.items():
            got = by[(tab, "売上")]["ex"]
            assert got == v, f"{month} {tab}: board {got:,} ≠ 既存スプシ {v:,}"
        for (tab, plrow), v in by.items():
            yield (tab, plrow, month, v["ex"], v["tax"], v["n"],
                   srclabel, sorted(v["detail"], key=lambda d: -d[2]))


def hold_rows():
    """グループが空でどの部門か決められなかった請求。(月, タブ, 内容, 理由)"""
    out = []
    for month, no, date, cust, anken, ex in NOGROUP:
        out.append((month, "（未定）",
                    f"board請求 No.{no} {cust} {anken}（税抜{ex:,}円）",
                    "boardの「グループ」列が空で、業務課／鳥害対策課／飲食事業部の"
                    "どれか決められない。データからは判別できないので計上していない。"
                    f"請求日{date}。グループを埋めてエクスポートし直すか、"
                    "どの部門か教えてもらえれば入ります"))
    return out


if __name__ == "__main__":
    for tab, plrow, m, ex, tax, n, src, det in sorted(rows(), key=lambda r: (r[2], r[0])):
        print(f"{m:<4}{tab:<10}{plrow:<8}{n:>3}件  税抜{ex:>12,}  消費税{tax:>10,}")
    if NOGROUP:
        print(f"\n★グループ空で保留 {len(NOGROUP)}件")
        for month, no, date, cust, anken, ex in NOGROUP:
            print(f"  {month} No.{no} {date} {cust} {anken} 税抜{ex:,}")
