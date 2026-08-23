#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PLの各セルに「この数字は何か」のメモを付ける

--- なぜ要るのか ---------------------------------------------------------
利用者から「ハナの7月、その他経費-50818円ってなんだっけ？」と聞かれた。
明細ログを見れば分かるが、2,000行を探しに行くのは現実的でない。
利用者の提案（2026-08-23）「コメント欄を利用してはどうかな」に沿って、
数字が入っているセルそのものにメモを付ける。カーソルを乗せれば出る。

    Excel(.xlsx)      … セルのコメント（赤い三角）
    スプレッドシート   … セルのメモ（note）。push_sheets.py が送る
    ※スプレッドシートの「コメント」（スレッド）ではなく「メモ」を使う。
      メモはAPIで書けて、人が付けたコメントと混ざらない。

--- 中身 -----------------------------------------------------------------
明細ログ（このワークブックのシート）を (店舗, 転記先PL行, 転記先の月) で束ねて、
1セルぶんのメモを組み立てる。ログが唯一の元ネタなので二重管理にならない。

    韓国酒場ハナ／その他経費／7月
    ────────────────────
    ※PLに載せていないもの
    ・神栖横丁［社内請求］ スポンサー売掛相殺 -50,818（税抜）
      買掛/21期/2607月/（横丁の請求書）2607_ハナ.pdf
      PLには載せていない（利用者確認 2026-08-23）。スポンサーへの請求は…

税抜が空のログ行は「PLに載せていないもの」として最後にまとめる。
金額の合計がセルの値と合わないときはメモの先頭に★を出す（取りこぼしの検知）。
"""
import collections

# 1件あたりの備考の長さ。長いメモはスプレッドシート側で読みにくい
MAX_NOTE = 260
# 1セルのメモ全体の長さ。Googleのメモは長すぎると途中で切れる
MAX_TOTAL = 4000
# 1セルに並べる明細の件数。これを超えたら「ほか n件」にまとめる
MAX_ITEMS = 12

COL = {"日付": 0, "店舗": 1, "取引先": 2, "摘要": 3, "税込": 4, "税率": 5,
       "税抜": 6, "消費税": 7, "PL行": 8, "月": 9, "元ファイル": 10, "備考": 12}


def _trim(s, n):
    s = " ".join(str(s or "").split())
    return s if len(s) <= n else s[:n - 1] + "…"


def _entries(wb):
    """{(店舗, PL行, 月): [ログ1行ぶんの dict, ...]}"""
    out = collections.defaultdict(list)
    ws = wb["明細ログ"]
    for r in ws.iter_rows(min_row=3, values_only=True):
        tab, plrow, m = r[COL["店舗"]], r[COL["PL行"]], r[COL["月"]]
        if not (tab and plrow and m):
            continue
        out[(tab, plrow, m)].append({k: r[i] for k, i in COL.items()})
    return out


def _one(tab, plrow, m, items, cell_value):
    """1セルぶんのメモを組み立てる。"""
    posted = [x for x in items if isinstance(x["税抜"], (int, float))]
    unposted = [x for x in items if not isinstance(x["税抜"], (int, float))]
    total = sum(int(x["税抜"]) for x in posted)

    head = f"{tab}／{plrow}／{m}"
    lines = [head, "─" * 22]
    if cell_value is not None and int(cell_value) != total:
        lines.append(f"★このメモの明細 {total:,} と セルの値 {int(cell_value):,} が"
                     f"合っていない（差 {int(cell_value) - total:,}）")
        lines.append("")
    elif posted:
        lines.append(f"計 {total:,} 円（税抜）／{len(posted)}件")

    def block(x):
        who = _trim(x["取引先"], 40) or "（取引先なし）"
        kind = f"［{_trim(x['摘要'], 20)}］" if x["摘要"] else ""
        amt = f" {int(x['税抜']):,}" if isinstance(x["税抜"], (int, float)) else ""
        b = [f"・{who}{kind}{amt}"]
        if x["日付"]:
            b[0] += f"  {_trim(x['日付'], 16)}"
        if x["元ファイル"]:
            b.append(f"   {_trim(x['元ファイル'], 90)}")
        if x["備考"]:
            b.append(f"   {_trim(x['備考'], MAX_NOTE)}")
        return b

    shown = posted[:MAX_ITEMS]
    for x in shown:
        lines += block(x)
    if len(posted) > MAX_ITEMS:
        rest = posted[MAX_ITEMS:]
        lines.append(f"・ほか {len(rest)}件 計 {sum(int(y['税抜']) for y in rest):,}"
                     f"（明細ログを見てください）")
    if unposted:
        lines += ["", "※PLに載せていないもの"]
        for x in unposted[:MAX_ITEMS]:
            lines += block(x)
    text = "\n".join(lines)
    return text if len(text) <= MAX_TOTAL else text[:MAX_TOTAL - 20] + "\n…（以下略）"


def build(wb, author="自動転記"):
    """ワークブックの各PLセルにコメントを付ける。付けた数を返す。"""
    import build2
    from openpyxl.comments import Comment
    n = mismatch = 0
    for (tab, plrow, m), items in _entries(wb).items():
        if tab not in wb.sheetnames or tab not in build2.RIDX:
            continue
        r = build2.RIDX[tab].get(plrow)
        if r is None or m not in build2.MCOL:
            continue
        c = wb[tab][f"{build2.MCOL[m]}{r}"]
        if c.value is None and not any(
                not isinstance(x["税抜"], (int, float)) for x in items):
            continue                      # 空セルにメモだけ残さない
        text = _one(tab, plrow, m, items, c.value)
        if text.splitlines()[2].startswith("★"):
            mismatch += 1
        cm = Comment(text, author)
        # だいたいの吹き出しの大きさ。長いメモでも読めるように
        cm.width = 460
        cm.height = min(420, 22 * (text.count("\n") + 2))
        c.comment = cm
        n += 1
    return n, mismatch
