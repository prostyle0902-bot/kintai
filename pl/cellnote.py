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
import os
import re

# 1件あたりの備考の長さ。長いメモはスプレッドシート側で読みにくい
MAX_NOTE = 260
# 1セルのメモ全体の長さ。Googleのメモは長すぎると途中で切れる
MAX_TOTAL = 4000
# 1セルに並べる明細の件数。これを超えたら「ほか n件」にまとめる
MAX_ITEMS = 12

COL = {"日付": 0, "店舗": 1, "取引先": 2, "摘要": 3, "税込": 4, "税率": 5,
       "税抜": 6, "消費税": 7, "PL行": 8, "月": 9, "元ファイル": 10, "備考": 12}


# ---- 出どころの言い換え（利用者要望 2026-08-23「ソースがわかるとありがたい」）----
# ログの「元ファイル名」はモジュールごとに書き方がバラバラなので、
# 「どこを開けば実物にたどり着けるか」が分かる形にそろえる。
DROPBOX = "Dropbox "                     # ※請求書※ など Dropbox の共有フォルダ
REPO = "このリポジトリ pl/"               # 手元に置いてあるCSV

# 千葉銀行の口座番号 → 何の口座か（銀行明細のファイル名だけでは分からない）
ACCOUNT = {
    "3351509": "本体（ビルメンのメイン口座）", "3543920": "神栖横丁",
    "3543939": "焼きたて屋", "3546725": "さわら十三里屋",
    "3548060": "タコとハイボール", "3555848": "もも焼きJAPAN",
    "3556801": "りゅうちゃん", "3563077": "韓国酒場ハナ",
}
# 既存21期PLのスプレッドシート（タブごとに別ファイル）
EXIST_SHEET = {
    "りゅうちゃん": "1N5QF8UO_", "もも焼きJAPAN": "16a9y8Ex", "韓国酒場ハナ": "1MbrqksG",
    "さわら十三里屋": "1K1U0hql", "タコとハイボール": "10PZ-wKL", "焼きたて屋": "1OB3yLta",
    "神栖横丁": "1aI4rq96", "鳥害対策課": "1q8C6Syu", "業務課": "1JKy9P6_", "本部": "1xO2MPtc",
}

_BANK = re.compile(r"小見川支店_普通_(\d{7})_(\d{4})(\d{2})")


def source(src):
    """ログの「元ファイル名」→ 実物にたどり着ける書き方。"""
    s = " ".join(str(src or "").split())
    if not s:
        return ""
    if s.startswith("/"):                       # すでにDropboxのフルパス
        return DROPBOX + s
    m = _BANK.search(s)
    if m:
        acct, y, mo = m.groups()
        who = ACCOUNT.get(acct, "")
        return (f"千葉銀行 小見川支店 普通 {acct}"
                + (f"（{who}）" if who else "")
                + f" {y}年{int(mo)}月の明細CSV ／ {REPO}bank/{os.path.basename(s)}")
    if s.startswith("NBG"):
        return f"PayPay銀行の明細CSV ／ {REPO}bank/{s}"
    if s.startswith("銀行明細/"):
        return ("千葉銀行の店舗口座＋PayPay銀行の明細CSV ／ "
                f"{REPO}bank/（{s.split('/')[-1]}）")
    if s.startswith("bank/"):
        return f"{REPO}{s}"
    if s.startswith("買掛/"):
        return DROPBOX + "/※請求書※/" + s
    if s.startswith("freeeカード明細/"):
        return DROPBOX + "/※請求書※/" + s
    if s.startswith("statement-"):
        return (f"freeeの経費精算CSV ／ {REPO}csv/{s}"
                "（Dropbox /※請求書※/freeeカード明細/21期/ と同じもの）")
    if s.startswith("既存21期PL"):
        tab = s.replace("既存21期PL", "").strip()
        u = EXIST_SHEET.get(tab)
        return ("会計士さんの既存21期PLスプレッドシート"
                + (f"「{tab}」タブ" if tab else "")
                + (f"（ID {u}…）" if u else ""))
    if s.startswith("既存PLスプシ"):
        return "会計士さんの既存21期PLスプレッドシート（毎月同額の自動引落・自動振込）"
    if s.startswith("かめや"):
        return DROPBOX + "/※請求書※/" + s.replace("かめや（焼きたて屋本部）", "かめや")
    if s.startswith("Drive "):
        return "Google Drive " + s[len("Drive "):]
    if s.startswith("給与集計"):
        return DROPBOX + "/※プロスタイル給与※/ " + s.split("（")[0]
    return s


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
            b.append(f"   出どころ: {_trim(source(x['元ファイル']), 150)}")
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
