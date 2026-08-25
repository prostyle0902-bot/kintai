#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""請求書 → 振込一覧スプレッドシートへの転記スクリプト

Dropbox の ※請求書※/買掛/<期>/<YYMM月>/<店舗>/ に入った請求書から
抜き出した「取引先・税込金額」を、Google スプレッドシート
「21期_振込一覧」「22期_振込一覧」の該当月タブへ書き込む。

    21期: 1CF5lIjrs3N4fTrVHa6BzSZtvbb0fmMWGZC3m8zn7LuM
    22期: 1d4s7XqlTb9xK-o3ujDd8PGiDMo23l5TzSg_wpl3DzRU

使い方:
    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 transcribe.py entries.json
    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 transcribe.py entries.json --dry-run
    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 transcribe.py --layout 21期 7月

entries.json の形式（配列）:
    [{"period": "21期",            # 21期 / 22期
      "month": 7,                  # フォルダの月 = タブの月（同月。検証済み）
      "store": "タコハイ",          # Dropboxの店舗フォルダ名
      "vendor": "ヴィーナスダイニング",  # シートの振込先表記に合わせた短い名前
      "amount": 212703,            # 今回御請求額（税込・整数円）
      "section": "月末",           # 月末 / 10日 / 小見川 / 自動 / null(=既存行に従う)
      "source": "/※請求書※/買掛/21期/2607月/タコハイ/ヴィーナスダイニング　タコハイ8月末.pdf"}]

書き込みルール（README.md も参照）:
  - 振込先が既にその月タブ・その店舗列にあり金額が空 → 金額だけ入れる
  - 同じ振込先・同じ金額が既に入っている → 何もしない（転記済み）
  - 振込先があるが違う金額が入っている → 空きの同名行を探し、無ければ保留
  - 振込先が無い → 指定セクションの空き行に 振込先+金額 を新規記入
  - 既に入っている値は絶対に上書きしない
"""
import json
import os
import re
import sys
import unicodedata

SHEETS = {
    "21期": "1CF5lIjrs3N4fTrVHa6BzSZtvbb0fmMWGZC3m8zn7LuM",
    "22期": "1d4s7XqlTb9xK-o3ujDd8PGiDMo23l5TzSg_wpl3DzRU",
}

# Dropboxの店舗フォルダ名 → シートの列見出し
# 本部フォルダの請求書は業務課の列に入れる（既存運用: ピカピカ・CCMO・佐倉クレーン等）
STORE_TO_COLUMN = {
    "業務": "業務課",
    "本部": "業務課",
    "鳥害対策": "鳥害対策課",
    "横丁": "神栖横丁",
    "もも焼き": "もも焼きJAPAN",
    "タコハイ": "タコハイ",
    "りゅうちゃん": "りゅうちゃん",
    "ハナ": "ハナ",
    "焼きたて屋": "焼きたて屋",
    "十三里屋": "十三里屋",
}

SECTION_PATTERNS = {
    "月末": re.compile(r"月末に.*支払う分"),
    "10日": re.compile(r"10日?に.*振り込む分|/10"),
    "小見川": re.compile(r"小見川"),
    "自動": re.compile(r"自動引落"),
}


def norm(s):
    """照合用の正規化: 全角→半角、空白・記号ゆらぎを吸収"""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = re.sub(r"[\s　]+", "", s)
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"^(株式会社|有限会社|\(株\)|\(有\)|カ\)|ユ\))", "", s)
    s = re.sub(r"(株式会社|有限会社|\(株\)|\(有\))$", "", s)
    return s.lower()


def vendor_match(sheet_name, entry_vendor):
    a, b = norm(sheet_name), norm(entry_vendor)
    if not a or not b:
        return False
    return a == b or a in b or b in a


def creds():
    from google.oauth2 import service_account
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    raw = os.environ.get("GOOGLE_SA_KEY")
    if raw:
        return service_account.Credentials.from_service_account_info(
            json.loads(raw), scopes=scopes)
    path = os.environ.get("GOOGLE_SA_KEY_FILE")
    if path and os.path.exists(path):
        return service_account.Credentials.from_service_account_file(path, scopes=scopes)
    sys.exit("鍵がない: GOOGLE_SA_KEY か GOOGLE_SA_KEY_FILE を設定すること。\n"
             "実物は Dropbox /※請求書※/sa-key.json.json（pl/push.py の冒頭参照）")


def svc():
    from googleapiclient.discovery import build
    return build("sheets", "v4", credentials=creds(), cache_discovery=False)


def col_letter(idx0):
    s = ""
    n = idx0 + 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class Tab:
    """月タブ1枚の構造: 店舗列・セクション行範囲・全セル値"""

    def __init__(self, api, spreadsheet_id, month):
        self.api = api
        self.sid = spreadsheet_id
        meta = api.spreadsheets().get(
            spreadsheetId=spreadsheet_id,
            fields="sheets(properties(sheetId,title,index))").execute()
        self.all_tabs = [s["properties"] for s in meta["sheets"]]
        want = None
        pat = re.compile(rf"^{month}月")
        for p in self.all_tabs:
            if pat.match(p["title"].strip()):
                want = p
                break
        if want is None:
            raise LookupError(
                f"{month}月 のタブが見つからない。タブ一覧: "
                + ", ".join(f"{p['title']}(gid={p['sheetId']})" for p in self.all_tabs))
        self.title = want["title"]
        self.gid = want["sheetId"]
        got = api.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id,
            range=f"'{self.title}'!A1:V300",
            valueRenderOption="UNFORMATTED_VALUE").execute()
        self.rows = got.get("values", [])
        self._parse()

    def cell(self, r, c):
        if r < len(self.rows) and c < len(self.rows[r]):
            v = self.rows[r][c]
            return "" if v is None else v
        return ""

    def _parse(self):
        # 店舗列: 「業務課」「鳥害対策課」… の見出し行を探す
        # （結合セルなので値は左上セルにだけ入る。見出しの次行が 振込先/金額）
        self.store_cols = {}   # 列見出し → (振込先col, 金額col)
        header_row = None
        for r in range(min(8, len(self.rows))):
            hits = 0
            for c in range(len(self.rows[r])):
                if str(self.cell(r, c)).strip() in STORE_TO_COLUMN.values():
                    hits += 1
            if hits >= 5:
                header_row = r
                break
        if header_row is None:
            raise LookupError(f"タブ『{self.title}』で店舗見出し行が見つからない")
        for c in range(len(self.rows[header_row])):
            name = str(self.cell(header_row, c)).strip()
            if name in STORE_TO_COLUMN.values() and name not in self.store_cols:
                self.store_cols[name] = (c, c + 1)
        self.header_row = header_row

        # セクション: A列のラベル行から次のラベルまで。
        # 結合セルのため値は先頭行のみ。「合計」行・空行は書き込み対象外。
        self.sections = {}     # キー(月末/10日/小見川/自動) → (開始行, 終了行) 0-based
        marks = []
        for r in range(header_row, len(self.rows)):
            a = str(self.cell(r, 0)).strip()
            if not a:
                continue
            for key, pat in SECTION_PATTERNS.items():
                if pat.search(a):
                    marks.append((r, key))
                    break
        for i, (r, key) in enumerate(marks):
            end = marks[i + 1][0] - 1 if i + 1 < len(marks) else len(self.rows) - 1
            # セクション末尾の合計行（金額列に数式合計が入る行）は除く:
            # 最後の2行程度に「合計」文字 or 全店舗の金額が数値で埋まる行がある
            self.sections.setdefault(key, (r, end))

    def describe(self):
        out = [f"タブ『{self.title}』 gid={self.gid}"]
        out.append("店舗列: " + ", ".join(
            f"{k}={col_letter(v[0])}/{col_letter(v[1])}" for k, v in self.store_cols.items()))
        for key, (s, e) in self.sections.items():
            out.append(f"セクション {key}: 行{s + 1}〜{e + 1}")
        return "\n".join(out)

    def section_rows(self, key):
        if key not in self.sections:
            raise LookupError(f"タブ『{self.title}』にセクション {key} が無い")
        s, e = self.sections[key]
        return range(s, e + 1)

    def is_total_row(self, r):
        """「合計」マーカーがある行・総計行は避ける"""
        for c in range(0, 22):
            if str(self.cell(r, c)).strip() == "合計":
                return True
        return False


def amount_of(v):
    if v in ("", None):
        return None
    if isinstance(v, (int, float)):
        return int(round(v))
    s = re.sub(r"[¥,\s円]", "", str(v))
    try:
        return int(s)
    except ValueError:
        return None


def plan_entry(tab, entry):
    """書き込み先セルを決める。返り値: (action, updates, note)
    action: write / already / hold
    updates: [(row0, col0, value)]"""
    col_name = STORE_TO_COLUMN.get(entry["store"])
    if not col_name:
        return ("hold", [], f"店舗フォルダ名が未知: {entry['store']}")
    if col_name not in tab.store_cols:
        return ("hold", [], f"タブ『{tab.title}』に列『{col_name}』が無い")
    vcol, acol = tab.store_cols[col_name]
    vendor, amount = entry["vendor"], int(entry["amount"])

    sec = entry.get("section")
    search_secs = [sec] if sec in tab.sections else list(tab.sections)

    empty_named = None      # 振込先一致・金額空
    same_value = None       # 振込先一致・同額入り
    conflict = None         # 振込先一致・別額入り
    for key in search_secs:
        for r in tab.section_rows(key):
            if tab.is_total_row(r):
                continue
            name = tab.cell(r, vcol)
            if not vendor_match(name, vendor):
                continue
            cur = amount_of(tab.cell(r, acol))
            if cur is None and empty_named is None:
                empty_named = r
            elif cur == amount and same_value is None:
                same_value = r
            elif cur is not None and cur != amount:
                conflict = (r, cur)

    if same_value is not None:
        return ("already", [], f"行{same_value + 1}に同額 {amount:,} が入力済み")
    if empty_named is not None:
        return ("write", [(empty_named, acol, amount)],
                f"既存の振込先行 行{empty_named + 1} に金額を記入")
    if conflict is not None and sec is None:
        return ("hold", [],
                f"行{conflict[0] + 1}に別額 {conflict[1]:,} が入力済み。"
                f"別請求なら section を指定して再実行")

    # 新規行: 指定セクション（無指定なら月末）の空きスロットへ
    key = sec if sec in tab.sections else "月末"
    for r in tab.section_rows(key):
        if tab.is_total_row(r):
            continue
        if str(tab.cell(r, vcol)).strip() == "" and tab.cell(r, acol) in ("", None):
            return ("write", [(r, vcol, vendor), (r, acol, amount)],
                    f"セクション{key}の空き行{r + 1}に新規記入")
    return ("hold", [], f"セクション{key}の{col_name}列に空き行が無い")


def main():
    args = [a for a in sys.argv[1:]]
    dry = "--dry-run" in args
    args = [a for a in args if a != "--dry-run"]

    if args and args[0] == "--layout":
        period, month = args[1], int(re.sub(r"\D", "", args[2]))
        tab = Tab(svc(), SHEETS[period], month)
        print(tab.describe())
        return

    if not args:
        sys.exit(__doc__)
    entries = json.load(open(args[0], encoding="utf-8"))
    api = svc()

    tabs = {}
    results = []
    writes_by_tab = {}
    for e in entries:
        key = (e["period"], int(e["month"]))
        if key not in tabs:
            try:
                tabs[key] = Tab(api, SHEETS[e["period"]], int(e["month"]))
            except Exception as ex:
                results.append({**e, "action": "hold", "note": f"タブ取得失敗: {ex}"})
                tabs[key] = None
                continue
        tab = tabs[key]
        if tab is None:
            results.append({**e, "action": "hold", "note": "タブ取得失敗(前述)"})
            continue
        action, updates, note = plan_entry(tab, e)
        results.append({**e, "action": action, "note": note,
                        "cells": [f"{col_letter(c)}{r + 1}" for r, c, _ in updates]})
        if action == "write":
            writes_by_tab.setdefault(key, []).append((tab, updates))
            # 同一実行内の後続エントリと衝突しないよう、ローカルにも反映
            for r, c, v in updates:
                while len(tab.rows) <= r:
                    tab.rows.append([])
                while len(tab.rows[r]) <= c:
                    tab.rows[r].append("")
                tab.rows[r][c] = v

    if not dry:
        for (period, month), items in writes_by_tab.items():
            data = []
            for tab, updates in items:
                for r, c, v in updates:
                    data.append({
                        "range": f"'{tab.title}'!{col_letter(c)}{r + 1}",
                        "values": [[v]],
                    })
            if data:
                api.spreadsheets().values().batchUpdate(
                    spreadsheetId=SHEETS[period],
                    body={"valueInputOption": "USER_ENTERED", "data": data},
                ).execute()

    print(json.dumps(results, ensure_ascii=False, indent=2))
    n = {"write": 0, "already": 0, "hold": 0}
    for r in results:
        n[r["action"]] += 1
    mode = "（dry-run: 書き込みなし）" if dry else ""
    print(f"\n書き込み {n['write']} / 転記済み {n['already']} / 保留 {n['hold']} {mode}",
          file=sys.stderr)
    if n["hold"]:
        sys.exit(2)


if __name__ == "__main__":
    main()
