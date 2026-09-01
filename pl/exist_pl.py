#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""既存の21期PLシート（exist/*.json）を読んで (タブ, 行名, 月) → 値 にする

exist_dump.py が Sheets API で落としたJSONを読むだけ。ネットにはつながない。
既存シートには絶対に書き込まない（読むだけ）。

10枚とも同じ形だった（2026-08-21 確認）:
    4行目がヘッダ  列1=項目 ／ 列2=年計 ／ 列3〜14=9月〜8月
    以降の行       列0=大分類（売上/売上原価/…）／ 列1=項目名 ／ 列3〜=月次

★同名の行が2枚ある（2026-08-21 に全10枚を走査して確認）:
    焼きたて屋「消費税」×2 … 1つ目は売上の消費税、2つ目は出前館の消費税
    りゅうちゃん「〃」×2   … どちらも作業用
    2つ目以降は「消費税#2」のように連番を付けて別の行として返す。
    最初に書いたときは後勝ちで潰していて、焼きたて屋の11か月を取り違えていた。
"""
import glob, json, os

BASE = os.path.dirname(os.path.abspath(__file__))
DIR = os.path.join(BASE, "exist")
MONTHS = ["9月", "10月", "11月", "12月", "1月", "2月", "3月",
          "4月", "5月", "6月", "7月"]


def _num(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return int(v)
    return None


def load():
    """{タブ: {行名: {月: 値}}} と {タブ: [行名の並び]} を返す。"""
    out, order = {}, {}
    for p in sorted(glob.glob(os.path.join(DIR, "*.json"))):
        tab = os.path.basename(p)[:-5]
        d = json.load(open(p))
        # 「9月」「10月」が並ぶヘッダを持つタブを1枚選ぶ（PL本体）
        for name, rows in d["sheets"].items():
            hi = next((i for i, r in enumerate(rows)
                       if "9月" in r and "10月" in r), None)
            if hi is None:
                continue
            h = rows[hi]
            col = {m: h.index(m) for m in MONTHS if m in h}
            vals, seq = {}, []
            for r in rows[hi + 1:]:
                if len(r) < 2 or not str(r[1]).strip():
                    continue
                item = str(r[1]).strip()
                mv = {m: _num(r[c]) for m, c in col.items()
                      if c < len(r) and _num(r[c]) is not None}
                # ★同名の行がある。あとから来たほうで潰さず、別の行として扱う。
                #   焼きたて屋は「消費税」が2つ（売上ぶんと出前館ぶん）。
                #   りゅうちゃんは「〃」が2つ。潰すと11か月ぶん取り違える。
                if item in vals:
                    k = 2
                    while f"{item}#{k}" in vals:
                        k += 1
                    item = f"{item}#{k}"
                seq.append(item)
                vals[item] = {}
                vals[item].update(mv)
            out[tab] = vals
            order[tab] = seq
            break
    return out, order


if __name__ == "__main__":
    ex, order = load()
    print(f"{'タブ':<14}{'行':>4}{'値のあるセル':>10}{'年計(9〜7月)':>14}")
    print("-" * 46)
    for tab in order:
        cells = sum(len(v) for v in ex[tab].values())
        print(f"{tab:<14}{len(order[tab]):>4}{cells:>10}")
