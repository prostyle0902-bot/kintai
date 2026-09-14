#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dropbox に置かれたPL用の元データを pl/ のどこに置くか決め、取り込み済みを覚える

定期チェック（Routine）から使う。Dropbox へ触れるのは Claude のコネクタだけなので、
**一覧と実ダウンロードは Claude がやり、置き場所の判定と台帳はここが持つ**。

    # 1) Claude が Dropbox を再帰リストして [{file_id, name, path}, …] を作る
    python3 ingest.py --plan listing.json
        → 未取り込みのファイルについて「どこへ置くか」を出す。
          置き場所が決まらないものは unknown として出る（★推測で置かない）

    # 2) Claude が download_link + curl で plan のとおりに落とす

    # 3) 落とせたものを台帳に記録する
    python3 ingest.py --record records.json
        records.json は [{file_id, name, path, local}, …]

台帳 `ingest_ledger.json` は file_id で照合する。Dropbox でファイルが
支払い済/ へ移動されても ID は変わらないので、二重取り込みにならない
（furikomi/ledger.json と同じ考え方）。

--- 置き場所のルール -----------------------------------------------------
元データごとに「ファイル名 → pl/ のどこ」が決まっている。月のずれ方が
ソースごとに違うので、ここに1か所へまとめてある（status8.py の冒頭も参照）。

| Dropboxのファイル名                        | 置き場所                    | 月のずれ |
|--------------------------------------------|-----------------------------|----------|
| 小見川支店_普通_<口座>_<YYYYMM>_<時刻>.csv | bank/小見川支店_普通_<口座>_<YYYYMM>.csv | 取引月 |
| statement-YYYY-MM.csv                      | csv/ 同名                   | −1か月＝利用月 |
| <YYYYMM>meisai.csv（JCB）                  | cards/ 同名                 | −1か月＝PL列 |
| <YYYYMM>.csv（三井住友）                   | cards/ 同名                 | −1か月＝PL列 |
| ENEOS_<YYYYMM>.csv                         | eneos/ENEOS_<PL月>.csv      | −2か月＝PL列 |
| 0000042400_高速_CP請求鑑_<YYYYMM>.csv      | rikuji/CP請求鑑_<PL月>.csv  | 同月＝PL列 |
| <店舗>_<YYYYMM>-<YYYYMM>.csv（Airレジ21期）| airegi/ 同名                | 行ごとに月 |
| <店舗>_<YYYYMMDD>-<YYYYMMDD>.csv（22期）   | airegi/<YYMM月>/ 同名       | 行ごとに日 |

★ここに無いもの（NBGの乱数名CSV・なめがた/陸事のPDF・かめや・出前館など）は
  unknown で出す。置き場所や読み取りに判断が要るので、利用者に報告してから入れる。

★**ファイルを置いただけでは8月が埋まらないソースがある。**
  statement / JCB / ENEOS / 陸事 は、モジュール側の月リストに1行足して
  はじめて拾われる（pl/README.md「8月分の受け入れ」の表）。plan はそれも出す。
"""
import json
import os
import re
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
LEDGER = os.path.join(BASE, "ingest_ledger.json")

# 引落月・支払月から PL列 へ戻す月数
_BACK = {"eneos": 2}


def _pl_month(yyyymm, back):
    """'202610' と 2 から '8月' を作る"""
    y, m = int(yyyymm[:4]), int(yyyymm[4:6])
    m -= back
    while m <= 0:
        m += 12
        y -= 1
    return f"{m}月"


def classify(name):
    """Dropboxのファイル名 → (置き場所, 追記が要るモジュール or None, メモ)

    置き場所が決まらないときは (None, None, 理由) を返す。
    """
    m = re.match(r"^(小見川支店_普通_\d+_\d{6})_\d+\.csv$", name)
    if m:
        return f"bank/{m.group(1)}.csv", None, "千葉銀行。置けば自動で拾う"

    if re.match(r"^statement-\d{4}-\d{2}\.csv$", name):
        return (f"csv/{name}", "engine.py",
                "freeeカード明細。ファイル年月−1か月＝利用月。engine.py の一覧に1行要る")

    m = re.match(r"^(\d{6})meisai\.csv$", name)
    if m:
        return (f"cards/{name}", "cards.py",
                f"JCB。支払{m.group(1)}＝PL列は{_pl_month(m.group(1), 1)}。"
                "cards.py の FILES に1行要る")

    m = re.match(r"^(\d{6})\.csv$", name)
    if m:
        return (f"cards/{name}", "cards.py",
                f"三井住友。支払{m.group(1)}＝PL列は{_pl_month(m.group(1), 1)}")

    m = re.match(r"^ENEOS_(\d{6})\.csv$", name)
    if m:
        pl = _pl_month(m.group(1), _BACK["eneos"])
        return (f"eneos/ENEOS_{pl}.csv", "eneos.py",
                f"ENEOS。引落{m.group(1)}＝PL列は{pl}。"
                "eneos.py の MONTHS と PAY_MONTH に1行要る")

    m = re.match(r"^\d+_高速_CP請求鑑_(\d{6})\.csv$", name)
    if m:
        pl = _pl_month(m.group(1), 0)
        return (f"rikuji/CP請求鑑_{pl}.csv", "rikuji.py",
                f"陸事総合のCP請求鑑。利用{m.group(1)}＝PL列は{pl}。"
                "金額の正は請求書PDFなので rikuji.py の PDF_DATA にも1行要る")

    m = re.match(r"^(.+)_(\d{6})-(\d{6})\.csv$", name)
    if m:
        return f"airegi/{name}", None, "Airレジ集計（21期・年間）。置けば自動で拾う"

    m = re.match(r"^(.+)_(\d{8})-(\d{8})\.csv$", name)
    if m:
        ym = m.group(2)
        return (f"airegi/{ym[2:4]}{ym[4:6]}月/{name}", None,
                "Airレジ集計（22期・日次）。置けば自動で拾う")

    return None, None, "置き場所のルールが無い。利用者に報告してから入れる"


def load_ledger():
    if not os.path.exists(LEDGER):
        return {"note": "PL用元データの取り込み台帳。file_idで照合する",
                "ingested": {}, "skipped": {}}
    with open(LEDGER, encoding="utf-8") as f:
        return json.load(f)


def save_ledger(d):
    with open(LEDGER, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False, indent=1)
        f.write("\n")


def plan(listing_path):
    with open(listing_path, encoding="utf-8") as f:
        entries = json.load(f)
    led = load_ledger()
    known = set(led["ingested"]) | set(led["skipped"])

    todo, unknown, mods = [], [], {}
    for e in entries:
        if e["file_id"] in known:
            continue
        local, mod, note = classify(e["name"])
        if local is None:
            unknown.append({**e, "note": note})
            continue
        if os.path.exists(os.path.join(BASE, local)):
            note += "（★同名がもう pl/ にある。中身が違えば上書きになる）"
        todo.append({**e, "local": local, "module": mod, "note": note})
        if mod:
            mods.setdefault(mod, []).append(e["name"])

    print(json.dumps({"todo": todo, "unknown": unknown}, ensure_ascii=False, indent=1))
    print(f"\n--- 取り込む {len(todo)}件 ／ 置き場所不明 {len(unknown)}件", file=sys.stderr)
    for mod, names in sorted(mods.items()):
        print(f"★{mod} に追記が要る: {'、'.join(names)}", file=sys.stderr)
    return 0


def record(records_path):
    with open(records_path, encoding="utf-8") as f:
        recs = json.load(f)
    led = load_ledger()
    for r in recs:
        dst = led["skipped"] if r.get("skip") else led["ingested"]
        dst[r["file_id"]] = {k: v for k, v in r.items() if k != "file_id"}
    save_ledger(led)
    print(f"台帳: 取り込み {len(led['ingested'])}件 ／ skip {len(led['skipped'])}件")
    return 0


def main():
    a = sys.argv[1:]
    if len(a) == 2 and a[0] == "--plan":
        return plan(a[1])
    if len(a) == 2 and a[0] == "--record":
        return record(a[1])
    sys.exit(__doc__)


if __name__ == "__main__":
    sys.exit(main())
