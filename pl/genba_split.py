#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""業務課スタッフの人件費を、現場カレンダーの「鳥害対策」の日数で鳥害対策課へ分ける（22期）

--- 利用者指示 2026-10-06 -------------------------------------------------
「昨日作ったアプリ（genba.html）から人件費の仕訳出来ないかな」
→ 提案「人ごとに『鳥害の日数 ÷ 現場の日数』で給料と社会保険を業務課・鳥害対策課に分ける」
   に「はい」。小林俊樹（名簿では鳥害対策課）も同じく割合で分ける（「はい」）。

--- 元データ -------------------------------------------------------------
現場管理カレンダーの保存先スプレッドシート「予定」（シート名も「予定」）。
genba.html / prostyle-calendar が読み書きしているのと同じもの。
    列: id date place client type outsource staff start end amount memo ...
    type が "bird" の予定 ＝ 鳥害対策
pl-writer（サービスアカウント）に閲覧で共有済み（利用者 2026-10-06）。

--- 数え方（genba.html の computeSites と同じ） -----------------------------
1日を1人1日として、その日に行った予定で分ける。
時間（start/end）が入っていれば時間の長さで、どれにも入っていなければ均等に分ける。
    例: 同じ日に 鳥害6時間＋清掃2時間 → 鳥害0.75日・清掃0.25日
月の合計で  鳥害の割合 ＝ 鳥害の日数 ÷ 現場の日数ぜんぶ。

★カレンダーに載っていない人（日常清掃のアルバイトさんなど）は今までどおり名簿の店。
★外注（SUN-X・RISE など）は給料一覧表に載らないので関係しない（請求書から入る）。
★本部の人（飯田栄＝社長）は分けない。分けるのは名簿で業務課・鳥害対策課の人だけ。

--- 控え（genba/days.json） ------------------------------------------------
人ごと・月ごとの「鳥害の日数・現場の日数」だけを残す（給与額は入らない）。
★一度控えた月は、カレンダーが後から書き換わっても動かさない（PLが勝手に変わらないように）。
  やり直すときは  python3 genba_split.py --refresh 2026-09
取り込み:
    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 genba_split.py          # 終わった月を足す
"""
import collections
import datetime
import json
import os
import re
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
DAYS = os.path.join(BASE, "genba", "days.json")
SHEET_ID = "1yreBklsmvcpnbvNJv1eE8wCLhmg2fC4HJrHhDbPxEQQ"   # 予定
SHEET = "予定"

# 分ける対象の店（名簿でこのどちらかの人だけ）と、鳥害ぶんの行き先
TABS = ("業務課", "鳥害対策課")
BIRD_TAB = "鳥害対策課"
OTHER_TAB = "業務課"
PERIODS = {"22期": ("2026-09", "2027-08")}


def _hours(start, end):
    t = lambda v: (lambda m: int(m[1]) + int(m[2]) / 60 if m else None)(
        re.match(r"^(\d{1,2}):(\d{2})$", (v or "").strip()))
    s, e = t(start), t(end)
    if s is None or e is None:
        return 0
    h = e - s
    return h + 24 if h <= 0 else h


def count(rows, ym):
    """{名前: [鳥害の日数, 現場の日数]}（ym = 'YYYY-MM'）"""
    byday = collections.defaultdict(list)
    for r in rows:
        if r.get("date", "").startswith(ym) and (r.get("place") or "").strip():
            byday[r["date"]].append(r)
    acc = collections.defaultdict(lambda: [0.0, 0.0])
    for lst in byday.values():
        per = collections.defaultdict(list)
        for r in lst:
            for n in [x.strip() for x in (r.get("staff") or "").split(",") if x.strip()]:
                per[n].append(r)
        for n, mine in per.items():
            tot = sum(_hours(x.get("start"), x.get("end")) for x in mine)
            for x in mine:
                h = _hours(x.get("start"), x.get("end"))
                sh = h / tot if tot > 0 else 1 / len(mine)
                acc[n][1] += sh
                if x.get("type") == "bird":
                    acc[n][0] += sh
    return {n: [round(b, 4), round(t, 4)] for n, (b, t) in sorted(acc.items())}


def _load():
    if not os.path.exists(DAYS):
        return {}
    with open(DAYS, encoding="utf-8") as f:
        return json.load(f)


def fetch(refresh=()):
    """カレンダーを読んで、終わった月（と refresh で指定した月）を控えに足す。"""
    import push
    from googleapiclient.discovery import build
    svc = build("sheets", "v4", credentials=push._creds())
    v = svc.spreadsheets().values().get(spreadsheetId=SHEET_ID, range=f"'{SHEET}'").execute()
    vals = v.get("values", [])
    hdr = vals[0]
    rows = [dict(zip(hdr, r + [""] * (len(hdr) - len(r)))) for r in vals[1:]]
    data = _load()
    this = datetime.date.today().strftime("%Y-%m")
    lo, hi = min(p[0] for p in PERIODS.values()), max(p[1] for p in PERIODS.values())
    months = sorted({r["date"][:7] for r in rows if r.get("date")})
    added = []
    for ym in months:
        if not (lo <= ym <= hi) or ym >= this:
            continue                     # 期の外・まだ終わっていない月は控えない
        if ym in data and ym not in refresh:
            continue                     # 控え済みの月は動かさない
        data[ym] = count(rows, ym)
        added.append(ym)
    os.makedirs(os.path.dirname(DAYS), exist_ok=True)
    with open(DAYS, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1, sort_keys=True)
    return added


def days(period, month):
    """22期の 'N月' → {名前: (鳥害の日数, 現場の日数)}。控えが無ければ None。"""
    if period not in PERIODS:
        return None
    n = int(month.rstrip("月"))
    y = int(PERIODS[period][0][:4]) + (0 if n >= 9 else 1)
    got = _load().get(f"{y}-{n:02d}")
    return None if got is None else {k: tuple(v) for k, v in got.items()}


def bird_part(amount, b, t):
    """鳥害ぶん（円未満切り捨て）。残りは業務課。"""
    if t <= 0 or b <= 0:
        return 0
    return int(amount * round(b / t, 6))


if __name__ == "__main__":
    if "--refresh" in sys.argv:
        added = fetch(refresh=tuple(sys.argv[sys.argv.index("--refresh") + 1:]))
    else:
        added = fetch()
    print("控えた月:", ", ".join(added) or "なし")
