#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""神栖横丁の口座引落 → 神栖横丁タブの水道光熱費・通信費ほか

--- なぜ銀行から入れるのか ----------------------------------------------
横丁の光熱費・電話・USENは請求書が2606月・2607月ぶんしかDropboxに無い。
9月〜5月が空欄のまま埋まらない。これらは全部【口座振替】なので、
神栖横丁の口座（千葉銀行 小見川支店 普通3543920）の出金がそのまま請求額になる。

★銀行を転記元に使うのはここだけ。bank.py は検算専用のまま。
  ここで使えるのは「毎月同じ相手に振替で落ちる固定的な費用」に限る。
  行の割り当ては利用者の指示（2026-08-21）。推測していない。

--- 行の割り当て（利用者指示 2026-08-21）-------------------------------
    ＡＰ（チイキソウセイ    → 水道光熱費（電気料金）
    カ）カタノシヨウテン     → 水道光熱費（ガス料金）
    デンワ７８／デンワ７７   → 通信費
    ＮＳ　ユウセン         → 通信費（USEN）
    ＭＨＦ）セキシヨウ      → 事務消耗品費
    ＤＦ．カミスシカンコウキ → 雑費
    ソンポジヤパン         → 保険料（非課税。2026-08-21 追加指示）
既に請求書から入っていた トウキヨウデンリヨク・スイドウリヨウキン も同じ考え方で入れる。

--- ★月ズレは取引先ごとに違う（これが要注意）---------------------------
2606月・2607月は請求書PDFから既に入っているので、それと銀行の引落額を
突き合わせて、取引先ごとのズレを実測した。1円まで合っている。

    地域創生   引落2026/07 366,794 ≒ 請求6月 333,450×1.1        → LAG=1
    片野商店   引落2026/07  20,050 ＝ 請求6月  18,228×1.1 ✅     → LAG=1
    水道       引落2026/07  72,979 ＝ 請求6月  66,345×1.1 ✅     → LAG=1
    東京電力   引落2026/07   1,449 ＝ 請求7月   1,318×1.1 ✅     → LAG=0
    電話78     引落2026/07   8,287 ＝ 請求7月   7,534×1.1 ✅     → LAG=0
    電話77     引落2026/06  11,880 ＝ 請求6月  10,800×1.1 ✅     → LAG=0
LAG=1 は「PL n月ぶんが n+1月に落ちる」。だから PL 9月は引落2025/10を見る。
引落2025/09 は20期8月ぶんなので使わない。

--- 税（非課税に注意）-------------------------------------------------------------------
損保ジャパンだけ非課税（損害保険料）。引落額をそのまま計上する。
残りは10%課税。税抜は 課税X ＋ floor(X*0.1) ＝ 引落額 を満たすXを探す（_ex）。
上の検算どおり、この求め方だと請求書の税抜と1円まで一致する。
解が無い月（端数が合わない月）は floor(引落×100/110) に落とす。

--- 入れない月 -----------------------------------------------------------
既に請求書から入っている月は飛ばす（二重計上防止）。飛ばすときは金額が
一致することを assert する。fixed_costs.py と同じ作法。
"""
import csv, glob, os, collections

DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bank")
ACCOUNT = "3543920"
TAB = "神栖横丁"
# ★8月を足した（2026-08-21）。利用者方針「8月分は全部CSVかPDFで入れていく」。
# 銀行明細のあるぶんだけ出るので、ファイルが無い月は自動的に空のまま。
MONTHS = ["9月", "10月", "11月", "12月", "1月", "2月", "3月",
          "4月", "5月", "6月", "7月", "8月"]
# PL列 → 銀行明細ファイルの年月
YM = {"9月": "202509", "10月": "202510", "11月": "202511", "12月": "202512",
      "1月": "202601", "2月": "202602", "3月": "202603", "4月": "202604",
      "5月": "202605", "6月": "202606", "7月": "202607", "8月": "202608"}
_ORDER = ["202509", "202510", "202511", "202512", "202601", "202602",
          "202603", "202604", "202605", "202606", "202607", "202608", "202609"]

# 摘要 → (PL行, LAG, メモ)
#   LAG=1: PL n月ぶんが n+1月に引き落とされる
VENDORS = {
    "ＡＰ（チイキソウセイ":     ("水道光熱費（電気料金）", 1, "地域創生ホールディングス。横丁の電気"),
    "トウキヨウデンリヨク":      ("水道光熱費（電気料金）", 0, "東京電力エナジーパートナー。共用部とみられる少額"),
    "カ）カタノシヨウテン":      ("水道光熱費（ガス料金）", 1, "片野商店"),
    "スイドウリヨウキン":        ("水道光熱費（水道料金）", 1, "神栖市水道事業"),
    "デンワ７８":              ("通信費", 0, "NTTファイナンス（回線1）"),
    "デンワ７７":              ("通信費", 0, "NTTファイナンス（回線2）"),
    "ＮＳ　ユウセン":           ("通信費（USEN）", 0, "USEN。毎月5,830円の定額"),
    "ＭＨＦ）セキシヨウ":        ("事務消耗品費", 0, "関彰商事"),
    "ＤＦ．カミスシカンコウキ":   ("雑費", 0, "神栖市観光協会"),
    # 2026-08-21 追加指示「損保ジャパン、入れて」
    "ソンポジヤパン":            ("保険料", 0, "損保ジャパン。損害保険料なので非課税"),
}

# 消費税がかからないもの。引落額をそのまま税抜として計上する。
# 本部のプルデンシャル47,090・ジブラルタ17,345も同じ扱いにしてある（fixed_costs.py）。
TAXFREE = {"ソンポジヤパン"}

# 入れないと決めたもの（いまは無し）
NOT_POSTED = {}

# 請求書PDFから既に入っている月（明細ログで確認 2026-08-21）。
# ★【取引先ごと】に持つこと。行ごとにすると事故る。
#   例: 事務消耗品費の6月は請求書がアスクル、口座振替は関彰商事。別の相手なので
#       両方が正しく、足すのが正解。行ごとに「6月は入っているから飛ばす」と
#       するとアスクルだけ残って関彰が落ちる。
FROM_INVOICE = {
    ("ＡＰ（チイキソウセイ", "6月"): 333450,     # 地域創生ホールディングス
    ("トウキヨウデンリヨク", "6月"): 1313,
    ("トウキヨウデンリヨク", "7月"): 1318,
    # ★2510月の請求書。横丁は食洗機と共同⑦の2メーターあり、引落を合算してから
    #   逆算すると1円ずれる（逆算10,119／請求書 8,372＋1,748＝10,120）。inv2509.py が入れる。
    ("カ）カタノシヨウテン", "10月"): 10120,
    ("カ）カタノシヨウテン", "6月"): 18228,      # 片野商店
    ("カ）カタノシヨウテン", "7月"): 19245,
    ("スイドウリヨウキン", "6月"): 66345,        # 神栖市水道事業
    ("スイドウリヨウキン", "7月"): 76810,
    ("デンワ７８", "6月"): 7646,                # NTTファイナンス 回線1
    ("デンワ７７", "6月"): 10800,               # NTTファイナンス 回線2（引落11,880＝10,800×1.1 ✅）
    ("デンワ７８", "7月"): 7534,
    ("ＮＳ　ユウセン", "6月"): 2560,             # USEN 1,280×2枚
    ("ＮＳ　ユウセン", "7月"): 2560,
}
# 請求書にはあるが口座振替に相手がいないもの（＝銀行から入れる対象外）
#   アスクル 事務消耗品費 6月 2,198 … カードか別口座払い。ここでは触らない


def _ex(inc):
    """引落（税込）から課税税抜を逆算。X + floor(X*0.1) == inc なるX。
    無ければ floor(inc*100/110)。"""
    base = inc * 10 // 11
    for x in (base - 1, base, base + 1, base + 2):
        if x + x // 10 == inc:
            return x, True
    return inc * 100 // 110, False


def _debits():
    """{摘要: {年月: 引落合計}}

    ★年月はファイル名ではなく【取引日の列】から取る。
      千葉銀行のエクスポートは 小見川支店_普通_3543920_202607_202608210925.csv の
      ように末尾に書き出し日時が付く。ファイル名の位置で切ると、
      そのまま置いたファイルで壊れる。取引日から取れば名前が何でも動く。
    """
    d = collections.defaultdict(lambda: collections.defaultdict(int))
    seen = set()
    for path in sorted(glob.glob(os.path.join(DIR, f"*_{ACCOUNT}_*.csv"))):
        for r in csv.DictReader(open(path, encoding="utf-8-sig")):
            if not r["出金金額(円)"]:
                continue
            # ★同じ月のCSVが2本置かれても二重に数えない。
            #   銀行のエクスポートは同じ期間を何度でも書き出せるので、
            #   古い名前のファイルが残っていると簡単に二重計上になる。
            #   日付・摘要・金額・残高が全部同じ行は同じ取引とみなす。
            key = (r["取引日"], r["摘要"], r["出金金額(円)"], r["残高(円)"])
            if key in seen:
                continue
            seen.add(key)
            k = r["摘要"].strip()
            if k in VENDORS:
                ym = r["取引日"].replace("/", "")[:6]   # 2026/07/03 -> 202607
                d[k][ym] += int(r["出金金額(円)"])
    return d


def _bank_ym(plmonth, lag):
    i = _ORDER.index(YM[plmonth]) + lag
    return _ORDER[i] if i < len(_ORDER) else None


def _by_row():
    """{(PL行, 月): [(摘要, 引落, 税抜, 端数ぴったりか, 銀行年月), ...]}"""
    d = _debits()
    out = collections.defaultdict(list)
    for raw, (plrow, lag, _note) in VENDORS.items():
        for m in MONTHS:
            ym = _bank_ym(m, lag)
            if ym is None or ym not in d[raw]:
                continue
            inc = d[raw][ym]
            ex, exact = (inc, True) if raw in TAXFREE else _ex(inc)
            out[(plrow, m)].append((raw, inc, ex, exact, ym))
    return out


def rows(wb=None):
    """(タブ, PL行, 月, 税抜, 元ファイル, メモ) を列挙。
    請求書から既に入っている【取引先】は、その月だけ飛ばす。"""
    agg = collections.defaultdict(list)
    for (plrow, m), items in _by_row().items():
        for raw, inc, ex, exact, ym in items:
            if (raw, m) in FROM_INVOICE:
                continue                     # その相手はもう請求書から入っている
            agg[(plrow, m)].append((raw, inc, ex, exact, ym))
    for (plrow, m), items in sorted(agg.items(), key=lambda x: (x[0][0], MONTHS.index(x[0][1]))):
        ex = sum(x[2] for x in items)
        inc = sum(x[1] for x in items)
        who = "＋".join(VENDORS[x[0]][2].split("。")[0] for x in items)
        yield (TAB, plrow, m, ex,
               f"銀行明細/21期/小見川支店_普通_{ACCOUNT}_{items[0][4]}*.csv",
               f"{who}。引落{inc:,}（税込）→ 税抜{ex:,}"
               + ("" if all(x[3] for x in items) else "。端数が割り切れず切り捨て"))


def skipped():
    """請求書と重なって飛ばした (取引先, PL行, 月, 請求書の額, 銀行から逆算した額)"""
    for (plrow, m), items in sorted(_by_row().items(), key=lambda x: MONTHS.index(x[0][1])):
        for raw, inc, ex, exact, ym in items:
            if (raw, m) in FROM_INVOICE:
                yield raw, plrow, m, FROM_INVOICE[(raw, m)], ex


def hold_rows():
    """埋まらない月と、銀行と請求書が食い違う月"""
    d = _debits()
    for raw, (plrow, lag, note) in sorted(VENDORS.items(), key=lambda x: x[1][0]):
        miss = [m for m in MONTHS
                if (raw, m) not in FROM_INVOICE
                and (_bank_ym(m, lag) is None or _bank_ym(m, lag) not in d[raw])]
        if 2 <= len(miss) < len(MONTHS):
            yield ("・".join(miss), TAB, plrow,
                   f"{note}。この月は口座振替が見つからず請求書も無いので空欄のまま。"
                   f"引落日が明細の期間外か、翌月にまとめて落ちている可能性")
    for raw, plrow, m, inv, bnk in skipped():
        if abs(inv - bnk) > 2:
            yield (m, TAB, plrow,
                   f"{VENDORS[raw][2].split('。')[0]}: 請求書からの計上 {inv:,} に対し、"
                   f"口座振替から逆算すると {bnk:,}（差 {bnk-inv:+,}）。"
                   f"請求書を正として計上済み。差の理由は未確認")
    for raw, why in NOT_POSTED.items():
        yield ("通年", TAB, raw, why)


def check(wb=None):
    """請求書と重なる月で、銀行から出した額が請求書と近いことを確認"""
    bad = []
    for raw, plrow, m, inv, bnk in skipped():
        if abs(inv - bnk) > 2:
            bad.append((plrow, m, inv, bnk))
    # 逆算のやり方が正しいことの確認: 請求書と1円まで合う月が4つ以上あること
    good = sum(1 for _, _, _, inv, bnk in skipped() if abs(inv - bnk) <= 1)
    assert good >= 4, f"請求書と一致した月が {good} しかない。逆算か月ズレが違う"
    return bad


if __name__ == "__main__":
    bad = check()
    print(f"{TAB} — 口座振替（千葉銀行 {ACCOUNT}）からの計上\n")
    print(f"{'PL行':<22}{'月':<5}{'引落(税込)':>11}{'税抜':>10}   内訳")
    print("-" * 88)
    t = 0
    for tab, plrow, m, ex, src, note in rows():
        inc = note.split("引落")[1].split("（")[0]
        print(f"{plrow:<22}{m:<5}{inc:>11}{ex:>10}   {note.split('。引落')[0]}")
        t += ex
    print("-" * 88)
    print(f"{'計':<27}{'':>11}{t:>10,}\n")
    print("請求書から既に入っていて飛ばした月（金額の突き合わせ）:")
    for raw, plrow, m, inv, bnk in skipped():
        mark = "✅" if abs(inv - bnk) <= 1 else f"★差 {bnk-inv:+,}"
        print(f"   {plrow:<20}{m:<4}{VENDORS[raw][2].split('。')[0][:12]:<14}"
              f"請求書 {inv:>8,} ／ 銀行 {bnk:>8,}  {mark}")
    print("\n保留:")
    for m, tab, item, why in hold_rows():
        print(f"   [{m}] {item}\n       {why}")
