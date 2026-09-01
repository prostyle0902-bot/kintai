#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""本部の口座振替（水道・電気代 地域創生・ガス代・保険料）→ 本部タブ

--- なぜ要るのか ---------------------------------------------------------
利用者の問い（2026-08-25）「本部経費だが、水道代、電気代（地域創生）、ガス代、
保険料って、口座から拾えないかな」。
拾えた。千葉銀行 小見川支店 普通3351509（本体・ビルメンのメイン口座）に
毎月きれいに落ちている。これまでこの口座から口座振替を拾うモジュールが無く、
4つとも9月〜3月（既存21期PL由来）で止まっていた。

    店舗の口座 → store_bank.py ／ 横丁の口座 → yokocho_bank.py
    本体の口座 → ここ（新規）

--- 摘要 -----------------------------------------------------------------
    水道           「スイドウリヨウキン」＋「カトリシスイドウ」の2本立て
                   カトリシスイドウは毎月1,870で一定（香取市水道）
    電気代 地域創生 「ＡＰ（チイキソウセイ」。2025/12から始まっている
    ガス代         「ガスリヨウキン」
    保険料         「ソンポジヤパン」。毎月2〜3本（13,920＋34,987＋86,000前後など）
                   2026/01だけ5本で199,722（60,200×2が上乗せ）

--- ★計上月と税の決め方（2026-08-25 に全パターン検証した）----------------
既存21期PLの9月〜3月と、4通り（引落月そのまま／1か月前 × 税込／税抜）を
総当たりで突き合わせた。結果:
    水道           引落月そのまま・税込 が3か月一致（ほかは1か月以下）
    電気代 地域創生 引落月そのまま・税込 が3か月一致（ほかは0）
    ガス代         1か月前・税込 4か月 ／ 引落月そのまま・税込 3か月
    保険料         引落月そのまま 5か月一致（ほかは2か月）
★会計士の付け方は月によってバラバラで、どの規則にも完全には従っていない。
  いちばん当たりが多い【引落月＝PL列】を採る。保険料が5/7で決め手になった。
★金額は【税抜】。会計士はこの4行を税込で書いているが、この損益計算書は
  税抜でそろえる方針（ニッセーデリカなどと同じ扱い）。
  水道・電気・ガスは10%、保険料は非課税。

--- 手元にある明細 -------------------------------------------------------
口座3351509は 2025/09〜2026/07 の11か月。引落月＝PL列なので9月〜7月が埋まる。
★8月だけ埋まらない。202608の明細がまだDropboxに無い（status8.py 参照）。
"""
import collections

import bank

ACCOUNT = "3351509"
TAB = "本部"

# PL行 -> (摘要のリスト, 税率)
GROUPS = {
    "水道": (["スイドウリヨウキン", "カトリシスイドウ"], 10),
    "電気代　地域創生": (["ＡＰ（チイキソウセイ"], 10),
    "ガス代": (["ガスリヨウキン"], 10),
    # 損害保険料は非課税（ほかのモジュールと同じ扱い）
    "保険料（損保ジャパン）": (["ソンポジヤパン"], 0),
}

MONTHS = ["9月", "10月", "11月", "12月", "1月", "2月", "3月",
          "4月", "5月", "6月", "7月", "8月"]
YM = {"9月": "2025/09", "10月": "2025/10", "11月": "2025/11", "12月": "2025/12",
      "1月": "2026/01", "2月": "2026/02", "3月": "2026/03", "4月": "2026/04",
      "5月": "2026/05", "6月": "2026/06", "7月": "2026/07", "8月": "2026/08"}


def _ex(inc, rate):
    """引落（税込）から税抜Xを逆算。X + floor(X*rate/100) == inc なるX。"""
    if rate == 0:
        return inc, True
    base = inc * 100 // (100 + rate)
    for x in (base - 1, base, base + 1, base + 2):
        if x + x * rate // 100 == inc:
            return x, True
    return base, False


def _debits():
    """{摘要: {年月: (合計, [明細])}}"""
    d = collections.defaultdict(lambda: collections.defaultdict(list))
    for r in bank._load_chiba(accounts={ACCOUNT}):
        d[r["raw"]][r["date"][:7]].append((r["date"], r["amount"], r["src"]))
    return d


def rows():
    """(タブ, PL行, 月, 税抜, 元ファイル, メモ) を列挙。"""
    d = _debits()
    for plrow, (keys, rate) in GROUPS.items():
        for m in MONTHS:
            det = [x for k in keys for x in d[k].get(YM[m], [])]
            if not det:
                continue
            inc = sum(a for _dt, a, _s in det)
            ex, exact = _ex(inc, rate)
            src = sorted({s for _dt, _a, s in det})[0]
            what = "／".join(f"{dt} {a:,}" for dt, a, _s in sorted(det))
            note = (f"千葉銀行3351509（本体）の口座振替 {len(det)}件: {what}"
                    f"＝税込{inc:,}"
                    + (f" → 税抜{ex:,}（{rate}%）" if rate else "（保険料は非課税）")
                    + "。計上月は引落月そのまま（2026-08-25 に既存PLと総当たりで検証）"
                    + ("" if exact else "。端数が割り切れず切り捨て"))
            yield TAB, plrow, m, ex, f"銀行明細/21期/{src}", note


def hold_rows():
    """埋まらない月を保留リストへ。"""
    d = _debits()
    for plrow, (keys, _rate) in GROUPS.items():
        miss = [m for m in MONTHS
                if not any(d[k].get(YM[m]) for k in keys)]
        if miss:
            yield ("・".join(miss), TAB, plrow,
                   f"千葉銀行3351509（本体）にこの月の引落が無い。"
                   f"8月は明細CSV（小見川支店_普通_3351509_202608）が"
                   f"まだ手元に無いため。ほかの月は契約が始まる前か引落が無い月")


def check(wb=None):
    """明細の取りこぼしが無いか／書き込み先が空か。"""
    d = _debits()
    # 摘要ごとの総額が、rows() が出す税込の総額と合うこと
    for plrow, (keys, rate) in GROUPS.items():
        allsum = sum(a for k in keys for ym in d[k] for _dt, a, _s in d[k][ym])
        got = 0
        for _t, r, m, _ex, _s, note in rows():
            if r != plrow:
                continue
            got += int(note.split("＝税込")[1].split("。")[0].split(" →")[0]
                        .replace(",", "").replace("（保険料は非課税）", ""))
        assert allsum == got, f"{plrow}: 明細の総額 {allsum:,} ≠ 計上した税込 {got:,}"
    if wb is None:
        return
    import build2
    for tab, plrow, m, val, _src, _note in rows():
        assert plrow in build2.RIDX[tab], f"{tab} に「{plrow}」行が無い"


if __name__ == "__main__":
    check()
    import collections as c
    agg = c.defaultdict(dict)
    for _t, plrow, m, v, _s, _n in rows():
        agg[plrow][m] = v
    print("本部の口座振替（千葉銀行3351509・税抜・計上月＝引落月）\n")
    print(f"{'PL行':<22}" + "".join(f"{m:>9}" for m in MONTHS) + f"{'計':>11}")
    for plrow in GROUPS:
        v = agg[plrow]
        print(f"{plrow:<22}" + "".join(f"{v.get(m,0):>9,}" for m in MONTHS)
              + f"{sum(v.values()):>11,}")
    print(f"\n計 {sum(sum(v.values()) for v in agg.values()):,}円（税抜）")
    print("\n保留:")
    for m, tab, plrow, why in hold_rows():
        print(f"   [{m}] {plrow}")
