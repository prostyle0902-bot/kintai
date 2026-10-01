#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""22期（2026年9月〜2027年8月）を組むときだけ、各モジュールの「期」を差し替える

21期のモジュール（debits / store_bank / yokocho_bank / honbu_bank）は、
月ラベル（"9月"…"8月"）→ 年月（"202509"…）の対応表をモジュールの先頭に持っている。
22期でも月ラベルは同じなので、対応表を22期の年月に差し替えれば同じロジックが使える。

★21期の月に結びついた例外表（請求書で入っている月・行の付け替え・既存PLとの検算値）は
  22期では意味が無い（同じ "9月" でも別の年）。ここで空にする。
  22期で例外が要るようになったら、各モジュールに 〜22 の表を作ってここで入れる。

★これを import するのは fill22.py だけ。21期の fill2.py は触らない
  （同じプロセスで21期と22期を混ぜないこと）。
"""
MONTHS = ["9月", "10月", "11月", "12月", "1月", "2月", "3月",
          "4月", "5月", "6月", "7月", "8月"]
START = (2026, 9)


def ym6(i):
    y, m = START[0] + (START[1] - 1 + i) // 12, (START[1] - 1 + i) % 12 + 1
    return f"{y:04d}{m:02d}"


YM6 = {m: ym6(i) for i, m in enumerate(MONTHS)}                 # "9月" -> "202609"
YM_SLASH = {m: f"{v[:4]}/{v[4:]}" for m, v in YM6.items()}      # "9月" -> "2026/09"
ORDER = [YM6[m] for m in MONTHS] + [ym6(12)]                   # 翌期の9月まで（LAG=1用）


def apply():
    import debits, store_bank, yokocho_bank, honbu_bank

    # --- debits（口座引落）: 期の初月を変えるだけで YM2M が作り直せる ------------
    debits.KESSAN_START = START
    debits.YM2M = {v: m for m, v in YM6.items()}
    debits._NBG_CACHE.clear()
    debits.CHECK = {}            # 既存21期PLとの検算値。22期には無い
    debits.EXCEPT = {}
    debits.DIFF = {}
    debits.ADD_ON = {}
    # ★利用者指示 2026-09-19「ハナリース料は本部で落としてるので、22期からはハナに転記して」
    for r in debits.RULES:
        if r["name"] == "千葉銀リース（本部）":
            r["split"] = [("韓国酒場ハナ", "リース料", 1, 1)]
            r["note"] = ("ちばぎんリース。本体口座（3351509）から引落23,980＝21,800×1.1。"
                         "中身は韓国酒場ハナのリースなので22期からはハナ「リース料」へ"
                         "（利用者指示 2026-09-19）")

    # --- 店舗口座・横丁口座 ---------------------------------------------------
    for mod in (store_bank, yokocho_bank):
        mod.YM = dict(YM6)
        mod._ORDER = list(ORDER)
        mod.FROM_INVOICE = {}
    store_bank.ROW_OVERRIDE = {}
    store_bank.EXIST_YAKITATE_GAS = {}

    # --- 本部口座 ------------------------------------------------------------
    honbu_bank.YM = dict(YM_SLASH)
