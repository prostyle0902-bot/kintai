#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""JCB法人カード / 三井住友カードの明細 → 本部タブ

freeeカード（engine.py）とは別系統。置き場所は同じ
`※請求書※/freeeカード明細/21期/` だが、発行会社が違うので明細の形式も違う。

【計上月：支払日ベース】（利用者確認済 2026-08-20）
    ファイル名の年月 − 1ヶ月 = PL列。
      202608meisai.csv（お支払日2026/08/10） → 7月列
      202609.csv      （支払月2026/09）      → 8月列
    理由：JCBは15日締め翌月10日払いで、1枚の明細が3ヶ月分の利用日にまたがる
    （8/10支払分の利用日は 2026/05/31〜07/15）。利用日ベースで切ると毎月2枚必要になり、
    月次の確定が約1ヶ月遅れるため。既存本部スプシの年計とも一致する。
    明細ログには利用日をそのまま残すので、後から利用日ベースへ切り替えることは可能。

【消費税】
    カード明細に税額の記載がないため 10% として逆算（円未満切り捨て・行単位）。
    軽減8%の可能性がある少額品（BASE FOOD 5,320／セブン-イレブン 2,558）も10%扱い。
    影響は624,249中の約7,878円ぶん。
"""

# 取引先キーワード → 本部のPL行
JCB_MASTER = {
    "旅費・交通費": ["楽天トラベル", "Suica", "ENEOS", "出光", "三井のリパーク",
                 "宮崎空港ビル", "東京国際空港", "都営宝町駐車場", "コスモ石油"],
    "接待交際費": ["創作和食藍とう", "花助", "阿見ゴルフクラブ", "ハイビスカスゴルフクラブ",
                "成田の森カントリークラブ", "ノースショアカントリークラブ", "R9TY龍ヶ崎",
                "はま寿司", "セブン-イレブン"],
    "通信費": ["ソフトバンクM", "NTTファイナンス", "NTTドコモビジネス", "日本郵便"],
    "事務消耗品費": ["AMAZON.CO.JP", "Amazon Market Place", "AMZN DIGITAL", "カインズホーム"],
    "租税公課": ["地方税共同機構", "システム利用料〔eLTAX〕"],
    "電気代　東京電力": ["東京電力"],
    # 新規行（利用者指示 2026-08-20）
    "福利厚生費": ["エルバランシア", "BASE FOOD"],
    "ソフトウェア利用料": ["fungoal", "ジョブカン", "board", "PEATIX", "ANTHROPIC", "OPENAI",
                   "アドビカブシキガイシャ", "GOOGLE", "Salon.jp", "MICROSOFT", "Voicy"],
}

# 三井住友は件数が少ないので個別指定
SMCC_MASTER = {
    "ハクロマーク製作所": "事務消耗品費",
    "テラルネツサンス": "その他経費",
}

# (ファイル, 発行会社, PL列, 明細合計の検算値)
FILES = [
    ("202608meisai.csv", "JCB", "7月", 624249),
    ("202609.csv", "三井住友", "8月", 112625),
]


# ---------------------------------------------------------------- 読み取り
import csv, io, os, unicodedata
from decimal import Decimal, ROUND_FLOOR

DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cards")


def _ex(tax_inc, rate=10):
    """税込 → 税抜（円未満切り捨て）"""
    return int((Decimal(tax_inc) / (1 + Decimal(rate) / 100)).to_integral_value(ROUND_FLOOR))


def _read_jcb(path):
    raw = open(path, "rb").read().decode("cp932")
    started = False
    for r in csv.reader(io.StringIO(raw)):
        if r and r[0] == "ご利用者":
            started = True; continue
        if started and len(r) >= 9 and r[2].strip():
            yield (r[2].strip(),                                   # ご利用日
                   unicodedata.normalize("NFKC", r[3]).strip(),    # ご利用先
                   int(r[8].replace(",", "")))                     # お支払い金額(税込)


def _read_smcc(path):
    for line in open(path, encoding="cp932"):
        p = line.rstrip("\n").split(",")
        if len(p) > 6 and p[0].strip():
            yield (p[0].strip(), unicodedata.normalize("NFKC", p[1]).strip(), int(p[6]))


def _classify(issuer, merchant):
    if issuer == "三井住友":
        return SMCC_MASTER.get(merchant)
    hit = [k for k, ws in JCB_MASTER.items() if any(w in merchant for w in ws)]
    return hit[0] if len(hit) == 1 else None


def rows():
    """(タブ, 取引先, PL行, 税抜, 消費税, 元ファイル, 月, 利用日) を列挙。
    分類できないものは PL行=None で返す（呼び出し側が保留リストへ）。"""
    for fname, issuer, month, expect in FILES:
        path = os.path.join(DIR, fname)
        reader = _read_jcb if issuer == "JCB" else _read_smcc
        det = list(reader(path))
        total = sum(v for _, _, v in det)
        assert total == expect, f"{fname}: 明細合計 {total:,} ≠ 検算値 {expect:,}"
        for used, merchant, inc in det:
            plrow = _classify(issuer, merchant)
            ex = _ex(inc)
            yield ("本部", merchant, plrow, ex, inc - ex,
                   f"freeeカード明細/21期/{fname}", month, used, issuer)
