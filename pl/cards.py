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

【カード明細のルール】（利用者指示 2026-08-21）
    明細を読んだ月は、明細1行ずつを費目へ仕訳して入れる。
    一括行（本部の「JCBカード」「三井住友カード」）には入れない。
    明細の無い月だけ、既存21期PLの一括額を exist_fill.py が暫定で置く。
    同じ月に両方入ると二重計上になるので exist_fill.check(wb) が止める。
    ★下の FILES に1行足せば、その月は自動で一括→費目別に切り替わる。
      exist_fill.split_by_cards() が FILES を見て判断しているため。

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
# (ファイル, 発行会社, PL列, 明細合計の検算値)
# ★PL列 = ファイル年月 −1か月（支払日ベース。冒頭の説明を参照）。
#   JCBは11枚とも既存21期PL「JCBカード」の月額と1円まで一致した（2026-08-22 実測）。
FILES = [
    ("202510meisai.csv", "JCB", "9月", 991395),
    ("202511meisai.csv", "JCB", "10月", 497745),
    ("202512meisai.csv", "JCB", "11月", 635785),
    ("202601meisai.csv", "JCB", "12月", 789988),
    ("202602meisai.csv", "JCB", "1月", 532674),
    ("202603meisai.csv", "JCB", "2月", 347904),
    ("202604meisai.csv", "JCB", "3月", 410596),
    ("202605meisai.csv", "JCB", "4月", 339754),
    ("202606meisai.csv", "JCB", "5月", 511265),
    ("202607meisai.csv", "JCB", "6月", 538175),
    ("202608meisai.csv", "JCB", "7月", 624249),
    ("202510.csv", "三井住友", "9月", 30000),
    ("202511.csv", "三井住友", "10月", 77377),
    # ★11月だけ既存21期PLは30,771で、明細の合計行30,071と700円違う。
    #   書類を正とする方針どおり明細の額を採る。差は「既存PLとの食い違い」タブに出る。
    ("202512.csv", "三井住友", "11月", 30071),
    ("202601.csv", "三井住友", "12月", 8000),
    ("202602.csv", "三井住友", "1月", 5000),
    ("202603.csv", "三井住友", "2月", 74500),
    ("202604.csv", "三井住友", "3月", 5000),
    ("202605.csv", "三井住友", "4月", 32400),
    ("202606.csv", "三井住友", "5月", 8000),
    # 6月・7月は既存21期PLが空欄。明細から新たに入る
    ("202607.csv", "三井住友", "6月", 7360),
    ("202608.csv", "三井住友", "7月", 33100),
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
        if not started or len(r) < 9:
            continue
        # ★カテゴリ ≪その他≫ の行は利用日が空のことがある
        #   （「ＥＴＣスルーカードＮ ご利用おまとめ」。その月ぶんの合算なので日付が無い）。
        #   列の位置はふつうの明細と同じで、金額は列8。利用日だけで弾くと
        #   検算が合わなくなる（202601で2,580円ぶん足りなかった）。
        #   年会費（法人ゴールドカード年会費）も同じ ≪その他≫ だが利用日は入っている。
        other = "≪その他≫" in r[1]
        if not (r[2].strip() or other):
            continue
        # ★訂正サインが「取消」の行はお支払い金額が空。プラスとマイナスの
        #   2行で打ち消してあるので、そのまま飛ばす（202510で実測）。
        if not r[8].strip():
            continue
        yield (r[2].strip(),                                   # ご利用日（無いこともある）
               unicodedata.normalize("NFKC", r[3]).strip(),    # ご利用先
               int(r[8].replace(",", "")))                     # お支払い金額(税込)


def _read_smcc(path):
    """三井住友。★2つの形式がある（2026-08-22 に判明）。

    ① 支払照会（202609.csv だけがこれ）
         利用日,利用先,本人区分,支払区分,,支払月,金額,金額,...
         列0=利用日／列1=利用先／列6=金額。ヘッダ行なし。
    ② 明細照会（202510〜202608。今回まとめて入った12枚）
         カード名義の行で始まり、明細は
           利用日,利用先,金額,回数,回数,金額,備考
         ★年会費とWEB明細書年会費割引だけ列2が空で、列5に金額が入る。
           列2だけ見ていると825円ぶん取りこぼす（202511で実測）。
         ★最終行が合計行（列0が空・列5に総額）。ここは足さない。
         ★カードが2枚あると名義行が途中にもう一度出る（202607がそう）。
    どちらも金額に返品のマイナスが混じる。
    """
    lines = open(path, encoding="cp932").read().splitlines()
    fmt2 = any("様" in l.split(",")[0] for l in lines[:1])
    for line in lines:
        p = line.rstrip("\n").split(",")
        if not fmt2:                                   # ① 支払照会
            if len(p) > 6 and p[0].strip():
                yield (p[0].strip(), unicodedata.normalize("NFKC", p[1]).strip(),
                       int(p[6]))
            continue
        # ② 明細照会
        if len(p) < 6:
            continue
        if "様" in p[0]:                                # カード名義の行
            continue
        if not p[0].strip() and not p[1].strip():       # 最終行の合計
            continue
        amt = p[2].strip() or p[5].strip()              # 年会費は列5にある
        if not amt:
            continue
        try:
            v = int(amt.replace(",", ""))
        except ValueError:
            continue
        used = p[0].strip() or "（年会費）"
        yield (used, unicodedata.normalize("NFKC", p[1]).strip(), v)


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
