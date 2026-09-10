#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""出前館（焼きたて屋）の売上・手数料 — 月次支払通知書PDFから

--- 元データ -------------------------------------------------------------
Google Drive の共有フォルダ（オーナー: brothertak83davab@gmail.com）
    2.焼きたて屋 / 取引集計（出前館） / 13608422_YYYYMM_payment.pdf
    フォルダID: 14TTFrm69mDNHiAFo68PTVrdNic3DXgvV
    2025年分は「2025」サブフォルダ: 1dmfYPNaMBGwCb9s4aUZp-8NOmmLvgTiS

21期は 2025.09〜2026.07 の全月がそろっている。
サービスアカウントには共有されていないので、Driveコネクタで read_file_content する。
店舗コード YM8QF_0101（焼きたて屋コメリパワー鹿嶋店）の1店舗のみ。

--- PDFの構造 -----------------------------------------------------------
    加盟店売上合計①      = 現金決済② ＋ ネット決済③ ＋ ポイント/クーポン④
    出前館利用料（税込）⑥ = サービス利用料10% ＋ 配達代行手数料25%
                            ＋ 振込手数料 ＋ 決済手数料
                            「10％対象計 ￥x (内税 ￥y)」に消費税が書いてある
    お戻し金額⑦          = 商品代金補填（不課税）。件数ぶんの返金
    支払金額 = ① − ⑥ + ⑦

--- PL行への割り当て -----------------------------------------------------
21期:
    出前館売上（税込）   ← ①              （既存PLと全月一致）
    支払手数料（出前館）  ← ⑥の税抜（＝⑥ − 内税）
    支払手数料（出前館返金）← ⑦をマイナス計上
    出前館消費税は既存の計算方法のまま触らない

22期以降:
    出前館消費税        ← ①から内税抽出（8/108・軽減税率）に変更する

★訂正（2026-08-20）
  「支払手数料（出前館）は21期に計上漏れ」と以前このファイルに書いたが、
  それは誤り。既存PLスプシ（焼きたて屋鹿島店 2025.9~2026.8損益計算書
  file_id 1OB3yLta9y1d14DbN3jeqCoL9MKFknUbBnqR1WBOvLYQ）を実際に読んだところ、
  支払手数料（出前館）年計 ¥140,843、支払手数料（出前館返金）年計 ▲¥20,890 が
  すでに全月入っていた。しかも本モジュールの計算値と11か月とも1円まで一致する。
  → 計上漏れではなく、逆に本モジュールの検算になっている（EXIST_21 で assert）。

★出前館消費税だけは既存と方法が違う
  既存は ①×8%（四捨五入）。正しくは ①×8/108 の内税抽出。
  例: 9月 ① 66,670 → 既存 5,334 ／ 内税抽出 4,938。
  期の途中で方法を変えると混在するので、yakitateya.py の A-2 と同じ考え方で
  22期から揃える。
"""
from decimal import Decimal, ROUND_FLOOR, ROUND_HALF_UP

FOLDER_ID = "14TTFrm69mDNHiAFo68PTVrdNic3DXgvV"
FOLDER_ID_2025 = "1dmfYPNaMBGwCb9s4aUZp-8NOmmLvgTiS"
TAB = "焼きたて屋"

# 月 -> (①加盟店売上税込, ⑥利用料税込, ⑥の内税, ⑦お戻し, 元ファイル)
DATA = {
    "9月":  (66670, 25637, 2330, 3273, "13608422_202509_payment.pdf"),
    "10月": (43720, 17115, 1555, 5475, "13608422_202510_payment.pdf"),
    "11月": (54700, 21408, 1946, 1071, "13608422_202511_payment.pdf"),
    "12月": (42730, 16612, 1510,  536, "13608422_202512_payment.pdf"),
    "1月":  (27500, 10673,  970, 2202, "13608422_202601_payment.pdf"),
    "2月":  (15530,  6182,  562,    0, "13608422_202602_payment.pdf"),
    "3月":  (42140, 16381, 1489, 7202, "13608422_202603_payment.pdf"),
    "4月":  (20180,  8001,  727,    0, "13608422_202604_payment.pdf"),
    "5月":  (38080, 14739, 1339,    0, "13608422_202605_payment.pdf"),
    "6月":  (22450,  8845,  804, 1131, "13608422_202606_payment.pdf"),
    "7月":  (24530,  9330,  848,    0, "13608422_202607_payment.pdf"),
    "8月":  (34450,  9915,  901,  903, "13608422_202608_payment.pdf"),
}
SRC = "Drive 2.焼きたて屋/取引集計（出前館）/"

# ★2026-09-10 に見つかった【2つめの出前館アカウント】--------------------
#   13575701 ＝ タコとハイボール神栖横丁店（店舗コード TAQV9_0101）。
#   フォルダID 1xKGTn1SwZbDHIfhK9atsk2MRt3GsrDZT（202601〜202608＋2025サブフォルダ）。
#   焼きたて屋（13608422）とは別物で、これまで一度も読んでいなかった。
#
#   ★売上は二重計上しない。エアレジの会計明細に「出前館支払合計額」列があり、
#     タコハイ8月は 64,100 で支払通知書の①と1円まで一致した。つまり
#     タコとハイボールの「売上（税込）」には出前館ぶんが最初から入っている。
#     既存21期PLも手数料だけを計上していて、同じ考え方だった。
#     → このモジュールが入れるのは【手数料だけ】。
#
#   検算（PDFから ⑥−内税 を計算して既存PLと突き合わせ）:
#       7月 19,420−1,765＝17,655  … 既存PL 17,655 と1円まで一致 ✅
#       6月 26,878−2,443＝24,435  … 既存PL 24,430（5円ちがい。既存PLは手入力なので
#                                    丸め違いか打ち間違いとみられる。書類のほうが正）
TACO_TAB = "タコとハイボール"
TACO = {
    "8月": (64100, 18351, 1668, 0, "13575701_202608_payment.pdf"),
}
TACO_SRC = "Drive 3.タコとハイボール/取引集計（出前館）/"
TACO_CHECK = {"7月": (19420, 1765, 17655), "6月": (26878, 2443, 24430)}


def sales_tax(inc):
    """加盟店売上（税込・軽減8%）から内税を抽出。22期から使う。"""
    return int((Decimal(inc) * 8 / 108).to_integral_value(ROUND_FLOOR))


def fee_rows():
    """21期に入れるぶん: 支払手数料（出前館）＝⑥の税抜"""
    for month, (inc, fee, fee_tax, back, src) in DATA.items():
        yield TAB, "支払手数料（出前館）", month, fee - fee_tax, SRC + src, (fee, fee_tax)
    for month, (inc, fee, fee_tax, back, src) in TACO.items():
        yield TACO_TAB, "支払手数料（出前館）", month, fee - fee_tax, TACO_SRC + src, (fee, fee_tax)


# 焼きたて屋の「出前館売上（税込）」「出前館消費税」は9月〜7月は既存PL（sales.py）
# から入る。8月は既存PLに無いので支払通知書から入れる。
#   ★消費税は21期のやり方（①×8%を四捨五入）に揃える。内税抽出（①×8/108）に
#     変えるのは22期から（このファイル冒頭の★を参照）。混ぜると期の中で方法が割れる。
SALES_FROM_PDF = ["8月"]


def sales_rows():
    """(タブ, PL行, 月, 金額, 元ファイル, メモ) — 既存PLに無い月の売上・消費税"""
    for month in SALES_FROM_PDF:
        inc, fee, fee_tax, back, src = DATA[month]
        tax = int((Decimal(inc) * 8 / 100).quantize(Decimal("1"), ROUND_HALF_UP))
        yield (TAB, "出前館売上（税込）", month, inc, SRC + src,
               f"加盟店売上合計① {inc:,}（現金決済＋ネット決済＋ポイント/クーポン）。"
               "既存21期PLに8月は無いので支払通知書から入れた")
        yield (TAB, "出前館消費税", month, tax, SRC + src,
               f"①{inc:,}×8%を四捨五入して {tax:,}。"
               "21期は既存PLと同じやり方に揃えている（22期から①×8/108の内税抽出に変える）")


# 既存PLスプシ（21期）の実測値。本モジュールの計算値と一致することを assert する。
#   file_id 1OB3yLta9y1d14DbN3jeqCoL9MKFknUbBnqR1WBOvLYQ  2026-08-20 に読み取り
EXIST_21_FEE = {          # 支払手数料（出前館）  年計 140,843
    "9月": 23307, "10月": 15560, "11月": 19462, "12月": 15102, "1月": 9703,
    "2月": 5620, "3月": 14892, "4月": 7274, "5月": 13400, "6月": 8041, "7月": 8482,
}
EXIST_21_REFUND = {       # 支払手数料（出前館返金）年計 ▲20,890（費用のマイナス）
    "9月": 3273, "10月": 5475, "11月": 1071, "12月": 536, "1月": 2202,
    "2月": 0, "3月": 7202, "4月": 0, "5月": 0, "6月": 1131, "7月": 0,
}


def refund_rows():
    """お戻し金額⑦ → 支払手数料（出前館返金）にマイナスで計上

    既存PLがこの扱い（費用のマイナス）なので、雑収入ではなくこちらに揃える。
    2026-08-20 に既存スプシを読んで確認済み。以前は保留リストに出していた。
    """
    for month, (inc, fee, fee_tax, back, src) in DATA.items():
        if back:
            yield TAB, "支払手数料（出前館返金）", month, -back, SRC + src, back
    for month, (inc, fee, fee_tax, back, src) in TACO.items():
        if back:
            yield TACO_TAB, "支払手数料（出前館返金）", month, -back, TACO_SRC + src, back


def check(sales_module):
    """PDFから求めた値が既存PLと一致することを確認する（①・手数料・返金）"""
    import build2
    vals = sales_module.SALES[TAB]["出前館売上（税込）"]
    for i, m in enumerate(build2.MONTHS):
        if m in DATA and vals[i] is not None:      # 8月は既存PLに無い
            got, want = DATA[m][0], vals[i]
            assert got == want, f"{m}: 売上① PDF {got:,} ≠ 既存PL {want:,}"
    for tab, _, m, v, _, _ in fee_rows():
        if tab != TAB or m not in EXIST_21_FEE:    # 8月・タコハイは既存PLと比べない
            continue
        assert v == EXIST_21_FEE[m], f"{m}: 手数料 PDF {v:,} ≠ 既存PL {EXIST_21_FEE[m]:,}"
    for tab, _, m, v, _, back in refund_rows():
        if tab != TAB or m not in EXIST_21_REFUND:
            continue
        assert back == EXIST_21_REFUND[m], f"{m}: 返金 PDF {back:,} ≠ 既存PL {EXIST_21_REFUND[m]:,}"
    # タコハイは7月のPDFで方法の裏を取っている（6月は既存PLが5円ちがい・下の★参照）
    fee, tax, want = TACO_CHECK["7月"]
    assert fee - tax == want, f"タコハイ7月の検算がずれた: {fee - tax:,} ≠ {want:,}"
    for tab, plrow, m, v, _s, _n in sales_rows():
        assert plrow in build2.RIDX[tab], f"{tab} に「{plrow}」行が無い"
    return True


if __name__ == "__main__":
    import sales
    check(sales)
    print("① は既存PLの出前館売上（税込）と全月一致 ✅\n")
    tot = 0
    for tab, row, m, v, src, (fee, tax) in fee_rows():
        print(f"  {m:<4} 利用料税込{fee:>7,} − 内税{tax:>6,} → {v:>7,}")
        tot += v
    print(f"  {'計':<4} {'':>20} {tot:>7,}")
    print("\n返金（お戻し金額⑦→ 支払手数料（出前館返金）にマイナス計上）")
    for tab, row, m, v, src, back in refund_rows():
        print(f"  {m:<4} {v:>8,}")
    print(f"  {'計':<4} {sum(v for *_, v, _, _ in refund_rows()):>8,}")
    print("\n手数料・返金とも既存PLと全月一致 ✅")
