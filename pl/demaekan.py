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
from decimal import Decimal, ROUND_FLOOR

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
}
SRC = "Drive 2.焼きたて屋/取引集計（出前館）/"


def sales_tax(inc):
    """加盟店売上（税込・軽減8%）から内税を抽出。22期から使う。"""
    return int((Decimal(inc) * 8 / 108).to_integral_value(ROUND_FLOOR))


def fee_rows():
    """21期に入れるぶん: 支払手数料（出前館）＝⑥の税抜"""
    for month, (inc, fee, fee_tax, back, src) in DATA.items():
        yield TAB, "支払手数料（出前館）", month, fee - fee_tax, SRC + src, (fee, fee_tax)


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


def check(sales_module):
    """PDFから求めた値が既存PLと一致することを確認する（①・手数料・返金）"""
    import build2
    vals = sales_module.SALES[TAB]["出前館売上（税込）"]
    for i, m in enumerate(build2.MONTHS):
        if m in DATA:
            got, want = DATA[m][0], vals[i]
            assert got == want, f"{m}: 売上① PDF {got:,} ≠ 既存PL {want:,}"
    for _, _, m, v, _, _ in fee_rows():
        assert v == EXIST_21_FEE[m], f"{m}: 手数料 PDF {v:,} ≠ 既存PL {EXIST_21_FEE[m]:,}"
    for _, _, m, v, _, back in refund_rows():
        assert back == EXIST_21_REFUND[m], f"{m}: 返金 PDF {back:,} ≠ 既存PL {EXIST_21_REFUND[m]:,}"
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
