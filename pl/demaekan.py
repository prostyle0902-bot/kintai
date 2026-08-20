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
21期（利用者判断 B-2 / 2026-08-20）:
    支払手数料（出前館） ← ⑥の税抜（＝⑥ − 内税）  ★これだけ入れる
    出前館売上（税込）・出前館消費税は既存のまま触らない

22期以降:
    出前館売上（税込）  ← ①
    出前館消費税        ← ①から内税抽出（8/108・軽減税率）
    支払手数料（出前館） ← ⑥の税抜

★なぜ21期は手数料だけか
  ①は全月とも既存PLと一致していて直す必要がない。
  消費税は既存が「①×8%を四捨五入」で内税抽出になっていないが、
  これは計算方法の違いなので22期から揃える（yakitateya.py の A-2 と同じ考え方）。
  手数料は方法の違いではなく単純な計上漏れなので、21期にも入れる。

★お戻し金額⑦は保留
  21期で合計20,890円ある（商品代金補填・不課税）。
  PLに「支払手数料（出前館返金）」行はあるが、費用のマイナスなのか
  雑収入なのか判断できないため入れていない。保留リストに出す。
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


def hold_rows():
    """保留: お戻し金額⑦（扱いが未定）"""
    for month, (inc, fee, fee_tax, back, src) in DATA.items():
        if back:
            yield (month, back, SRC + src)


def check(sales_module):
    """①が既存PLの出前館売上（税込）と一致することを確認する"""
    import build2
    vals = sales_module.SALES[TAB]["出前館売上（税込）"]
    for i, m in enumerate(build2.MONTHS):
        if m in DATA:
            got, want = DATA[m][0], vals[i]
            assert got == want, f"{m}: PDF {got:,} ≠ 既存PL {want:,}"
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
    print(f"\n保留（お戻し金額⑦）: {sum(b for _, b, _ in hold_rows()):,}円 / "
          f"{len(list(hold_rows()))}か月")
