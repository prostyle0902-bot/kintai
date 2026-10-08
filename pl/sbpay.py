#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""さわら十三里屋のPayPay決済端末（SBペイメント）の手数料 → 22期 十三里屋「支払手数料（Softbank）」

--- 利用者指示 2026-10-08 ---------------------------------------------------
「十三里屋のPayPay決済端末（SBペイメント）の収納明細書を入れるフォルダ
  『08十三里屋カード決済（PayPay端末）』と、保管先『十三里屋カード決済明細/21期・22期』を
  Dropboxに作りました。定期チェックの対象に加えて、取り込んだら保管先へ移すようにしてください。
  損益計算書の十三里屋タブに、焼きたて屋と同じ『支払手数料（Softbank）』の行を作り、
  22期9月に 15,748円（税抜、半月2枚の合計）を入れてください。」

--- 元データ -----------------------------------------------------------------
Dropbox /※請求書※/★毎月ここに入れる/08_十三里屋カード決済（PayPay端末）/
    【Ｐｒｏｓｔｙｌｅ株式会社（J9787-001）さわら十三里屋】<YYYYMM>_収納明細書<MMDD-MMDD>_<番号>.pdf
取り込んだら /※請求書※/十三里屋カード決済明細/<期>/ へ移す（ingest.py）。

半月ごとに1枚。ファイル名の <YYYYMM> は【発行月】で、集計期間の月とは限らない
（9月後半＝0916-0930 は 202610_ で届く）。★月は PDF の「集計期間」で決める。

--- 金額 ---------------------------------------------------------------------
合計行の「手数料」（税抜）を取る。半月2枚を足してその月の値にする。
焼きたて屋（kessai.py の 支払手数料（Softbank））と同じ取り方。
★クレジットカード VISA/Master の手数料は非課税、そのほかは外税10%。
  合計行の「税込金額」−「手数料」を消費税として持っておく（PLには税抜だけ入る）。

--- なぜPDFを持たないか ------------------------------------------------------
pl/ には置かない（rikuji_pdf・yokocho_pdf と同じ。.gitignore 済み）。
読んだ数字だけを下の DATA に書く。新しい明細が届いたら
    python3 sbpay22.py <PDF> …
で合計行を読み、出てきた1行を DATA に足す。
"""
import re
import sys

TAB = "さわら十三里屋"
ROW = "支払手数料（Softbank）"
SRC = "Dropbox /※請求書※/十三里屋カード決済明細/22期/"

# (集計開始, 集計終了, 手数料（税抜）, 税込金額, 決済処理金額, ファイル名の頭)
DATA = [
    ("2026-09-01", "2026-09-15", 5009, 5366, 174610,
     "202609_収納明細書0901-0915_20260921776393"),
    ("2026-09-16", "2026-09-30", 10739, 11644, 304270,
     "202610_収納明細書0916-0930_20261022014861"),
]

PERIOD22 = ("2026-09", "2027-08")


def month_of(day):
    return f"{int(day[5:7])}月"


def rows():
    """(タブ, PL行, 月, 税抜, 消費税, 取引先, 出どころ, メモ)。月ごとに半月ぶんを足す。"""
    by = {}
    for start, end, fee, inc, total, name in DATA:
        assert start[:7] == end[:7], f"{name}: 集計期間が月をまたいでいる"
        if not (PERIOD22[0] <= start[:7] <= PERIOD22[1]):
            continue
        m = month_of(start)
        b = by.setdefault(m, {"fee": 0, "inc": 0, "parts": []})
        b["fee"] += fee
        b["inc"] += inc
        b["parts"].append(f"{start[5:]}〜{end[5:]} 手数料{fee:,}（決済{total:,}）")
    for m, b in by.items():
        yield (TAB, ROW, m, b["fee"], b["inc"] - b["fee"], "SBペイメントサービス",
               SRC, "収納明細書の合計行の手数料（税抜）。" + "＋".join(b["parts"]))


def parse(path):
    """収納明細書PDFの合計行を読んで DATA の1行を返す。"""
    import pymupdf
    t = "\n".join(p.get_text() for p in pymupdf.open(path))
    m = re.search(r"集計期間：(\d{4})/(\d{2})/(\d{2})～(\d{4})/(\d{2})/(\d{2})", t)
    assert m, f"{path}: 集計期間が読めない"
    start = f"{m[1]}-{m[2]}-{m[3]}"
    end = f"{m[4]}-{m[5]}-{m[6]}"
    # 合計行: 件数 / 売上金額 / 返金件数 / 返金金額 / 合計件数 / 合計金額 / 手数料 / 税込金額 / お振込金額
    i = t.index("\n合計\n")
    v = [x for x in t[i:].split("\n")[2:11]]
    num = lambda s: int(s.replace("¥", "").replace(",", ""))
    total, fee, inc, pay = num(v[5]), num(v[6]), num(v[7]), num(v[8])
    assert total - inc == pay, f"{path}: 合計{total:,}−手数料税込{inc:,}≠振込{pay:,}"
    name = re.search(r"(\d{6}_収納明細書\d{4}-\d{4}_\d+)", path)
    return (start, end, fee, inc, total, name[1] if name else path)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        for p in sys.argv[1:]:
            print(parse(p))
    else:
        for r in rows():
            print(r[:5])
