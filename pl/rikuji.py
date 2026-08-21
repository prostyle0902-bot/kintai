#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""陸事総合協同組合（ETC高速代）→ 業務課「陸自総合（高速代）」

--- 元データ -------------------------------------------------------------
請求書PDF  0000042400_高速_請求書_YYYYMM.pdf   ← 【正】。課税/消費税/不課税が明記
CP請求鑑   0000042400_高速_CP請求鑑_YYYYMM.csv  CP932。ETCカード別の内訳。検算用
Dropbox    /※請求書※/買掛/21期/YYMM月/…/業務/
ローカル    rikuji_pdf/YYYYMM.pdf ／ rikuji/CP請求鑑_<PL月>.csv

★★ ファイル名の年月は【利用月】。そのままPL列になる ★★
    引落は利用月の2か月後。
        202509 → 9月列 →「【2025年09月利用分】」→ 振替日 2025/11/04
        202607 → 7月列 →「【2026年07月利用分】」→ 振替日 2026/09/03
    ※ENEOS（eneos.py）はファイル名が【引落月】。逆なので間違えないこと。

--- 請求書PDFの構造 -----------------------------------------------------
1ページ目に全部載っている。テキスト層あり（pdftotext -layout）。
    【2025年09月利用分】
    ご請求金額  \126,143
    【請求金額内訳】
      課税対象額（10%）   消費税額（10%）   不課税・非課税対象額
        \114,676           \11,467             \0
    【当月分ご請求内訳】
      コーポレートカード 126,143 ／ マイレージカード 0 ／ 年会費 0 ／ 調整額 0
PLに入れるのは【課税対象額】。

--- なぜPDFが要るのか（2026-08-20 → 21 の経緯）--------------------------
最初はCP請求鑑CSVだけで組もうとした。CSVは
    ＪＨ請求総額 ＋ 本四調整額 ＋ 首都請求総額 ＋ 阪神請求総額
を全カード合計すれば請求書の総額になる……はずが、11か月中2か月ズレた。
    2月  CSV 111,980 ／ 請求書 112,639  （差 +659）
    4月  CSV  92,193 ／ 請求書  98,124  （差 +5,931）
千葉銀行の実引落は請求書のほうと一致した。CSVはETC利用分しか写しておらず、
月によっては請求書に別の費目が乗る。だからCSVは正にできない。
利用者が11か月ぶんのPDFを揃えてくれたので、PDFを正とする（2026-08-21）。

不課税・非課税は 2月 30円／4月 270円 の2件だけ（年計300円）。
利用者判断「不課税分はいいや」（2026-08-20）によりPLには入れない。

--- ★既存PLの誤り ------------------------------------------------------
既存の業務課シート 陸自総合（高速代）年計 975,862 の内訳:
    9月・12月・1月・4月・5月 … 課税対象額。正しい
    11月・2月・3月           … 税込のまま入っている
    10月 143,552             … 2025/10/03の引落額。これは20期8月の利用分。
                                21期10月の利用分は 133,727（引落2025/12/03）
    6月・7月                 … 空欄
このモジュールは全部を正しい利用月・正しい課税対象額で入れ替える。
"""
import csv, os, re, subprocess

BASE = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(BASE, "rikuji_pdf")
CSV_DIR = os.path.join(BASE, "rikuji")
TAB, ROW = "業務課", "陸自総合（高速代）"
MONTHS = ["9月", "10月", "11月", "12月", "1月", "2月", "3月",
          "4月", "5月", "6月", "7月"]
# PL列 -> 請求書ファイル名の年月（＝利用月）
YM = {"9月": "202509", "10月": "202510", "11月": "202511", "12月": "202512",
      "1月": "202601", "2月": "202602", "3月": "202603", "4月": "202604",
      "5月": "202605", "6月": "202606", "7月": "202607"}

# 請求書PDFから読んだ値（2026-08-21）。PDFが手元に無くても動くように焼いてある。
#   PL月: (課税対象額, 消費税額, 不課税・非課税, ご請求金額, 振替日)
PDF_DATA = {
    "9月":  (114676, 11467,   0, 126143, "2025/11/04"),
    "10月": (121570, 12157,   0, 133727, "2025/12/03"),
    "11月": (130458, 13045,   0, 143503, "2026/01/05"),
    "12月": (133058, 13305,   0, 146363, "2026/02/03"),
    "1月":  ( 64296,  6429,   0,  70725, "2026/03/03"),
    "2月":  (102372, 10237,  30, 112639, "2026/04/03"),
    "3月":  ( 90366,  9036,   0,  99402, "2026/05/07"),
    "4月":  ( 88959,  8895, 270,  98124, "2026/06/03"),
    "5月":  ( 75777,  7577,   0,  83354, "2026/07/03"),
    "6月":  ( 97461,  9746,   0, 107207, "2026/08/03"),
    "7月":  (126187, 12618,   0, 138805, "2026/09/03"),
}

# 千葉銀行の実引落（bank.py「ＲＫＳ．リクジソウゴウ」）。21期の明細は〜2026/07。
BANK = {"9月": 126143, "10月": 133727, "11月": 143503, "12月": 146363,
        "1月": 70725, "2月": 112639, "3月": 99402, "4月": 98124,
        "5月": 83354, "6月": None, "7月": None}

# 既存PLの値（2026-08-20 読み取り）。入れ替えるので検算には使わない。参考。
EXIST_21 = {"9月": 114676, "10月": 143552, "11月": 143503, "12月": 133058,
            "1月": 64296, "2月": 112639, "3月": 99402, "4月": 88959,
            "5月": 75777, "6月": None, "7月": None}

_YEN = re.compile(r"\\([\d,]+)")


def _parse_pdf(month):
    """請求書PDFの1ページ目から (課税, 消費税, 不課税, 請求金額, 振替日, 利用月表記)。
    PDFが無ければ None。"""
    path = os.path.join(PDF_DIR, f"{YM[month]}.pdf")
    if not os.path.exists(path):
        return None
    txt = subprocess.run(["pdftotext", "-layout", "-f", "1", "-l", "1", path, "-"],
                         capture_output=True).stdout.decode("utf-8", "replace")
    lines = txt.splitlines()
    i = next(n for n, s in enumerate(lines) if "請求金額内訳" in s)
    tax = [int(x.replace(",", "")) for x in _YEN.findall("\n".join(lines[i:i + 4]))]
    assert len(tax) == 3, f"{path}: 請求金額内訳が3つ取れない → {tax}"
    total = int(re.search(r"コーポレートカード\s+([\d,]+)", txt).group(1).replace(",", ""))
    payday = re.search(r"振替日\s+([\d/]+)", txt).group(1)
    used = re.search(r"【(\d{4})年(\d{2})月利用分】", txt)
    return (*tax, total, payday, f"{used.group(1)}{used.group(2)}")


def _csv_total(month):
    """CP請求鑑CSVのETC高速代合計（税込）と、使用のあったカード枚数。検算用。"""
    path = os.path.join(CSV_DIR, f"CP請求鑑_{month}.csv")
    if not os.path.exists(path):
        return None, 0
    rows = list(csv.reader(open(path, "rb").read().decode("cp932").splitlines()))
    h = {k: i for i, k in enumerate(rows[0])}
    cols = ["ＪＨ請求総額", "本四調整額", "首都請求総額", "阪神請求総額"]
    total = cards = 0
    for r in rows[1:]:
        if not r or not r[0]:
            continue
        total += sum(int(r[h[c]] or 0) for c in cols)
        if int(r[h["利用総額"]] or 0):
            cards += 1
    return total, cards


def rows():
    """(タブ, PL行, 月, 税抜, 元ファイル, メモ) を列挙。入れるのは課税対象額。"""
    for m in MONTHS:
        ex, tax, free, total, payday = PDF_DATA[m]
        yy = "25" if m in ("9月", "10月", "11月", "12月") else "26"
        note = (f"請求書の課税対象額。請求総額{total:,}＝課税{ex:,}＋消費税{tax:,}"
                + (f"＋不課税{free:,}" if free else "") + f"。振替日 {payday}")
        if free:
            note += f"／不課税{free:,}円は計上しない（利用者判断）"
        yield (TAB, ROW, m, ex,
               f"買掛/21期/{yy}{int(m[:-1]):02d}月/…/業務/"
               f"0000042400_高速_請求書_{YM[m]}.pdf", note)


def hold_rows():
    """CP請求鑑CSVと請求書が合わない月を、参考として出す（金額はPDFが正）"""
    for m in MONTHS:
        c, _ = _csv_total(m)
        total = PDF_DATA[m][3]
        if c is not None and c != total:
            yield (m, TAB, ROW,
                   f"CP請求鑑CSVの合計 {c:,} が請求書 {total:,} と {total-c:+,} 違う。"
                   f"CSVはETC利用分しか写らないため。金額は請求書PDFを採用済みで、"
                   f"計上に問題は無い。CSVだけで組むと足りなくなる、という記録。")


def check():
    """PDFの実物・焼いた表・CSV・銀行明細、4つを突き合わせる"""
    for m in MONTHS:
        ex, tax, free, total, payday = PDF_DATA[m]
        assert ex + tax + free == total, \
            f"{m}: 課税{ex:,}＋消費税{tax:,}＋不課税{free:,} ≠ 請求{total:,}"
        assert tax == ex // 10, f"{m}: 消費税{tax:,} ≠ 課税{ex:,}の10%"
        got = _parse_pdf(m)
        if got:                      # PDFが手元にあれば実物と照合
            assert got[:5] == (ex, tax, free, total, payday), \
                f"{m}: PDF実物 {got[:5]} ≠ 焼いた値 {(ex, tax, free, total, payday)}"
            assert got[5] == YM[m], \
                f"{m}: PDFの利用月 {got[5]} がファイル名 {YM[m]} と違う"
        if BANK[m]:                  # 銀行の実引落と照合
            assert BANK[m] == total, \
                f"{m}: 請求{total:,} ≠ 千葉銀行の引落{BANK[m]:,}"
        c, cards = _csv_total(m)
        if c is not None:
            assert cards > 0, f"{m}: CP請求鑑の使用カードが0枚"
    return True


if __name__ == "__main__":
    check()
    have = sum(1 for m in MONTHS if _parse_pdf(m))
    print(f"{TAB} / {ROW}")
    print(f"請求書PDF {have}/11本と照合、全項目一致 ✅\n")
    print(f"{'利用月':<6}{'請求金額':>10}{'課税':>10}{'消費税':>9}{'不課税':>7}"
          f"{'CSV':>10}   既存PL")
    print("-" * 74)
    t = 0
    for tab, row, m, ex, src, note in rows():
        _, tax, free, total, _ = PDF_DATA[m]
        c = _csv_total(m)[0]
        cs = f"{c:,}" if c == total else f"{c:,}★"
        e = EXIST_21[m]
        mark = "空欄" if e is None else \
            (f"{e:,}" if e == ex else (f"{e:,}（税込）" if e == total else f"{e:,}（?）"))
        print(f"{m:<6}{total:>10,}{ex:>10,}{tax:>9,}{free:>7,}{cs:>10}   {mark}")
        t += ex
    print("-" * 74)
    print(f"{'計':<6}{'':>10}{t:>10,}")
    old = sum(v for v in EXIST_21.values() if v)
    print(f"\n既存PL年計 {old:,} → 正しい課税対象額 {t:,}（差 {t-old:+,}）")
    print("★＝CP請求鑑CSVが請求書と合わない月\n")
    for m, tab, row, why in hold_rows():
        print(f"  {m}: {why}")
