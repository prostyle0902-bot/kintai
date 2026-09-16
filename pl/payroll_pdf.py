#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""給料一覧表PDFから組む人件費・法定福利費（9月〜3月 と 8月）

--- なぜ書類から組み直したのか -------------------------------------------
利用者確認 2026-08-27「21期も無視していいよ。税理士が作ってるわけじゃなくて、
俺が作ってたから」。21期PLシートは正ではないと分かったので、
給与も書類（給料一覧表PDF）から組み直す。

    9月〜3月 … このモジュール（PDF）      ★2026-08-27 新規
    4月〜7月 … payroll.py（給与集計スプレッドシート）
    8月     … このモジュール（PDF）

★4月〜7月をここに寄せなかった理由:
  給与集計スプレッドシート（202604〜202607のみ存在）は【人ごと・月ごと】に
  店舗を決めている。こちらの規則は「いまの勤怠アプリの名簿」で決めるので、
  当時と配属が違う人がいる月はズレる（実測で4月3件・5月4件・6月3件ズレた。
  7月は0件で完全一致）。月ごとの人別割り当てがある資料のほうが根拠が強いので、
  その4か月は payroll.py のままにした。
  9月〜3月はその資料が存在しないので、PDF＋名簿がいちばん確かな資料になる。

--- 割り振りの規則 -------------------------------------------------------
kyuyo_split.py にまとめてある（22期の折半ルールも同じ場所）。
    店舗 … 勤怠アプリの名簿（roster.py）。引けない人は部門コード（社員番号の頭4桁）
    行  … 社員番号 0002-xxxx は【人件費（店長）】、それ以外は【人件費（アルバイト）】。
           本部だけ 0001-0001＝人件費　社長、0001-0002＝人件費　純子
    社会保険料 … 【法定福利費】
★この規則で7月を組み直すと payroll.py（給与集計スプシ由来）と完全に一致する。
  check() が毎回それを確かめている。

--- ★2026年1月のPDFが2種類ある --------------------------------------------
    /※プロスタイル給与※/プロスタイル給与R8年/給料一覧表-202601.pdf（2026-01-27）
    /※プロスタイル給与※/プロスタイル給与R7年/給料一覧表-202601.pdf（2026-02-18）
違うのは 0002-0014 小林俊樹（鳥害対策課）1人だけ。
    R8年フォルダ版 総支給174,970／社保45,906
    R7年フォルダ版 総支給251,326／社保46,326  （差 76,356）
★2026-08-31 利用者確認「新しい方で。少ない金額の方」→【R8年フォルダ版】で確定。
  2026年1月は令和8年なので年フォルダとして正しいのもR8年のほう、
  金額が少ないのもR8年フォルダ版（174,970 < 251,326）で、どちらの見方も同じ答えになる。
  （R7年フォルダのほうがDropboxの更新日は新しいが、更新日は再アップや移動でも動くので
  中身の新しさの証拠にはならない。）

--- ★給料一覧表に載らない振込 ---------------------------------------------
りゅうちゃん店長 コウノリュウジ への毎月10日 300,000 は給与ではなく別経路の振込で、
給料一覧表PDFには載らない。銀行明細に実物があるので OUTSIDE で足している。
"""
import collections

import kyuyo_parse
import kyuyo_split

PERIOD = "21期"
# PL列 -> 給料一覧表のファイル名（kyuyo/<ym>.pdf）
MONTHS = {
    "9月": "202509", "10月": "202510", "11月": "202511", "12月": "202512",
    "1月": "202601_R8", "2月": "202602", "3月": "202603",
    # 4月〜7月は payroll.py（給与集計スプレッドシート）が持っている
    "8月": "202608",
}
SRC = "給料一覧表-{ym}.pdf（Dropbox /※プロスタイル給与※/）"

# 給料一覧表には載らないが銀行明細に実物がある人件費
# (タブ, PL行, 金額, 月, 氏名, 元ファイル, メモ)
OUTSIDE = [
    ("りゅうちゃん", "人件費（店長）", 300000,
     ["9月", "10月", "11月", "12月", "1月", "2月", "3月", "8月"],
     "コウノリュウジ",
     "銀行明細/21期/（千葉銀行の店舗口座・PayPay銀行）",
     "毎月10日の振込。給料一覧表には載らない別経路（利用者確認 2026-08-31「入れて」）。"
     "千葉銀行ぶん（9月・10月）は300,330だが330は振込手数料なので本体300,000を計上。"
     "4月〜7月は payroll.py が同額を持っている。"
     "★8月は 2026-09-01 に202608の明細が届いて 2026/08/10 PayPay 300,000 が確認できたので追加"),
]


def _outside_check():
    """OUTSIDE が実際の振込と合っているか、銀行明細で毎回確かめる。"""
    import bank
    got = [(d, a) for d, _bk, raw, _v, a, _s, _k in bank.payments() if "コウノ" in raw]
    assert len(got) == 12, f"コウノリュウジへの振込が{len(got)}件（9月〜8月の12件のはず）"
    for d, a in got:
        assert a in (300000, 300330), f"{d} の振込 {a:,} が300,000でも300,330でもない"


def rows():
    """(タブ, PL行, 月, 金額, メモ) を列挙。給与は消費税の対象外なのでそのまま。"""
    for m, ym in MONTHS.items():
        split, _fb, _pool = kyuyo_split.split(ym, PERIOD)
        emp = kyuyo_parse.parse(ym)[0]
        nm = kyuyo_parse.names(ym)
        who = collections.defaultdict(list)
        for no in sorted(emp):
            tab, _how = kyuyo_split.tab_of(no, nm.get(no, ""))
            who[(tab, kyuyo_split.row_of(no))].append(f"{nm.get(no,'')}({no})")
        for (tab, row), v in sorted(split.items()):
            if row == "法定福利費":
                note = f"給料一覧表{m}分の社会保険料計（{tab}ぶん）"
            else:
                names = who[(tab, row)]
                note = (f"給料一覧表{m}分の総支給額（{len(names)}人）: "
                        + "／".join(names[:12])
                        + (f" ほか{len(names)-12}人" if len(names) > 12 else ""))
            yield tab, row, m, int(v), note, SRC.format(ym=ym.split("_")[0])
    for tab, row, amount, months, who, src, note in OUTSIDE:
        for m in months:
            yield tab, row, m, amount, f"{who}。{note}", src


def fallbacks():
    """名簿で引けず部門コードで決めた人 (月, 社員番号, 氏名, 金額, タブ)。"""
    for m, ym in MONTHS.items():
        for no, name, g, tab in kyuyo_split.split(ym, PERIOD)[1]:
            yield m, no, name, g, tab


def hold_rows():
    """判断が割れるところ。推測で決めずに記録に残す。"""
    fb = list(fallbacks())
    if fb:
        by = collections.defaultdict(list)
        for m, no, name, g, tab in fb:
            by[(no, name, tab)].append(m)
        for (no, name, tab), ms in sorted(by.items()):
            yield ("・".join(ms), tab, f"{name}（{no}）",
                   f"勤怠アプリの名簿に無いので部門コード{no[:4]}で{tab}に入れた。"
                   "退職者などで名簿から消えている人。"
                   "★2026-08-31 利用者確認「配属OK」＝この割り振りで正しい（記録として残す）")


def check(wb=None):
    """PDFの総合計と一致するか／7月を組み直して payroll.py と合うか。"""
    for m, ym in MONTHS.items():
        tot = kyuyo_parse.parse(ym)[1]
        split = kyuyo_split.split(ym, PERIOD)[0]
        g = sum(v for (t, r), v in split.items() if r != "法定福利費")
        s = sum(v for (t, r), v in split.items() if r == "法定福利費")
        assert g == tot[0], f"{m}: 割り振った総支給 {g:,} ≠ PDFの総合計 {tot[0]:,}"
        assert s == tot[1], f"{m}: 割り振った社会保険 {s:,} ≠ PDFの総合計 {tot[1]:,}"
    _outside_check()
    # ★同じ規則で7月を組み直し、給与集計スプシ由来の payroll.py と突き合わせる。
    import payroll
    want = {(t, r): v[3] for (t, r), v in payroll.PAYROLL.items()}
    got = kyuyo_split.split("202607", PERIOD)[0]
    ALLOW = {("りゅうちゃん", "人件費（店長）"),           # コウノ（別経路）
             ("タコとハイボール", "人件費（アルバイト）"),   # 松岡実麻
             ("韓国酒場ハナ", "人件費（アルバイト）"),       # 同上
             ("さわら十三里屋", "法定福利費"),
             ("業務課", "法定福利費")}
    bad = [(k, got.get(k, 0), want.get(k, 0)) for k in set(want) | set(got)
           if k not in ALLOW and got.get(k, 0) != want.get(k, 0)]
    assert not bad, f"7月を組み直すと給与集計スプシと合わない: {bad}"
    if wb is None:
        return
    import build2
    for tab, row, m, v, _note, _src in rows():
        assert row in build2.RIDX[tab], f"{tab} に「{row}」行が無い"
        c = wb[tab][f"{build2.MCOL[m]}{build2.RIDX[tab][row]}"]
        assert not c.value, f"{tab} {row} {m} に既に {c.value} が入っている"


if __name__ == "__main__":
    check()
    agg = collections.defaultdict(dict)
    for tab, row, m, v, _n, _s in rows():
        agg[(tab, row)][m] = v
    ms = list(MONTHS)
    print(f"{'タブ':<12}{'行':<20}" + "".join(f"{m:>10}" for m in ms) + f"{'計':>12}")
    tot = 0
    for k in sorted(agg):
        v = agg[k]; tot += sum(v.values())
        print(f"{k[0]:<12}{k[1]:<20}" + "".join(f"{v.get(m,0):>10,}" for m in ms)
              + f"{sum(v.values()):>12,}")
    print(f"\n計 {tot:,}円 ／ {len(list(rows()))}セル")
    print("\n保留:")
    for m, tab, item, why in hold_rows():
        print(f"   [{m}] {tab} {item}")
