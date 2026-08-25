#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""8月の給与 → 各タブの人件費・法定福利費（給料一覧表PDFから直に組む）

--- なぜ別モジュールなのか -----------------------------------------------
payroll.py は4月〜7月ぶんを「給与集計_月次累積_202607」スプレッドシートから
写している。8月分のスプレッドシートは無い。
利用者指示（2026-08-23）「8月は基本、CSVやPDFから全部数字入れてください。
既存は無視してくださいね」に沿って、給料一覧表PDFだけで組む。
    Dropbox /※プロスタイル給与※/プロスタイル給与R8年/給料一覧表-202608.pdf
    ローカル kyuyo/202608.pdf   読み取りは kyuyo_parse.py

--- 割り振りの規則（2026-08-25 に7月で検証した）--------------------------
① 店舗 … 勤怠アプリの名簿（roster.py）で氏名から引く。
   名簿に無い人は【部門コード】（社員番号の頭4桁）で引く。部門コードは
   給料一覧表そのものが持っている区分なので、書類だけで完結する。
② 行  … 社員番号 0002-xxxx は【人件費（店長）】、それ以外は【人件費（アルバイト）】。
   本部だけ例外で、0001-0001＝人件費　社長、0001-0002＝人件費　純子。
③ 社会保険料 … そのまま【法定福利費】。

★この規則で7月を組み直すと、会計士の既存21期PLと次の3点以外は完全に一致した。
  （検算は check() が毎回やっている）
    ・りゅうちゃん「人件費（店長）」300,000
      … コウノリュウジ。給料一覧表に載っていない別経路の振込。下の HOLD 参照
    ・本部の 805,000
      … 会計士は社長700,000／純子105,000に分けている。②の例外で合わせた
    ・松岡実麻 16,110
      … 部門は0091（神栖横丁）、名簿はタコとハイボール、会計士は7月に
        韓国酒場ハナへ入れていた。3つとも食い違う。下の HOLD 参照

--- 8月のPDFの中身 -------------------------------------------------------
    56人 ／ 総支給額 7,317,242 ／ 社会保険計 618,404
名簿で引けなかったのは2人だけで、どちらも部門コードで決まった。
    0081-0093 茂木多美子 50,202 … 部門0081（他の2人とも名簿は焼きたて屋）
    0091-0101 齋藤良江    9,722 … 部門0091（他の6人とも名簿は神栖横丁）
"""
import collections

import kyuyo_parse
import roster

YM = "202608"
MONTH = "8月"
SRC = "給料一覧表-202608.pdf（Dropbox /※プロスタイル給与※/プロスタイル給与R8年/）"

# 部門コード（社員番号の頭4桁）→ タブ。名簿で引けない人の受け皿。
# ★7月・8月のデータで「その部門の人が名簿でどこに割れているか」を数えて決めた。
BUMON = {
    "0001": "本部", "0002": "業務課", "0011": "業務課", "0021": "業務課",
    "0022": "業務課", "0031": "業務課", "0071": "業務課", "0081": "焼きたて屋",
    "0091": "神栖横丁", "0092": "りゅうちゃん", "0101": "さわら十三里屋",
    "0131": "業務課", "0141": "業務課",
}

# 本部だけ行が人ごとに分かれている
HONBU_ROW = {"0001-0001": "人件費　社長", "0001-0002": "人件費　純子"}


def _tab(no, name):
    """氏名（勤怠アプリの名簿）→ タブ。引けなければ部門コードで引く。"""
    t = roster.tab_of(name)
    if t:
        return t, "名簿"
    return BUMON.get(no[:4]), "部門コード"


def _split(ym):
    """{(タブ, PL行): 金額} と、名簿で引けなかった人。"""
    emp = kyuyo_parse.parse(ym)[0]
    nm = kyuyo_parse.names(ym)
    out = collections.Counter()
    fallback = []
    for no in sorted(emp):
        gross, ins = emp[no]
        name = nm.get(no, "")
        tab, how = _tab(no, name)
        assert tab, f"{no} {name}: 店舗が決められない（名簿にも部門コード表にも無い）"
        if how == "部門コード":
            fallback.append((no, name, gross, tab))
        row = HONBU_ROW.get(no) or (
            "人件費（店長）" if no.startswith("0002") else "人件費（アルバイト）")
        out[(tab, row)] += gross
        if ins:
            out[(tab, "法定福利費")] += ins
    return out, fallback


def rows():
    """(タブ, PL行, 月, 金額, メモ) を列挙。給与は消費税の対象外なのでそのまま。"""
    split, _fb = _split(YM)
    emp = kyuyo_parse.parse(YM)[0]
    nm = kyuyo_parse.names(YM)
    who = collections.defaultdict(list)
    for no in sorted(emp):
        tab, _how = _tab(no, nm.get(no, ""))
        row = HONBU_ROW.get(no) or (
            "人件費（店長）" if no.startswith("0002") else "人件費（アルバイト）")
        who[(tab, row)].append(f"{nm.get(no,'')}({no})")
    for (tab, row), v in sorted(split.items()):
        if row == "法定福利費":
            note = f"給料一覧表{MONTH}分の社会保険料計（{tab}ぶん）"
        else:
            names = who[(tab, row)]
            note = (f"給料一覧表{MONTH}分の総支給額（{len(names)}人）: "
                    + "／".join(names[:12])
                    + (f" ほか{len(names)-12}人" if len(names) > 12 else ""))
        yield tab, row, MONTH, int(v), note


def hold_rows():
    """判断が割れるところ。推測で決めずに記録に残す。"""
    return [
        (MONTH, "りゅうちゃん", "人件費（店長）",
         "9月〜7月は毎月300,000（コウノリュウジ）が入っているが、これは給料一覧表に"
         "載らない別経路の振込で、8月分の銀行明細（PayPay・千葉銀行の202608）が"
         "まだ手元に無い。書類で確認できないので8月は入れていない。"
         "明細が届いたら payroll8.py に足すこと"),
        (MONTH, "タコとハイボール", "人件費（アルバイト）（松岡実麻 21,480）",
         "この1人だけ3つの資料が食い違う。給料一覧表の部門コードは0091（神栖横丁）、"
         "勤怠アプリの名簿はタコとハイボール、会計士の既存21期PLは7月に"
         "韓国酒場ハナへ入れていた（16,110）。名簿を優先してタコとハイボールに"
         "入れてある（4〜7月と同じ扱い）。正しい店舗を教えてもらえれば直す"),
    ]


def check(wb=None):
    """PDFの総合計と一致するか／7月を組み直して既存PLと合うか。"""
    emp, tot = kyuyo_parse.parse(YM)
    split, fallback = _split(YM)
    g = sum(v for (t, r), v in split.items() if r != "法定福利費")
    s = sum(v for (t, r), v in split.items() if r == "法定福利費")
    assert g == tot[0], f"割り振った総支給 {g:,} ≠ PDFの総合計 {tot[0]:,}"
    assert s == tot[1], f"割り振った社会保険 {s:,} ≠ PDFの総合計 {tot[1]:,}"
    # ★同じ規則で7月を組み直し、会計士の既存21期PLと突き合わせる。
    #   食い違ってよいのは docstring に書いた3点だけ。
    import payroll
    want = {(t, r): v[3] for (t, r), v in payroll.PAYROLL.items()}
    got, _ = _split("202607")
    ALLOW = {("りゅうちゃん", "人件費（店長）"),          # コウノ（別経路）
             ("タコとハイボール", "人件費（アルバイト）"),  # 松岡実麻
             ("韓国酒場ハナ", "人件費（アルバイト）"),      # 同上
             ("さわら十三里屋", "法定福利費"),            # 会計士のみ2,429
             ("業務課", "法定福利費")}                   # 会計士のみ+1,100
    bad = [(k, got.get(k, 0), want.get(k, 0)) for k in set(want) | set(got)
           if k not in ALLOW and got.get(k, 0) != want.get(k, 0)]
    assert not bad, f"7月を組み直すと既存PLと合わない: {bad}"
    if wb is None:
        return fallback
    import build2
    for tab, row, m, v, _note in rows():
        assert row in build2.RIDX[tab], f"{tab} に「{row}」行が無い"
        c = wb[tab][f"{build2.MCOL[m]}{build2.RIDX[tab][row]}"]
        assert not c.value, f"{tab} {row} {m} に既に {c.value} が入っている"
    return fallback


if __name__ == "__main__":
    fb = check()
    print(f"給料一覧表{MONTH}分（{SRC}）\n")
    tot = 0
    for tab, row, m, v, note in rows():
        print(f"{tab:<14}{row:<22}{v:>10,}")
        tot += v
    print(f"{'計':<37}{tot:>10,}")
    print(f"\n名簿で引けず部門コードで決めた人 {len(fb)}人")
    for no, name, g, tab in fb:
        print(f"   {no} {name} {g:,} → {tab}（部門{no[:4]}）")
    print("\n保留:")
    for m, tab, item, why in hold_rows():
        print(f"   [{m}] {tab} {item}")
    print("\n検算すべて通過")
