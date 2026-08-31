#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""給与を店舗に割り振る規則（22期の折半ルールを含む）

--- 22期の折半（利用者指示 2026-08-27）----------------------------------
「22期のPLだけど、ハナとタコハイの人件費、レナ、ナディシャ、ハルシャニの給与を
  3人分足して、半分ずつにしてほしい。」

3人は給料一覧表と勤怠アプリの名簿から特定した:
    0002-0019 宮川玲奈           ＝ レナ      （名簿では韓国酒場ハナ）
    0002-0020 ゴラカウェラパティ  ＝ ハルシャニ（名簿ではタコとハイボール）
    0002-0021 ウィジェシンへアラ  ＝ ナディシャ（名簿ではタコとハイボール）
※カタカナ名は給料一覧表が姓、名簿が名で載っている。roster.ALIAS がつないでいる。

21期までは名簿どおり（レナ＝ハナ、ほか2人＝タコハイ）。
22期からは3人の総支給を合算して、韓国酒場ハナとタコとハイボールで折半する。
3人とも社員番号 0002-xxxx なので行は【人件費（店長）】。

★端数（合計が奇数のとき）は韓国酒場ハナ側に寄せる。
★【法定福利費】も折半する（追加指示 2026-08-27「法定福利費も折半で」）。
  同じ3人の社会保険料を合算して半分ずつ。端数の寄せ先も総支給と同じ。

--- 割り振りの基本規則（21期・22期共通）--------------------------------
店舗 … 勤怠アプリの名簿（roster.py）。引けない人は部門コード（社員番号の頭4桁）
行  … 社員番号 0002-xxxx は【人件費（店長）】、それ以外は【人件費（アルバイト）】。
       本部だけ 0001-0001＝人件費　社長、0001-0002＝人件費　純子
社会保険料 … 【法定福利費】
★この規則で21期7月を組み直すと会計士の既存PLと一致する（payroll8.check）。
"""
import collections

import kyuyo_parse
import roster

# 部門コード（社員番号の頭4桁）→ タブ。名簿で引けない人の受け皿。
BUMON = {
    "0001": "本部", "0002": "業務課", "0011": "業務課", "0021": "業務課",
    "0022": "業務課", "0031": "業務課", "0071": "業務課", "0081": "焼きたて屋",
    "0091": "神栖横丁", "0092": "りゅうちゃん", "0101": "さわら十三里屋",
    # ★0111 は 2026-08-27 追加。9月〜3月にだけ出てくる部門で、
    #   この部門の人のうち名簿に載っている人は業務課だった（奥田京子・小倉孔子は未掲載）。
    "0111": "業務課",
    "0131": "業務課", "0141": "業務課",
}

# 本部だけ行が人ごとに分かれている
HONBU_ROW = {"0001-0001": "人件費　社長", "0001-0002": "人件費　純子"}

# 何人かをまとめて複数の店舗で分けるもの。期ごとに効かせる。
POOL = [
    {
        "期": "22期",
        "社員番号": ["0002-0019", "0002-0020", "0002-0021"],
        "名前": "レナ（宮川玲奈）／ナディシャ（ウィジェシンへアラ）／"
                "ハルシャニ（ゴラカウェラパティ）",
        "分ける先": ["韓国酒場ハナ", "タコとハイボール"],
        "行": "人件費（店長）",
        # ★2026-08-27 追加指示「法定福利費も折半で」。社会保険料も同じ3人ぶんを
        #   合算して半分ずつにする。総支給と同じく端数は先頭の店へ。
        "対象": ["総支給", "社会保険"],
        "指示": "利用者指示 2026-08-27「ハナとタコハイの人件費、レナ、ナディシャ、"
                "ハルシャニの給与を3人分足して、半分ずつにしてほしい」",
    },
]


def pools(period):
    return [p for p in POOL if p["期"] == period]


def tab_of(no, name):
    """氏名（名簿）→ タブ。引けなければ部門コードで引く。(タブ, 決め方)"""
    t = roster.tab_of(name)
    if t:
        return t, "名簿"
    return BUMON.get(no[:4]), "部門コード"


def row_of(no):
    return HONBU_ROW.get(no) or (
        "人件費（店長）" if no.startswith("0002") else "人件費（アルバイト）")


def split(ym, period="21期"):
    """{(タブ, PL行): 金額}、名簿で引けなかった人、折半の内訳 を返す。"""
    emp = kyuyo_parse.parse(ym)[0]
    nm = kyuyo_parse.names(ym)
    pooled = {no: p for p in pools(period) for no in p["社員番号"]}
    out = collections.Counter()
    fallback, pool_detail = [], []
    hold = collections.defaultdict(lambda: collections.Counter())

    for no in sorted(emp):
        gross, ins = emp[no]
        name = nm.get(no, "")
        tab, how = tab_of(no, name)
        assert tab, f"{no} {name}: 店舗が決められない（名簿にも部門コード表にも無い）"
        if how == "部門コード":
            fallback.append((no, name, gross, tab))
        p = pooled.get(no)
        if p and "総支給" in p["対象"]:
            hold[id(p)]["総支給"] += gross
            pool_detail.append((no, name, gross))
        else:
            out[(tab, row_of(no))] += gross
        if ins:
            if p and "社会保険" in p["対象"]:
                hold[id(p)]["社会保険"] += ins
            else:
                out[(tab, "法定福利費")] += ins

    # まとめたぶんを分ける。端数は先に書いた店へ寄せる。
    for p in pools(period):
        h = hold.get(id(p))
        if not h:
            continue
        tabs = p["分ける先"]
        for kind, row in (("総支給", p["行"]), ("社会保険", "法定福利費")):
            total = h.get(kind, 0)
            if not total:
                continue
            each = total // len(tabs)
            rest = total - each * (len(tabs) - 1)      # 端数は先頭の店へ
            for i, tab in enumerate(tabs):
                out[(tab, row)] += rest if i == 0 else each
    return out, fallback, pool_detail


if __name__ == "__main__":
    import sys
    ym = sys.argv[1] if len(sys.argv) > 1 else "202608"
    for period in ("21期", "22期"):
        s, fb, pd_ = split(ym, period)
        print(f"── {period} の規則で {ym} を組むと")
        for tab in ("韓国酒場ハナ", "タコとハイボール"):
            for row in ("人件費（店長）", "人件費（アルバイト）", "法定福利費"):
                if s.get((tab, row)):
                    print(f"   {tab:<12}{row:<16}{s[(tab,row)]:>10,}")
        if pd_:
            print("   ★折半にまわした3人:")
            for no, name, g in pd_:
                print(f"      {no} {name:<16}{g:>9,}")
            print(f"      合計 {sum(g for _n,_m,g in pd_):,} → "
                  f"半分ずつ {sum(g for _n,_m,g in pd_)//2:,}")
        print()
