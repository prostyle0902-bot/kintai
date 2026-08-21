#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""既存の21期PLシート10枚から、新シートの【空いているセルだけ】を埋める

--- なぜ要るのか ---------------------------------------------------------
請求書・カード明細・銀行明細から積み上げてきたが、それでも新シートは2割しか
埋まっていなかった。既存の21期PLは1年ぶん手で入れてある「記録」なので、
書類から作れないセルはこれを写すのがいちばん確か。
利用者指示「今まで追加した資料、Dropbox、既存スプシを見て、2025年9月から
21期を全部埋めてみて」（2026-08-21）。

--- 大原則 ---------------------------------------------------------------
★既存シートには絶対に書き込まない。exist_dump.py が values().get で読むだけ。
★新シートで【空いているセルだけ】を埋める。書類から入れた値は絶対に上書きしない。
  重なったセルは check() が拾って「食い違い一覧」に出す。判断は利用者に委ねる。
★行の対応がつかないものは推測で入れない。保留リストへ出す。

--- 既存シートの形（10枚とも同じ。2026-08-21 確認）---------------------
    4行目がヘッダ  列1=項目 ／ 列2=年計 ／ 列3〜14=9月〜8月
    以降の行       列0=大分類 ／ 列1=項目名 ／ 列3〜=月次
読み込みは exist_pl.py。

--- 行の対応づけ ---------------------------------------------------------
① 行名がそのまま一致 … 何もしなくてよい（大半がこれ。1,004セル）
② 名前が違うだけ     … ALIAS で読み替える。全部この下の実測で裏を取った
③ 受け皿の行が無い   … build2.py に行を足した（12行。2026-08-21）
④ 対応がつかない     … SKIP に理由を書いて保留リストへ

--- ②の裏取り（2026-08-21 実測）---------------------------------------
人件費がいちばん紛らわしかった。既存は店舗により「人件費」「人件費（社員）」と
まちまちで、新シートは「人件費（店長）」「人件費（アルバイト）」に分かれている。
payroll.py が入れている4〜7月と数値を突き合わせて決めた:
    タコハイ・業務課・鳥害対策課・神栖横丁 … 既存＝新「人件費（店長）」と完全一致
    さわら十三里屋                        … 既存＝新「人件費（アルバイト）」と完全一致
                                           （4〜7月とも1円まで一致）
ほかの読み替えも、重なる月の金額が同じ桁・同じ費目であることを確かめてある。

--- ④に落としたもの -----------------------------------------------------
・りゅうちゃん「沖縄六角堂（8％対象）」「〃」「沖縄六角堂（10％対象）」
  内訳の作業用の行。合計が「仕入（沖縄六角堂）」と合わない（1,278,799 対 1,181,018）。
  「〃」というディット記号の行まである。何を指すか決められないので入れない。
・さわら十三里屋「水道光熱費」
  実測すると中身は電気料金だけだった（10月 49,577 が新シートの電気と1円まで一致）。
  なめがたの請求書から11か月とも入っているので、写す必要がない。
・業務課「仕入（クリーン＆ケミカル）」全月0。写す意味がない。
"""
import exist_pl

MONTHS = exist_pl.MONTHS

# (タブ, 既存の行名) -> 新シートの行名
ALIAS = {
    ("さわら十三里屋", "仕入（Freeeカードリアル）"): "仕入（freeeカード）",
    ("さわら十三里屋", "消耗品費（Freeeカード）"): "消耗品費（freeeカード）",
    ("さわら十三里屋", "人件費"): "人件費（アルバイト）",   # 4〜7月がpayrollと完全一致
    ("さわら十三里屋", "広告宣伝費"): "広告宣伝費（共通宣伝費）",
    ("さわら十三里屋", "地代家賃"): "地代家賃（賃料）",
    ("りゅうちゃん", "売上（税込）"): "売上",
    ("りゅうちゃん", "仕入（freeeカード山中ストアー）"): "仕入（やまなか）",
    ("りゅうちゃん", "仕入（沖縄六角堂）"): "仕入（六角堂）",
    ("りゅうちゃん", "仕入（藤原ストアー）"): "仕入（藤原ストア）",
    ("りゅうちゃん", "仕入（平良洋酒店）"): "仕入（平洋酒店）",
    ("タコとハイボール", "人件費（社員）"): "人件費（店長）",
    ("業務課", "人件費（社員）"): "人件費（店長）",
    ("業務課", "地代家賃"): "地代家賃（賃料）",
    ("業務課", "外注費（SUN-X)"): "外注費（SUN-X）",        # 既存は括弧が半角
    ("業務課", "広告宣伝費"): "広告宣伝費（共通宣伝費）",
    ("焼きたて屋", "消費税#2"): "出前館消費税",   # 既存は「消費税」が2つある
    ("焼きたて屋", "広告宣伝費"): "広告宣伝費（共通宣伝費）",
    ("焼きたて屋", "耗品費"): "消耗品費",                   # 既存の誤記
    ("神栖横丁", "人件費（社員）"): "人件費（店長）",
    ("神栖横丁", "地代家賃（ともえ）"): "地代家賃（賃料）",
    ("神栖横丁", "電気"): "水道光熱費（電気料金）",
    ("神栖横丁", "水道"): "水道光熱費（水道料金）",
    ("神栖横丁", "ガス"): "水道光熱費（ガス料金）",
    ("神栖横丁", "ドリーム"): "居酒屋ドリーム",
    ("鳥害対策課", "人件費"): "人件費（店長）",
    ("鳥害対策課", "地代家賃"): "地代家賃（賃料）",
}

# 入れないもの。(タブ, 既存の行名) -> 理由
SKIP = {
    ("りゅうちゃん", "沖縄六角堂（8％対象）"):
        "内訳の作業用の行。8%1,097,585＋10%181,214＝1,278,799 で、"
        "「仕入（沖縄六角堂）」1,181,018 と合わない。どちらが正か決められない",
    ("りゅうちゃん", "沖縄六角堂（10％対象）"):
        "同上。「仕入（沖縄六角堂）」との関係がはっきりしない",
    ("りゅうちゃん", "〃"):
        "ディット記号だけの行。年計1,016,278。何を指すか決められない",
    ("りゅうちゃん", "〃#2"):
        "同上。「〃」の行がもう1つある",
    ("さわら十三里屋", "水道光熱費"):
        "中身は電気料金だけだった（10月49,577が新シートの電気と1円まで一致）。"
        "なめがたの請求書から11か月とも入っているので写す必要がない",
    ("業務課", "仕入（クリーン＆ケミカル）"): "全月0円",
}

# 小計・利益・比率の行。新シートは数式で持っているので写さない
import re
_CALC = re.compile(r"合計|利益|比率|\(\d+\)|＝|=|期首|期末|売上原価\(a\)")


def _target(tab, item):
    """既存の行名 → 新シートの行名。写さないものは None。"""
    if (tab, item) in SKIP or _CALC.search(item):
        return None
    return ALIAS.get((tab, item), item)


def rows(wb):
    """(タブ, PL行, 月, 値, 元, メモ) を列挙。空いているセルだけ。"""
    import build2
    ex, order = exist_pl.load()
    for tab in order:
        if tab not in build2.RIDX or tab not in wb.sheetnames:
            continue
        ws = wb[tab]
        for item in order[tab]:
            plrow = _target(tab, item)
            if plrow is None or plrow not in build2.RIDX[tab]:
                continue
            for m in MONTHS:
                v = ex[tab][item].get(m)
                if not v:            # 未入力・0円は写さない
                    continue
                c = ws[f"{build2.MCOL[m]}{build2.RIDX[tab][plrow]}"]
                if c.value:          # 書類から入っている。絶対に上書きしない
                    continue
                note = "既存21期PLから転記"
                if plrow != item:
                    note += f"（既存の行名「{item}」）"
                yield (tab, plrow, m, int(v), f"既存21期PL {tab}", note)


def conflicts(wb):
    """既存PLと新シートで金額が食い違うセル。(タブ, 行, 月, 新, 既存)"""
    import build2
    ex, order = exist_pl.load()
    out = []
    for tab in order:
        if tab not in build2.RIDX or tab not in wb.sheetnames:
            continue
        ws = wb[tab]
        for item in order[tab]:
            plrow = _target(tab, item)
            if plrow is None or plrow not in build2.RIDX[tab]:
                continue
            for m in MONTHS:
                v = ex[tab][item].get(m)
                c = ws[f"{build2.MCOL[m]}{build2.RIDX[tab][plrow]}"].value
                if not v or not c or isinstance(c, str):
                    continue
                if int(c) != int(v):
                    out.append((tab, plrow, m, int(c), int(v)))
    return out


def check(wb=None):
    """対応づけの取りこぼしが無いかを見る。"""
    import build2
    ex, order = exist_pl.load()
    lost = []
    for tab in order:
        if tab not in build2.RIDX:
            continue
        for item in order[tab]:
            if not any(ex[tab][item].values()):
                continue
            plrow = _target(tab, item)
            if plrow is None:
                continue
            if plrow not in build2.RIDX[tab]:
                lost.append((tab, item, sum(ex[tab][item].values())))
    assert not lost, ("既存PLに値があるのに新シートに行が無い（ALIAS か SKIP か "
                      "build2.py の行追加が要る）:\n  " +
                      "\n  ".join(f"{t} 「{i}」 {s:,}円" for t, i, s in lost))


def hold_rows():
    """保留リストへ出すもの。"""
    ex, _ = exist_pl.load()
    out = []
    for (tab, item), why in SKIP.items():
        s = sum(ex.get(tab, {}).get(item, {}).values())
        if not s:
            continue
        out.append(("", tab, f"既存PL「{item}」（年計{s:,}円）", why))
    return out


if __name__ == "__main__":
    import warnings; warnings.filterwarnings("ignore")
    import openpyxl, collections
    check()
    wb = openpyxl.load_workbook("損益計算書_21期テスト版.xlsx")
    rs = list(rows(wb))
    by = collections.Counter()
    amt = collections.Counter()
    for tab, _r, _m, v, _s, _n in rs:
        by[tab] += 1; amt[tab] += v
    print(f"{'タブ':<14}{'埋めるセル':>10}{'金額':>16}")
    print("-" * 42)
    for tab in by:
        print(f"{tab:<14}{by[tab]:>10}{amt[tab]:>16,}")
    print("-" * 42)
    print(f"{'合計':<14}{sum(by.values()):>10}{sum(amt.values()):>16,}")
    print(f"\n食い違い（新シート優先。上書きしない）: {len(conflicts(wb))}セル")
    print(f"保留へ出すもの: {len(hold_rows())}件")
