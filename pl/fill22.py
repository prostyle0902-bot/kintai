#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""22期（2026.9〜2027.8）の損益計算書を組み立てる

    python3 fill22.py            # → 損益計算書_22期.xlsx と 22期の転記一覧（out22_cells.csv）
    python3 fill22.py --list 9月 # その月に入るセルを一覧で出す（利用者確認用）

build22.py は「空の器＋inv22.py」だけだった。ここで元データからの転記を足す。
21期の fill2.py は触らない（21期の数字が動かないように、別のプロセスで組む）。

--- 入れているもの（2026-10-01 時点・9月から）------------------------------
  ① inv22.py              … 22期に入れる請求書
  ② エアレジ売上（日別）     … uriage/<YYMM月>/<店>_YYYYMMDD-YYYYMMDD.csv
                              （airregi.py の逆算をそのまま使う）
  ③ freeeカード            … csv/statement-YYYY-MM.csv（ファイル年月−1＝利用月）
                              （engine.py の仕分けをそのまま使う。計上月は利用日）
  ④ JCB・三井住友カード      … cards/YYYYMMmeisai.csv・cards/YYYYMM.csv（−1か月＝PL列）
  ⑤ 銀行の口座引落          … debits / store_bank / yokocho_bank / honbu_bank
                              （period22.py で22期の年月に差し替えて使う）
  ★先に入っているセル（請求書など）には銀行から足さない（debits.rows(wb) と同じ考え方）。

--- まだ入れていないもの（元データ待ち。届いたらここに足す）-----------------
  給与（給料一覧表）／横丁→各店の社内請求／かめや（焼きたて屋の売上・本部請求）／
  棚卸し／Airペイ等の手数料（kessai.py）／board売上（横丁）／出前館／
  各店の仕入先の請求書（鹿島食品・六角堂ほか。10月に届く）
"""
import csv
import glob
import os
import sys

import pandas as pd
from openpyxl.comments import Comment
from openpyxl.styles import PatternFill

import build2

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = "損益計算書_22期.xlsx"
CELLS = "out22_cells.csv"

F_POST = PatternFill("solid", fgColor="FFF9C4")

# 22期の月（フォルダ名 → PL列）
MONTHS22 = {"2609月": "9月", "2610月": "10月", "2611月": "11月", "2612月": "12月",
            "2701月": "1月", "2702月": "2月", "2703月": "3月", "2704月": "4月",
            "2705月": "5月", "2706月": "6月", "2707月": "7月", "2708月": "8月"}

# freeeカード明細（ファイル年月−1か月＝利用月）。届いたら1行足す
STATEMENTS22 = [("csv/statement-2026-10.csv", 2026, 9)]

# JCB・三井住友（支払月−1か月＝PL列）。届いたら1行足す
CARDS22 = [("202610meisai.csv", "JCB", "9月"),
           ("202610.csv", "三井住友", "9月")]

# ★十三里屋の固定費カード0538（ドコモ・USEN・ダスキン）は engine.py が「除外」にする。
#   21期は既存PLの定額（fixed_costs.py）と二重にならないようにするためだった。
#   22期には定額の表が無いので、除外にした行をここで十三里屋の行へ入れ直す。
#   利用者指示 2026-10-02「十三里屋のリース料（ダスキン）は freeeカードの明細に
#   入ってるから、そこから引っ張ってきてほしい。22期も同じように」。
#   ドコモ・USEN も同じカードで同じ理由で除外されていて、22期9月の通信費が空いていた。
#   税抜は 22期の決まりどおり ÷1.1 の円未満切り捨て。月は利用日で決める。
CARD0538 = [("ダスキン", "リース料（ダスキン）"),
            ("ドコモ", "通信費（USEN、Wi-Fi）"),
            ("USEN", "通信費（USEN、Wi-Fi）")]

# カード明細の PL行 → そのタブでの呼び名（fill2.REMAP と同じ）
REMAP = {("韓国酒場ハナ", "仕入（やまなか）"): "仕入（山中ストアー）",
         ("もも焼きJAPAN", "仕入（やまなか）"): "仕入（freeeカード）",
         ("タコとハイボール", "仕入（やまなか）"): "仕入（freeeカード）"}


class Book:
    """セルへの足し込みと、何をどこに入れたかの記録"""

    def __init__(self, wb):
        self.wb = wb
        self.cells = []        # (タブ, 行, 月, 金額, 区分, 元データ, メモ)
        self.missing = []      # 行が無かったもの
        self.hold = []         # 入れなかったもの (月, タブ, 何, 理由)

    def has(self, tab, plrow, m):
        c = self.wb[tab][f"{build2.MCOL[m]}{build2.RIDX[tab][plrow]}"]
        return bool(c.value)

    def add(self, tab, plrow, m, val, kind, src, note=""):
        if plrow not in build2.RIDX[tab]:
            self.missing.append((tab, plrow, m, val, kind, src))
            return
        c = self.wb[tab][f"{build2.MCOL[m]}{build2.RIDX[tab][plrow]}"]
        c.value = int(c.value or 0) + int(val)
        c.fill = F_POST
        c.number_format = build2.NUMFMT
        self.cells.append((tab, plrow, m, int(val), kind, src, note))


# ---------------------------------------------------------------- ① 請求書
def post_invoices(bk):
    import inv22
    inv22.check(bk.wb)
    for tab, plrow, m, ex, _tax, vendor, src, biko in inv22.rows():
        bk.add(tab, plrow, m, ex, "請求書", src, f"{vendor}。{biko}")


# ---------------------------------------------------------------- ② エアレジ
def post_airregi(bk):
    import airregi
    for folder, m in MONTHS22.items():
        d = os.path.join(BASE, "uriage", folder)
        if not os.path.isdir(d):
            continue
        for name, (tab, urirow, zeirow) in airregi.STORES.items():
            if name == "焼きたて屋":
                continue        # 焼きたて屋はかめや（税込・税抜）が要る。下で保留に出す
            hit = glob.glob(os.path.join(d, f"{name}_*.csv"))
            if not hit:
                continue
            inc, tax, days, t10, t8 = airregi._csv(hit[0])
            src = f"会計明細/22期/{folder}/{os.path.basename(hit[0])}"
            note = (f"エアレジ売上集計（日別）{days}日分。税込{inc:,}（10%標準 {t10:,}／8%軽減 {t8:,}）。"
                    f"消費税{tax:,}は日ごと・税率ごとの逆算の合計")
            if zeirow:
                bk.add(tab, urirow, m, inc, "エアレジ", src, note + "。売上行は【税込】")
                bk.add(tab, zeirow, m, tax, "エアレジ", src, note + "。消費税行")
            else:
                bk.add(tab, urirow, m, inc - tax, "エアレジ", src,
                       note + "。★この店は消費税行が無いので【税抜】")
        if glob.glob(os.path.join(d, "焼きたて屋_*.xlsx")):
            have = [os.path.basename(p) for p in glob.glob(os.path.join(d, "焼きたて屋_*.xlsx"))]
            if not any("税込" in h for h in have):
                bk.hold.append((m, "焼きたて屋", "売上（税込）・消費税",
                                "FCの月間売上集計一覧表が【税抜】版しか無い（"
                                + "・".join(have) + "）。税込版か、かめやの合計精算書が届けば入る"))


# ---------------------------------------------------------------- ③ freeeカード
def post_freee(bk):
    import engine
    engine.PERIOD = [(2026, m) for m in range(9, 13)] + [(2027, m) for m in range(1, 9)]
    frames = []
    for path, y, mo in STATEMENTS22:
        if not os.path.exists(os.path.join(BASE, path)):
            continue
        d = engine.load(os.path.join(BASE, path), y, mo)
        d["_srcfile"] = os.path.basename(path)
        frames.append(d)
    if not frames:
        return
    ok, hold = engine.classify(pd.concat(frames, ignore_index=True))
    ok.to_csv("out_meisai22.csv", index=False, encoding="utf-8-sig")
    hold.to_csv("out_hold22.csv", index=False, encoding="utf-8-sig")
    live = ok[ok["判定"] != "除外"].copy()
    live["税抜"] = live["税抜"].astype(int)
    for (tab, plrow, m), g in live.groupby(["店舗", "PL行", "計上月"]):
        plrow = REMAP.get((tab, plrow), plrow)
        bk.add(tab, plrow, m, int(g["税抜"].sum()), "freeeカード",
               "・".join(sorted(set(g["元ファイル"]))), f"{len(g)}件")
    for _, r in hold.iterrows():
        bk.hold.append(("", str(r.get("店舗", "")), f"freeeカード {r['利用日']} {r['取引先']} {r['税込']:,}",
                        str(r["理由"])))
    _post_0538(bk, ok[ok["判定"] == "除外"])


def _post_0538(bk, ex):
    """engine.py が除外にした 十三里屋カード0538 の固定費を 22期の行へ。"""
    for _, r in ex.iterrows():
        if not str(r["カード"]).endswith("(0538)"):
            continue
        name = engine_norm(r["取引先"])
        row = next((pl for k, pl in CARD0538 if engine_norm(k) in name), None)
        assert row, f"0538 の除外行で行き先が決まらない: {r['取引先']}"
        y, mo = int(r["利用日"][:4]), int(r["利用日"][5:7])
        m = f"{mo}月"
        if (y, mo) not in PERIOD22:
            bk.hold.append((m, "さわら十三里屋", f"freeeカード {r['利用日']} {r['取引先']} {int(r['税込']):,}",
                            "利用日が22期の外。21期8月ぶんなら21期に入れる（fill2.py 側）"))
            continue
        inc = int(r["税込"])
        bk.add("さわら十三里屋", row, m, inc * 10 // 11, "freeeカード",
               r["元ファイル"], f"{r['利用日']} {r['取引先']} 税込{inc:,}（カード0538）")


def engine_norm(s):
    import engine
    return engine.norm(s)


PERIOD22 = [(2026, m) for m in range(9, 13)] + [(2027, m) for m in range(1, 9)]


# ---------------------------------------------------------------- ④ JCB・三井住友
def post_cards(bk):
    import cards
    for fname, issuer, m in CARDS22:
        path = os.path.join(cards.DIR, fname)
        if not os.path.exists(path):
            continue
        reader = cards._read_jcb if issuer == "JCB" else cards._read_smcc
        det = list(reader(path))
        if issuer == "JCB":
            head = cards._jcb_total(path)
            total = sum(v for _, _, v in det)
            assert head is None or total == head, \
                f"{fname}: 明細合計 {total:,} ≠ 今回のお支払金額合計 {head:,}"
        per = {}
        for used, merchant, inc in det:
            plrow = cards._classify(issuer, merchant)
            ex = cards._ex(inc)
            if plrow == "接待交際費" and ex <= cards.KAIGI_LIMIT and not cards.is_gift(merchant):
                plrow = "会議費"
            if plrow is None:
                bk.hold.append((m, "本部", f"{issuer} {used} {merchant} {inc:,}",
                                "cards.py の取引先マスタに無い。費目を決めてください"))
                continue
            per.setdefault(plrow, []).append((used, merchant, ex))
        for plrow, items in per.items():
            bk.add("本部", plrow, m, sum(x[2] for x in items), issuer,
                   f"freeeカード明細/22期/{fname}",
                   " ／ ".join(f"{u} {mm} {e:,}" for u, mm, e in items))


# ---------------------------------------------------------------- ⑤ 銀行
def post_bank(bk):
    import period22
    period22.apply()
    import debits, store_bank, yokocho_bank, honbu_bank
    for tab, plrow, m, ex, _tax, name, src, note in debits.rows(bk.wb):
        bk.add(tab, plrow, m, ex, "口座引落", src, note)
    for name, mod in (("店舗口座", store_bank), ("横丁口座", yokocho_bank),
                      ("本部口座", honbu_bank)):
        for tab, plrow, m, ex, src, note in mod.rows():
            if _same_vendor_in(bk, tab, plrow, m, note):
                # ★請求書（inv22）が先に【同じ取引先】を入れている。銀行から足すと二重になる
                bk.hold.append((m, tab, f"{plrow}（{name}）",
                                f"請求書から入っているので銀行（{ex:,}）は足さない。{note}"))
                continue
            bk.add(tab, plrow, m, ex, name, src, note)


def _same_vendor_in(bk, tab, plrow, m, note):
    """そのセルに、銀行の引落と【同じ取引先】の請求書が既に入っているか。

    ★以前はセルに何か入っていれば銀行ぶんを捨てていた。2026-10-07 に 神栖横丁
      事務消耗品費 9月 へ アスクル（請求書）を入れたら、同じセルの 関彰商事（横丁口座の
      引落）が消えた。yokocho_bank の注意書き「行ごとに飛ばすと事故る」と同じこと。
      取引先の名前（メモの先頭『。』まで）で見て、別の相手なら両方入れる。
    """
    who = note.split("。")[0].strip()
    for t, r, mm, _v, _k, _src, n in bk.cells:
        if (t, r, mm) != (tab, plrow, m):
            continue
        head = (n or "").split("。")[0].strip()
        if who and head and (who in (n or "") or head in who):
            return True
    return False


# ---------------------------------------------------------------- ⑧ SBペイメント（十三里屋）
# PayPay決済端末の収納明細書（半月ごと）の手数料。sbpay22.py に読んだ数字がある。
def post_sbpay(bk):
    import sbpay22
    for tab, plrow, m, ex, _tax, vendor, src, note in sbpay22.rows():
        bk.add(tab, plrow, m, ex, "SBペイメント", src, f"{vendor}。{note}")


# ---------------------------------------------------------------- ⑦ board売上（業務課・鳥害対策課）
# 利用者が ★毎月ここに入れる/04_board売上/ に置く「合計請求書の一覧」エクスポート → cards/board_YYMM.csv。
# ★利用者指示 2026-10-06「飲食は全部出てないから、業務、鳥害だけ入れて」
#   → グループ 業務課・鳥害対策課 だけ入れる。飲食事業部はまだ入れない（横丁の売上は後日）。
#   グループ空欄は board.GROUP_BY_CUSTOMER（21期と同じ: PlusOne・シナネンアクシア・亀甲堂→業務課）。
#   TAKEOUTPARK神栖横丁（グリスト清掃）も21期と同じく業務課の売上。
# 主キーは 合計請求書No（この形式には請求書IDが無い）。請求日がその月でない行があれば止める。
BOARD22 = [("board_2609.csv", "9月", "2026-09")]
BOARD22_GROUPS = {"業務課", "鳥害対策課"}


def post_board(bk):
    import board
    for name, m, ym in BOARD22:
        path = os.path.join(BASE, "cards", name)
        if not os.path.exists(path):
            continue
        rows = {r["合計請求書No"]: r for r in board._read(path)}.values()
        by = {}
        for r in rows:
            assert r["請求日"][:7] == ym, f"{name}: 請求日 {r['請求日']} が {ym} でない"
            g = board._group(r)
            if g not in BOARD22_GROUPS:
                continue
            ex = int(float(r["請求金額（JPY・税抜）"]))
            by.setdefault(g, []).append((ex, r["顧客名"]))
        for g, items in sorted(by.items()):
            items.sort(key=lambda x: -x[0])
            note = f"board {m}分 {len(items)}件: " + "／".join(f"{c} {e:,}" for e, c in items[:15])
            bk.add(g, "売上", m, sum(e for e, _c in items), "board",
                   f"★毎月ここに入れる/04_board売上（cards/{name}）", note)


# ---------------------------------------------------------------- ⑥ 給与
# 給料一覧表PDF（Dropbox /※プロスタイル給与※/プロスタイル給与R8年/給料一覧表-YYYYMM.pdf）を
# pl/kyuyo/YYYYMM.pdf に落としておくと入る。★個人の給与が載るのでリポジトリには入れない。
# 割り振りは kyuyo_split.split(ym, "22期")（22期の折半ルールもそこにある）。
# ★りゅうちゃん店長（河野竜二）の業務委託料は給料一覧表に載らない。22期は請求書（inv22）から入れる。
def post_payroll(bk):
    import collections
    import kyuyo_parse
    import genba_split
    import kyuyo_split
    for folder, m in MONTHS22.items():
        ym = "20" + folder[:4]
        if not os.path.exists(os.path.join(BASE, "kyuyo", f"{ym}.pdf")):
            bk.hold.append((m, "（全タブ）", "人件費・法定福利費",
                            f"給料一覧表-{ym}.pdf がまだ無い（※プロスタイル給与※/プロスタイル給与R8年）"))
            break       # 先の月はまだ来ていない。いちばん早い未着の月だけ出す
        split, _fb, _pool = kyuyo_split.split(ym, "22期")
        emp = kyuyo_parse.parse(ym)[0]
        nm = kyuyo_parse.names(ym)
        who = collections.defaultdict(list)
        # ★現場カレンダーの鳥害日数で分けた人は、両方のタブに割合つきで書く（genba_split.py）
        genba = kyuyo_split.genba_moves(ym, "22期", emp, nm)
        for no in sorted(emp):
            tab, _how = kyuyo_split.tab_of(no, nm.get(no, ""))
            row = kyuyo_split.row_of(no)
            if no in genba:
                b, t = genba[no]
                who[("鳥害対策課", row)].append(f"{nm.get(no, '')}({no})鳥害{b:.1f}/{t:.1f}日")
                if b < t:
                    who[("業務課", row)].append(f"{nm.get(no, '')}({no})清掃{t - b:.1f}/{t:.1f}日")
                continue
            who[(tab, row)].append(f"{nm.get(no, '')}({no})")
        sc = genba_split.shacho("22期", m)
        if sc:
            nmS = genba_split.SHACHO_NAME
            who[("本部", "人件費　社長")] = [f"{nmS}（現場に出た{sc['鳥害日数'] + sc['清掃日数']:.1f}日ぶん"
                                         f" {sc['鳥害'] + sc['業務課']:,}円を鳥害対策課・業務課へ移した残り）"]
            if sc["鳥害"]:
                who[("鳥害対策課", "人件費（店長）")].append(
                    f"{nmS}(社長)鳥害{sc['鳥害日数']:.1f}日×日当{sc['日当']:,}")
            if sc["業務課"]:
                who[("業務課", "人件費（店長）")].append(
                    f"{nmS}(社長)清掃{sc['清掃日数']:.1f}日×日当{sc['日当']:,}")
        src = f"給料一覧表-{ym}.pdf（Dropbox /※プロスタイル給与※/）"
        if genba or sc:
            src += "＋現場カレンダー「予定」（鳥害の日数で業務課・鳥害対策課に分けた）"
        for (tab, row), v in sorted(split.items()):
            if row == "法定福利費":
                note = f"給料一覧表{m}分の社会保険料計（{tab}ぶん）"
            else:
                names = who[(tab, row)]
                note = f"給料一覧表{m}分の総支給額（{len(names)}人）: " + "／".join(names[:12])
            bk.add(tab, row, m, int(v), "給与", src, note)


# ---------------------------------------------------------------- 仕上げ
def _comment(bk):
    by = {}
    for tab, plrow, m, val, kind, src, note in bk.cells:
        by.setdefault((tab, plrow, m), []).append((val, kind, src, note))
    for (tab, plrow, m), items in by.items():
        c = bk.wb[tab][f"{build2.MCOL[m]}{build2.RIDX[tab][plrow]}"]
        lines = [f"{tab}／{plrow}／{m}  計 {sum(x[0] for x in items):,} 円（税抜）"]
        for val, kind, src, note in items:
            lines.append(f"・{kind} {val:,}  出どころ: {src}")
            if note:
                lines.append(f"   {note[:300]}")
        c.comment = Comment("\n".join(lines), "自動転記")


def build():
    wb = build2.new_wb(period="22期")
    bk = Book(wb)
    post_invoices(bk)
    post_airregi(bk)
    post_freee(bk)
    post_cards(bk)
    post_bank(bk)
    post_payroll(bk)
    post_board(bk)
    post_sbpay(bk)
    _comment(bk)
    assert not bk.missing, "行が見つからない: " + "／".join(
        f"{t} {r} {m} {v:,}（{k}）" for t, r, m, v, k, _s in bk.missing)
    wb.save(OUT)
    with open(CELLS, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["タブ", "PL行", "月", "税抜", "区分", "元データ", "メモ"])
        w.writerows(bk.cells)
    return bk


def listing(bk, month):
    import inv22
    rows = [c for c in bk.cells if c[2] == month]
    by = {}
    for tab, plrow, m, val, kind, src, note in rows:
        by.setdefault(kind, []).append((tab, plrow, val))
    print(f"■ 22期 {month} に入れるもの {len(rows)}件")
    for kind, items in by.items():
        print(f"\n【{kind}】 {len(items)}件 計 {sum(v for _t, _r, v in items):,}")
        for tab, plrow, val in items:
            mark = (" ★要確認" if kind == "請求書"
                    and (tab, plrow, month) in getattr(inv22, "REVIEW", {}) else "")
            print(f"  {tab:<10} {plrow:<24} {val:>11,}{mark}")
    hold = [h for h in bk.hold if h[0] in (month, "")]
    if hold:
        print(f"\n■ 入れなかったもの・保留 {len(hold)}件")
        for m, tab, what, why in hold:
            print(f"  {tab} {what}\n      {why[:200]}")


if __name__ == "__main__":
    bk = build()
    print(f"{OUT} を作成 ／ 転記 {len(bk.cells)} 件 ／ 計 {sum(c[3] for c in bk.cells):,}円（税抜）"
          f" ／ 保留 {len(bk.hold)} 件")
    if "--list" in sys.argv:
        listing(bk, sys.argv[sys.argv.index("--list") + 1])
