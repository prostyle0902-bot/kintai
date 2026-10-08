#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""毎月の元データ、届いているか／まだか（リマインド用）

--- 利用者指示 2026-10-08 ---------------------------------------------------
「毎月、ダウンロードする資料が沢山あるね〜 自動化するか、リマインドしてもらうかしないと
  忘れそうだな」→ 提案「毎月5日の朝に先月分でまだ届いていないものだけ知らせる。
  10日にまだ足りなければもう一度」に「うん、それお願い」。

    python3 remind.py              # 先月ぶん
    python3 remind.py 2026-09      # 月を指定
    python3 remind.py --have 給料一覧表   # Dropboxで確かめたものを「届いている」にする

定期チェック（Routine）が毎月5日・10日・25日の最初の回に呼ぶ。
「まだ」が0件なら何も知らせない。
★出そろう目安の日（期日）を過ぎたものだけ「まだ」にする。SBペイメントの後半（16日〜月末）は
  翌月22日ごろ発行なので、5日・10日には言わず25日に見る。

--- 何を見ているか -------------------------------------------------------
pl/ にファイルがあるか、または数字を焼いたモジュール（sbpay・airpay13・kessai）に
その月が入っているか。Dropbox に入れてもらったものは定期チェックで pl/ に落とすので
「pl/ にある＝取り込み済み」でよい。
★給料一覧表は pl/ に置かない（個人の給与）。Routine が Dropbox の
  ※プロスタイル給与※ で「給料一覧表-YYYYMM」を探し、あれば --have 給料一覧表 を付ける。
★置き方は22期（2026-09〜）に合わせてある。21期のぶんを聞くと名前が違って「まだ」と出るものがある。
★ENEOS・陸事（CP請求鑑）はファイル名が「N月」だけで期が分からないので、ここでは見ない。
"""
import argparse
import calendar
import datetime
import glob
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))


def _next(ym):
    y, m = int(ym[:4]), int(ym[5:])
    return f"{y + (m == 12)}-{m % 12 + 1:02d}"


def items(ym):
    """(名前, どこから取るか, 届いているか, 出そろう目安の日 YYYY-MM-DD)"""
    y, m = int(ym[:4]), int(ym[5:])
    yyyymm, yymm = f"{y}{m:02d}", f"{y % 100:02d}{m:02d}"
    nx = _next(ym)
    nx6 = nx.replace("-", "")
    last = calendar.monthrange(y, m)[1]
    g = lambda p: glob.glob(os.path.join(BASE, p))
    d1 = f"{nx}-01"          # 月が終われば取れるもの
    d5 = f"{nx}-05"          # カードの明細・給料一覧表（翌月の頭に出る）

    yield ("千葉銀行の明細（口座ごと）", "ちばぎんビジネスWeb → 01_",
           len(g(f"bank/小見川支店_普通_*_{yyyymm}.csv")) >= 8, d1)
    yield ("PayPay銀行の明細", "PayPay銀行 → 01_", bool(g(f"bank/NBG_{yyyymm}.csv")), d1)
    yield ("freeeカードの明細", f"freee（{nx} の明細）→ 02_",
           bool(g(f"csv/statement-{nx}.csv")), d5)
    yield ("JCBカードの明細", f"JCB（{nx6}meisai.csv）→ 03_",
           bool(g(f"cards/{nx6}meisai.csv")), d5)
    yield ("三井住友カードの明細", f"三井住友（{nx6}.csv）→ 03_",
           bool(g(f"cards/{nx6}.csv")), d5)
    yield ("board売上", "board「合計請求書の一覧」→ 04_",
           bool(g(f"cards/board_{yymm}.csv")), d1)
    yield ("エアレジ売上（5店）", "Airレジ → 05_",
           len(g(f"uriage/{yymm}月/*_{yyyymm}01-{yyyymm}{last}.csv")) >= 5, d1)
    yield ("かめや 月間売上集計（焼きたて屋）", "かめや → 06_",
           bool(g(f"uriage/{yymm}月/焼きたて屋_*.xlsx")), d1)

    import sbpay
    halves = {r[0][8:] for r in sbpay.DATA if r[0][:7] == ym}
    alone = ym in sbpay.COMPLETE_ALONE
    yield ("十三里屋 SBペイメント 収納明細書（1〜15日）", "SBペイメント → 08_",
           alone or "01" in halves, f"{ym}-20")
    yield ("十三里屋 SBペイメント 収納明細書（16日〜月末）", "SBペイメント → 08_",
           alone or "16" in halves, f"{nx}-22")

    import airpay13
    yield ("十三里屋 Airペイ 振込明細", "Airペイ（振込明細）→ 09_",
           any(r[0] == ym for r in airpay13.DATA), d1)

    if ym >= "2026-09":
        import kessai
        tabs = sorted({t for t, _m in kessai.AIRPAY22})
        miss = [t for t in tabs if (t, ym) not in kessai.AIRPAY22]
        yield ("ほかの4店 Airペイ 振込明細" + (f"（まだ: {'・'.join(miss)}）" if miss else ""),
               "店長さんがGoogleドライブの店舗フォルダに置く（私が取りに行く）", not miss, d1)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("month", nargs="?")
    ap.add_argument("--have", action="append", default=[])
    ap.add_argument("--today", help="試すとき用 YYYY-MM-DD")
    a = ap.parse_args()
    if a.month:
        ym = a.month
    else:
        base = datetime.date.fromisoformat(a.today) if a.today else datetime.date.today()
        t = base.replace(day=1) - datetime.timedelta(days=1)
        ym = t.strftime("%Y-%m")
    rows = list(items(ym))
    kyuyo = os.path.exists(os.path.join(BASE, "kyuyo", ym.replace("-", "") + ".pdf"))
    rows.append(("給料一覧表", "給与ソフト → ※プロスタイル給与※",
                 kyuyo or "給料一覧表" in a.have, f"{_next(ym)}-05"))
    today = a.today or datetime.date.today().isoformat()
    ng = [r for r in rows if not r[2] and r[3] <= today]          # 目安を過ぎてまだ
    later = [r for r in rows if not r[2] and r[3] > today]       # まだ出ていないはず
    print(f"{int(ym[5:])}月ぶん  届いている {sum(r[2] for r in rows)} ／ まだ {len(ng)}"
          + (f" ／ これから出るもの {len(later)}" if later else ""))
    for name, where, _ok, _d in ng:
        print(f"  × {name}（{where}）")
    for name, _w, _ok, d in later:
        print(f"  … {name}（{int(d[5:7])}/{int(d[8:])}ごろ）")
    return 0 if not ng else 1


if __name__ == "__main__":
    sys.exit(main())
