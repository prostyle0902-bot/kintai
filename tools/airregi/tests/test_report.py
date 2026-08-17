"""report.py のテスト。

実データ（2026年7月のタコとハイボール）と同じ数値でCSVを組み立て、
Airレジの実際のエクスポート（Shift-JIS・店舗ごとに列数が違う）を再現して検証する。

    python3 tools/airregi/tests/test_report.py
"""

import datetime as dt
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import report  # noqa: E402

# 決済手段の列まで持つ通常のエクスポート（列数が多い店舗）
FULL_HEADER = [
    "集計日", "売上", "客数", "客単価", "組数", "組単価", "商品点数",
    "現金売上高", "現金売上金額(10%対象)", "現金売上金額(8%軽減)", "現金売上金額(非課税)",
    "その他売上高", "その他売上金額(10%対象)", "その他売上金額(8%軽減)", "その他売上金額(非課税)",
    "値引額", "売上金額(10%対象)", "売上金額(8%軽減)",
    "クレジットカード(Airペイ) 売上金額", "QR(Airペイ QR) 売上金額",
]
# 決済手段の内訳を持たない店舗のエクスポート（列が少ない）
SHORT_HEADER = [
    "集計日", "売上", "客数", "客単価", "組数", "組単価", "商品点数",
    "現金売上高", "その他売上高", "値引額", "売上金額(10%対象)", "売上金額(8%軽減)",
]

# 2026-07 タコとハイボール: (日, 売上, 客数, 組数, 点数, 現金, その他)
TACO = [
    (2, 12350, 7, 8, 20, 11700, 650), (3, 74380, 30, 38, 108, 46650, 27730),
    (6, 29450, 17, 17, 41, 19550, 9900), (7, 20450, 13, 13, 29, 12000, 8450),
    (9, 53780, 18, 21, 75, 17750, 36030), (10, 55280, 18, 28, 85, 45150, 10130),
    (11, 58420, 32, 37, 87, 35550, 22870), (12, 31200, 15, 20, 50, 20400, 10800),
    (13, 46680, 20, 22, 61, 10050, 36630), (14, 32440, 19, 19, 47, 15200, 17240),
    (16, 28140, 17, 19, 47, 11350, 16790), (17, 46640, 21, 26, 70, 10000, 36640),
    (18, 61660, 33, 41, 90, 29100, 32560), (19, 69480, 49, 52, 95, 32450, 37030),
    (20, 38000, 30, 31, 60, 13350, 24650), (21, 21340, 16, 16, 31, 6950, 14390),
    (23, 24980, 15, 21, 37, 21350, 3630), (24, 35920, 10, 20, 53, 26900, 9020),
    (25, 64610, 32, 36, 95, 28600, 36010), (26, 26670, 11, 18, 41, 15000, 11670),
    (27, 45940, 16, 27, 69, 39700, 6240), (28, 23460, 12, 14, 32, 11050, 12410),
    (30, 51700, 17, 21, 66, 30600, 21100), (31, 42140, 17, 29, 66, 21900, 20240),
]
# 同店の2026-06（前月比の検証用・合計 852,270円 / 485組）
TACO_JUNE = [
    (1, 16250, 9, 10, 30, 10000, 6250), (2, 18050, 12, 14, 34, 12000, 6050),
    (4, 25980, 11, 12, 40, 15980, 10000), (5, 30640, 20, 22, 55, 20640, 10000),
    (6, 65580, 30, 33, 120, 40580, 25000), (7, 39000, 22, 25, 70, 24000, 15000),
    (8, 36070, 14, 16, 60, 21070, 15000), (9, 31440, 15, 17, 52, 19440, 12000),
    (11, 32280, 17, 19, 55, 20280, 12000), (12, 39970, 20, 22, 66, 24970, 15000),
    (13, 37920, 21, 24, 64, 22920, 15000), (14, 33830, 20, 22, 58, 20830, 13000),
    (15, 27200, 19, 22, 48, 17200, 10000), (16, 36140, 14, 16, 58, 22140, 14000),
    (18, 37890, 21, 24, 62, 23890, 14000), (19, 40520, 24, 27, 68, 25520, 15000),
    (20, 46220, 24, 27, 78, 28220, 18000), (21, 40730, 15, 17, 66, 25730, 15000),
    (22, 19630, 8, 9, 32, 12630, 7000), (23, 23660, 11, 13, 40, 14660, 9000),
    (25, 18320, 10, 11, 30, 11320, 7000), (26, 32090, 15, 17, 54, 20090, 12000),
    (27, 32530, 18, 20, 55, 20530, 12000), (28, 29580, 13, 15, 48, 18580, 11000),
    (29, 22400, 8, 9, 36, 14400, 8000), (30, 38350, 20, 22, 62, 23350, 15000),
]


def full_rows(month, rows):
    out = []
    for d, sales, guests, groups, items, cash, other in rows:
        out.append([
            f"2026{month:02d}{d:02d}", sales, guests, round(sales / guests), groups,
            round(sales / groups), items, cash, cash, 0, 0, other, other, 0, 0, 0,
            sales, 0, other, 0,
        ])
    return out


def short_rows(month, rows):
    out = []
    for d, sales, guests, groups, items, cash, other in rows:
        out.append([
            f"2026{month:02d}{d:02d}", sales, guests, round(sales / guests), groups,
            round(sales / groups), items, cash, other, 0, sales, 0,
        ])
    return out


def write_csv(path, header, rows):
    lines = [",".join(header)]
    lines += [",".join(str(c) for c in r) for r in rows]
    path.write_bytes("\n".join(lines).encode("cp932"))


class TestParsing(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.addCleanup(self.tmp.cleanup)

    def test_full_export_matches_known_totals(self):
        path = self.dir / "タコハイ７月.csv"
        write_csv(path, FULL_HEADER, full_rows(7, TACO))
        rows = report.load_csv(path)
        s = report.summarize(rows)
        self.assertEqual(s["sales"], 995_110)
        self.assertEqual(s["groups"], 594)
        self.assertEqual(s["days"], 24)
        self.assertEqual(round(s["unit"]), 1675)
        self.assertAlmostEqual(s["cashless"] * 100, 46.5, places=1)
        self.assertAlmostEqual(s["guest_input"] * 100, 82, delta=0.5)

    def test_short_export_reads_same_values(self):
        """列数が少ないエクスポートでも同じ数字になる（位置決め打ちにしない）。"""
        a = self.dir / "full.csv"
        b = self.dir / "short.csv"
        write_csv(a, FULL_HEADER, full_rows(7, TACO))
        write_csv(b, SHORT_HEADER, short_rows(7, TACO))
        sa, sb = report.summarize(report.load_csv(a)), report.summarize(report.load_csv(b))
        for key in ("sales", "groups", "guests", "items", "cash", "other"):
            self.assertEqual(sa[key], sb[key], key)

    def test_store_name_from_filename(self):
        cases = {
            "タコハイ７月.csv": "タコハイ",
            "もも焼きJAPAN7月.csv": "もも焼きJAPAN",
            "ハナ_2026-07.csv": "ハナ",
            "りゅうちゃん 2026年7月.csv": "りゅうちゃん",
        }
        for name, want in cases.items():
            self.assertEqual(report.store_name(Path(name)), want, name)

    def test_utf8_csv_is_accepted(self):
        path = self.dir / "utf8.csv"
        lines = [",".join(SHORT_HEADER)] + [
            ",".join(str(c) for c in r) for r in short_rows(7, TACO[:3])
        ]
        path.write_text("\n".join(lines), encoding="utf-8")
        self.assertEqual(len(report.load_csv(path)), 3)


class TestAggregation(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.cur = Path(self.tmp.name) / "07"
        self.prev = Path(self.tmp.name) / "06"
        self.cur.mkdir(parents=True)
        self.prev.mkdir(parents=True)
        write_csv(self.cur / "タコハイ７月.csv", FULL_HEADER, full_rows(7, TACO))
        write_csv(self.prev / "タコハイ６月.csv", FULL_HEADER, full_rows(6, TACO_JUNE))
        self.addCleanup(self.tmp.cleanup)

    def test_month_over_month_and_decomposition(self):
        rep = report.build(report.load_dir(self.cur), report.load_dir(self.prev))
        s = rep["stores"][0]
        self.assertEqual(s["prev"]["sales"], 852_270)
        self.assertAlmostEqual(s["mom"] * 100, 116.8, places=1)
        # 分解の合計が実際の増減額と一致する（誤差1円未満）
        self.assertAlmostEqual(
            s["volume_effect"] + s["price_effect"], s["sales"] - s["prev"]["sales"], delta=1
        )
        self.assertGreater(s["volume_effect"], 0)   # 組数は増えている
        self.assertLess(s["price_effect"], 0)       # 単価は下がっている

    def test_weekday_average(self):
        rows = report.load_csv(self.cur / "タコハイ７月.csv")
        wk = report.by_weekday(rows)
        self.assertEqual(round(wk["土"]), 61_563)   # 7/11,18,25 の平均
        self.assertEqual(round(wk["火"]), 24_422)

    def test_temporary_closure_and_loss(self):
        """他店が開いていた日に休んだ店だけを臨時休業として拾う。"""
        write_csv(self.cur / "となり店７月.csv", FULL_HEADER,
                  full_rows(7, TACO + [(4, 50000, 20, 25, 60, 30000, 20000)]))
        stores = report.load_dir(self.cur)
        closed = report.closed_days(stores)
        self.assertEqual(closed["タコハイ"], [dt.date(2026, 7, 4)])
        self.assertEqual(closed["となり店"], [])
        # 7/4は土曜 → 土曜平均で試算される
        loss = report.opportunity_loss(stores["タコハイ"], closed["タコハイ"])
        self.assertEqual(round(loss), 61_563)

    def test_html_output_is_self_contained(self):
        rep = report.build(report.load_dir(self.cur), report.load_dir(self.prev))
        out = report.render_html(rep, "2026年7月 売上レポート")
        self.assertIn("<!doctype html>", out)
        self.assertIn("995,110", out)
        self.assertNotIn("__DATA__", out)
        self.assertNotIn("http://", out)
        self.assertNotIn("https://", out)   # 外部読み込みなし＝オフラインで開ける


if __name__ == "__main__":
    unittest.main(verbosity=2)
