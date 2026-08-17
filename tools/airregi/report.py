#!/usr/bin/env python3
"""Airレジの「日別売上」CSVから月次レポート（コンソール＋HTML）を作る。

使い方:
    python3 tools/airregi/report.py CSVディレクトリ [--prev 前月のCSVディレクトリ]
                                    [--out report.html] [--month 2026-07]

Airレジ管理画面 → 売上 → 日別売上 → CSVダウンロード で店舗ごとに落としたCSVを
1つのフォルダに入れて実行する。店舗名はファイル名から取る（「タコハイ７月.csv」→「タコハイ」）。
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import html
import io
import json
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path

WEEKDAYS = "月火水木金土日"

# Airレジの列名（表記ゆれを含む）→ 内部キー。ヘッダー名で引き当てる。
COLUMN_KEYS = {
    "date": ("集計日", "日付"),
    "sales": ("売上",),
    "guests": ("客数",),
    "groups": ("組数",),
    "items": ("商品点数", "点数"),
    "cash": ("現金売上高", "現金売上"),
    "other": ("その他売上高", "その他売上"),
    "discount": ("値引額",),
}


def _norm(s: str) -> str:
    """全角スペース・カッコ差を吸収してヘッダーを比較しやすくする。"""
    return unicodedata.normalize("NFKC", s).replace(" ", "").strip()


def _read_text(path: Path) -> str:
    raw = path.read_bytes()
    for enc in ("cp932", "utf-8-sig", "utf-8"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise SystemExit(f"文字コードを判別できません: {path}")


def _resolve_columns(header: list[str]) -> dict[str, int]:
    """ヘッダー行から必要な列の位置を決める。

    Airレジのエクスポートは店舗の設定（決済手段・部門）で列数が変わるため、
    位置決め打ちにはしない。内訳列（「現金売上金額(10%対象)」など）は
    合計列より後ろに来るので、最初に一致したものを採用する。
    """
    cols: dict[str, int] = {}
    for i, name in enumerate(header):
        n = _norm(name)
        for key, candidates in COLUMN_KEYS.items():
            if key in cols:
                continue
            if any(n == c or n.startswith(c + "(") for c in candidates):
                cols[key] = i
    missing = [k for k in ("date", "sales") if k not in cols]
    if missing:
        raise SystemExit(f"必要な列が見つかりません（{', '.join(missing)}）")
    return cols


def _to_int(cell: str) -> int:
    cell = _norm(cell).replace(",", "").replace("¥", "")
    if not cell or cell == "-":
        return 0
    try:
        return int(float(cell))
    except ValueError:
        return 0


def store_name(path: Path) -> str:
    """ファイル名から店舗名を取る（末尾の「7月」「２０２６年７月」等を落とす）。"""
    stem = unicodedata.normalize("NFKC", path.stem)
    stem = re.sub(r"[_\s-]*\d{4}[-_年]?\d{1,2}月?$", "", stem)
    stem = re.sub(r"[_\s-]*\d{1,2}月$", "", stem)
    return stem.strip() or path.stem


def load_csv(path: Path) -> list[dict]:
    text = _read_text(path)
    rows = list(csv.reader(io.StringIO(text)))
    if not rows:
        return []
    cols = _resolve_columns(rows[0])
    out = []
    for row in rows[1:]:
        if not row or not _norm(row[cols["date"]]):
            continue
        raw_date = _norm(row[cols["date"]])
        m = re.match(r"^(\d{4})[-/]?(\d{2})[-/]?(\d{2})", raw_date)
        if not m:
            continue
        date = dt.date(*(int(g) for g in m.groups()))
        get = lambda key: _to_int(row[cols[key]]) if key in cols and cols[key] < len(row) else 0
        out.append(
            {
                "date": date,
                "dow": WEEKDAYS[date.weekday()],
                "sales": get("sales"),
                "guests": get("guests"),
                "groups": get("groups") or get("guests"),
                "items": get("items"),
                "cash": get("cash"),
                "other": get("other"),
                "discount": get("discount"),
            }
        )
    out.sort(key=lambda r: r["date"])
    return out


def load_dir(path: Path) -> dict[str, list[dict]]:
    files = sorted(p for p in path.glob("*.csv") if p.is_file())
    if not files:
        raise SystemExit(f"CSVが見つかりません: {path}")
    return {store_name(p): load_csv(p) for p in files}


# --------------------------------------------------------------------------- 集計


def summarize(rows: list[dict]) -> dict:
    sales = sum(r["sales"] for r in rows)
    groups = sum(r["groups"] for r in rows)
    guests = sum(r["guests"] for r in rows)
    items = sum(r["items"] for r in rows)
    cash = sum(r["cash"] for r in rows)
    other = sum(r["other"] for r in rows)
    days = len(rows)
    return {
        "sales": sales,
        "groups": groups,
        "guests": guests,
        "items": items,
        "cash": cash,
        "other": other,
        "days": days,
        "daily": sales / days if days else 0,
        "unit": sales / groups if groups else 0,          # 組単価（会計1件あたり）
        "items_per_group": items / groups if groups else 0,
        "cashless": other / (cash + other) if (cash + other) else 0,
        "guest_input": guests / groups if groups else 0,  # 人数入力率
    }


def by_weekday(rows: list[dict]) -> dict[str, float]:
    acc: dict[str, list[int]] = defaultdict(list)
    for r in rows:
        acc[r["dow"]].append(r["sales"])
    return {w: sum(v) / len(v) for w, v in acc.items()}


def closed_days(stores: dict[str, list[dict]]) -> dict[str, list[dt.date]]:
    """他店が営業していた日に休んでいた日（＝定休日ではない臨時休業）を拾う。"""
    open_days: dict[str, set] = {s: {r["date"] for r in rows} for s, rows in stores.items()}
    business = set().union(*open_days.values()) if open_days else set()
    return {s: sorted(business - days) for s, days in open_days.items()}


def opportunity_loss(rows: list[dict], missing: list[dt.date]) -> float:
    """休業日の同曜日平均を積み上げた機会損失の目安。"""
    wk = by_weekday(rows)
    return sum(wk.get(WEEKDAYS[d.weekday()], 0) for d in missing)


def decompose(cur: dict, prev: dict) -> tuple[float, float]:
    """増減額を「組数増による分」と「単価変動による分」に分解する。"""
    if not prev["groups"] or not cur["groups"]:
        return 0.0, 0.0
    volume = (cur["groups"] - prev["groups"]) * prev["unit"]
    price = (cur["unit"] - prev["unit"]) * cur["groups"]
    return volume, price


def build(stores: dict[str, list[dict]], prev: dict[str, list[dict]] | None) -> dict:
    order = sorted(stores, key=lambda s: -summarize(stores[s])["sales"])
    closed = closed_days(stores)
    report = {"stores": [], "month": None}
    dates = [r["date"] for rows in stores.values() for r in rows]
    if dates:
        report["month"] = f"{min(dates):%Y-%m}"
    for name in order:
        rows = stores[name]
        cur = summarize(rows)
        entry = {
            "name": name,
            **cur,
            "weekday": by_weekday(rows),
            "closed": closed.get(name, []),
            "loss": opportunity_loss(rows, closed.get(name, [])),
            "rows": rows,
        }
        if prev and name in prev:
            p = summarize(prev[name])
            vol, price = decompose(cur, p)
            entry["prev"] = p
            entry["mom"] = cur["sales"] / p["sales"] if p["sales"] else 0
            entry["mom_daily"] = (cur["daily"] / p["daily"]) if p["daily"] else 0
            entry["mom_unit"] = (cur["unit"] / p["unit"]) if p["unit"] else 0
            entry["volume_effect"] = vol
            entry["price_effect"] = price
        report["stores"].append(entry)
    report["total"] = summarize([r for rows in stores.values() for r in rows])
    report["total"]["days"] = sum(s["days"] for s in report["stores"])
    report["total"]["daily"] = (
        report["total"]["sales"] / report["total"]["days"] if report["total"]["days"] else 0
    )
    if prev:
        pt = summarize([r for rows in prev.values() for r in rows])
        report["prev_total"] = pt
        report["total_mom"] = report["total"]["sales"] / pt["sales"] if pt["sales"] else 0
        vol, price = decompose(report["total"], pt)
        report["total_volume_effect"] = vol
        report["total_price_effect"] = price
    return report


# --------------------------------------------------------------------------- 出力

def yen(v: float) -> str:
    n = round(v)
    return f"−¥{abs(n):,}" if n < 0 else f"¥{n:,}"


def pct(v: float) -> str:
    return f"{v * 100:.0f}%"


def print_console(rep: dict) -> None:
    t = rep["total"]
    print(f"■ {rep['month']} 合計 {yen(t['sales'])} / {t['groups']:,}組 / 組単価 {yen(t['unit'])}")
    if "total_mom" in rep:
        print(
            f"  前月比 {pct(rep['total_mom'])}"
            f"（組数増 {yen(rep['total_volume_effect']):>12} ／ 単価変動 {yen(rep['total_price_effect']):>12}）"
        )
    print()
    head = f"{'店舗':<18}{'売上':>11}{'前月比':>7}{'営業':>5}{'日商':>9}{'組数':>6}{'組単価':>8}{'点/組':>6}{'CL率':>6}{'人数入力':>8}"
    print(head)
    print("-" * len(head))
    for s in rep["stores"]:
        mom = pct(s["mom"]) if "mom" in s else "-"
        print(
            f"{s['name']:<18}{s['sales']:>11,}{mom:>7}{s['days']:>5}{s['daily']:>9,.0f}"
            f"{s['groups']:>6,}{s['unit']:>8,.0f}{s['items_per_group']:>6.1f}"
            f"{s['cashless'] * 100:>5.0f}%{s['guest_input'] * 100:>7.0f}%"
        )
    print()
    print("曜日別 1営業日あたり売上")
    print(f"{'店舗':<18}" + "".join(f"{w:>10}" for w in WEEKDAYS))
    for s in rep["stores"]:
        print(f"{s['name']:<18}" + "".join(
            f"{s['weekday'][w]:>10,.0f}" if w in s["weekday"] else f"{'-':>10}" for w in WEEKDAYS
        ))
    print()
    losses = [s for s in rep["stores"] if s["closed"]]
    if losses:
        print("臨時休業と機会損失の目安（同曜日平均で試算）")
        for s in losses:
            days = "、".join(f"{d:%-m/%-d}({WEEKDAYS[d.weekday()]})" for d in s["closed"])
            print(f"  {s['name']:<18}{len(s['closed'])}日  {days}  約{yen(s['loss'])}")
        print(f"  {'合計':<18}      約{yen(sum(s['loss'] for s in losses))}")
        print()
    low = [s for s in rep["stores"] if s["guest_input"] < 0.99]
    if low:
        print("※ 人数入力率が100%未満の店舗は、客単価が「1組あたり」になっています:")
        print("   " + "、".join(f"{s['name']} {s['guest_input'] * 100:.0f}%" for s in low))


# --------------------------------------------------------------------------- HTML

HTML_TEMPLATE = """<!doctype html>
<html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__TITLE__</title>
<style>
:root{--ground:#eceeed;--surface:#fbfcfb;--ink:#131a19;--ink-2:#3d4948;--muted:#65726f;
--line:rgba(19,26,25,.14);--line-soft:rgba(19,26,25,.07);--accent:#2b5f6b;--accent-soft:#8fb9c0;
--accent-wash:rgba(43,95,107,.10);--good:#2e7355;--bad:#a8483c;--flat:#7b8683;
--shadow:0 1px 2px rgba(19,26,25,.05),0 8px 24px rgba(19,26,25,.05);
--serif:"Hiragino Mincho ProN","Yu Mincho","Noto Serif JP",serif;
--sans:-apple-system,"Hiragino Sans","Noto Sans JP","Yu Gothic",sans-serif}
@media (prefers-color-scheme:dark){:root:not([data-theme="light"]){--ground:#0d1214;--surface:#151d1f;
--ink:#e6ecea;--ink-2:#c2cdcb;--muted:#8d9b98;--line:rgba(230,236,234,.16);--line-soft:rgba(230,236,234,.08);
--accent:#78b8c4;--accent-soft:#3d6d77;--accent-wash:rgba(120,184,196,.13);--good:#5fb68d;--bad:#dc8577;
--flat:#8d9b98;--shadow:0 1px 2px rgba(0,0,0,.4),0 10px 28px rgba(0,0,0,.35)}}
:root[data-theme="dark"]{--ground:#0d1214;--surface:#151d1f;--ink:#e6ecea;--ink-2:#c2cdcb;--muted:#8d9b98;
--line:rgba(230,236,234,.16);--line-soft:rgba(230,236,234,.08);--accent:#78b8c4;--accent-soft:#3d6d77;
--accent-wash:rgba(120,184,196,.13);--good:#5fb68d;--bad:#dc8577;--flat:#8d9b98;
--shadow:0 1px 2px rgba(0,0,0,.4),0 10px 28px rgba(0,0,0,.35)}
*{box-sizing:border-box}
body{margin:0;background:var(--ground);color:var(--ink);font-family:var(--sans);font-size:15px;line-height:1.75}
.wrap{max-width:1060px;margin:0 auto;padding:clamp(1.5rem,4vw,3rem) clamp(1rem,3vw,2rem) 4rem;
display:flex;flex-direction:column;gap:2.25rem}
.eyebrow{font-size:11px;letter-spacing:.18em;text-transform:uppercase;color:var(--muted);font-weight:600}
h1{font-family:var(--serif);font-weight:600;font-size:clamp(1.7rem,4vw,2.4rem);line-height:1.3;margin:.3rem 0 0}
h2{font-family:var(--serif);font-weight:600;font-size:1.25rem;margin:0}
.head{border-bottom:1px solid var(--line);padding-bottom:.55rem}
section{display:flex;flex-direction:column;gap:.9rem}
.note{font-size:13px;color:var(--muted);margin:0}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(185px,1fr));gap:.75rem}
.kpi{background:var(--surface);border:1px solid var(--line-soft);border-radius:4px;padding:1rem 1.1rem;
box-shadow:var(--shadow);display:flex;flex-direction:column;gap:.15rem}
.kpi .k{font-size:11.5px;letter-spacing:.08em;color:var(--muted);font-weight:600}
.kpi .v{font-family:var(--serif);font-size:1.8rem;line-height:1.2;font-variant-numeric:tabular-nums}
.kpi .s{font-size:12.5px;color:var(--ink-2);font-variant-numeric:tabular-nums}
.scroll{overflow-x:auto;background:var(--surface);border:1px solid var(--line-soft);border-radius:4px;box-shadow:var(--shadow)}
table{width:100%;border-collapse:collapse;font-variant-numeric:tabular-nums;font-size:13.5px}
th,td{padding:.55rem .7rem;text-align:right;white-space:nowrap;border-bottom:1px solid var(--line-soft)}
th{font-size:11px;letter-spacing:.06em;color:var(--muted);font-weight:600;border-bottom:1px solid var(--line)}
th:first-child,td:first-child{text-align:left}
tbody tr:last-child td{border-bottom:none}
tbody tr:hover td{background:var(--accent-wash)}
.chip{display:inline-block;font-size:12px;font-weight:600;padding:.1rem .45rem;border-radius:3px;font-variant-numeric:tabular-nums}
.up{color:var(--good);background:color-mix(in srgb,var(--good) 12%,transparent)}
.down{color:var(--bad);background:color-mix(in srgb,var(--bad) 12%,transparent)}
.flat{color:var(--flat);background:color-mix(in srgb,var(--flat) 12%,transparent)}
.split{display:grid;grid-template-columns:repeat(auto-fit,minmax(300px,1fr));gap:.75rem}
.card{background:var(--surface);border:1px solid var(--line-soft);border-radius:4px;padding:1rem 1.1rem .85rem;
box-shadow:var(--shadow);display:flex;flex-direction:column;gap:.6rem}
.card .cap{display:flex;justify-content:space-between;align-items:baseline;gap:.5rem}
.card h3{font-size:13px;font-weight:600;margin:0}
.card .cap span{font-size:12px;color:var(--muted);font-variant-numeric:tabular-nums}
.cols{display:flex;align-items:flex-end;gap:2px;height:104px}
.col{flex:1;display:flex;flex-direction:column;justify-content:flex-end;gap:4px;height:100%}
.col i{display:block;background:var(--accent-soft);border-radius:2px 2px 0 0}
.col.we i{background:var(--accent)}
.col span{font-size:9px;color:var(--muted);text-align:center;font-variant-numeric:tabular-nums}
.legend{display:flex;gap:1.1rem;font-size:12px;color:var(--muted);flex-wrap:wrap;align-items:center}
.legend b{display:inline-block;width:10px;height:10px;border-radius:2px;margin-right:.35rem;vertical-align:-1px}
footer{color:var(--muted);font-size:12.5px;border-top:1px solid var(--line);padding-top:1rem}
:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
</style></head><body><div class="wrap">
__BODY__
</div>
<script id="data" type="application/json">__DATA__</script>
<script>
const D=JSON.parse(document.getElementById("data").textContent);
document.getElementById("dailies").innerHTML=D.map(s=>{
  const max=Math.max(...s.rows.map(r=>r[2]));
  const bars=s.rows.map(([d,w,v])=>{
    const we=(w==="土"||w==="日");
    return `<div class="col${we?" we":""}" title="${d}（${w}） ¥${v.toLocaleString()}">`+
           `<i style="height:${(v/max*100).toFixed(1)}%"></i><span>${d.split("/")[1]}</span></div>`;
  }).join("");
  const avg=Math.round(s.rows.reduce((a,r)=>a+r[2],0)/s.rows.length);
  return `<div class="card"><div class="cap"><h3>${s.name}</h3>`+
         `<span>最高 ¥${max.toLocaleString()}／日商 ¥${avg.toLocaleString()}</span></div>`+
         `<div class="cols">${bars}</div></div>`;
}).join("");
</script></body></html>
"""


def chip(ratio: float) -> str:
    if not ratio:
        return '<span class="chip flat">-</span>'
    cls = "up" if ratio > 1.005 else "down" if ratio < 0.995 else "flat"
    mark = "▲" if cls == "up" else "▼" if cls == "down" else ""
    return f'<span class="chip {cls}">{mark}{ratio * 100:.0f}%</span>'


def render_html(rep: dict, title: str) -> str:
    t = rep["total"]
    e = html.escape
    parts = [
        '<header><div class="eyebrow">Airレジ 月次レポート</div>'
        f"<h1>{e(title)}</h1></header>",
        '<section><div class="kpis">',
        f'<div class="kpi"><span class="k">売上合計</span><span class="v">{yen(t["sales"])}</span>'
        + (
            f'<span class="s">前月 {yen(rep["prev_total"]["sales"])} → {chip(rep["total_mom"])}</span>'
            if "prev_total" in rep
            else '<span class="s">&nbsp;</span>'
        )
        + "</div>",
        f'<div class="kpi"><span class="k">組数（会計件数）</span><span class="v">{t["groups"]:,}</span>'
        f'<span class="s">1営業日あたり {t["groups"] / t["days"]:.1f}組</span></div>'
        if t["days"]
        else "",
        f'<div class="kpi"><span class="k">組単価</span><span class="v">{yen(t["unit"])}</span>'
        + (
            f'<span class="s">前月 {yen(rep["prev_total"]["unit"])} → '
            f'{chip(t["unit"] / rep["prev_total"]["unit"] if rep["prev_total"]["unit"] else 0)}</span>'
            if "prev_total" in rep
            else '<span class="s">&nbsp;</span>'
        )
        + "</div>",
        f'<div class="kpi"><span class="k">キャッシュレス比率</span><span class="v">{t["cashless"] * 100:.1f}%</span>'
        f'<span class="s">現金 {yen(t["cash"])}／CL {yen(t["other"])}</span></div>',
        "</div></section>",
    ]

    if "total_volume_effect" in rep:
        vol, price = rep["total_volume_effect"], rep["total_price_effect"]
        parts.append(
            '<section><div class="head"><h2>前月からの増減の内訳</h2></div>'
            '<div class="scroll"><table><thead><tr><th>店舗</th><th>増減額</th>'
            "<th>うち組数増による</th><th>うち単価変動による</th><th>前月組単価</th><th>今月組単価</th></tr></thead><tbody>"
            + "".join(
                f"<tr><td>{e(s['name'])}</td><td>{yen(s['sales'] - s['prev']['sales'])}</td>"
                f"<td>{yen(s['volume_effect'])}</td><td>{yen(s['price_effect'])}</td>"
                f"<td>{yen(s['prev']['unit'])}</td><td>{yen(s['unit'])}</td></tr>"
                for s in rep["stores"]
                if "prev" in s
            )
            + f"<tr><td>合計</td><td>{yen(t['sales'] - rep['prev_total']['sales'])}</td>"
            f"<td>{yen(vol)}</td><td>{yen(price)}</td>"
            f"<td>{yen(rep['prev_total']['unit'])}</td><td>{yen(t['unit'])}</td></tr>"
            "</tbody></table></div>"
            '<p class="note">組数寄与＝（今月組数−前月組数）×前月組単価、単価寄与＝（今月組単価−前月組単価）×今月組数。</p></section>'
        )

    parts.append(
        '<section><div class="head"><h2>店舗別</h2></div><div class="scroll"><table>'
        "<thead><tr><th>店舗</th><th>売上</th><th>前月比</th><th>営業日</th><th>日商</th>"
        "<th>組数</th><th>組単価</th><th>1組点数</th><th>キャッシュレス</th><th>人数入力率</th></tr></thead><tbody>"
        + "".join(
            f"<tr><td>{e(s['name'])}</td><td>{yen(s['sales'])}</td>"
            f"<td>{chip(s.get('mom', 0))}</td><td>{s['days']}</td><td>{yen(s['daily'])}</td>"
            f"<td>{s['groups']:,}</td><td>{yen(s['unit'])}</td><td>{s['items_per_group']:.1f}</td>"
            f"<td>{s['cashless'] * 100:.1f}%</td><td>{s['guest_input'] * 100:.0f}%</td></tr>"
            for s in rep["stores"]
        )
        + "</tbody></table></div>"
        '<p class="note">人数入力率＝Airレジの「客数」÷「組数」。100%未満のぶんは会計時に人数が入力されておらず、'
        "その場合の客単価は1組あたりの金額になります。</p></section>"
    )

    wk_used = [w for w in WEEKDAYS if any(w in s["weekday"] for s in rep["stores"])]
    parts.append(
        '<section><div class="head"><h2>曜日別・1営業日あたり売上</h2></div><div class="scroll"><table>'
        "<thead><tr><th>店舗</th>"
        + "".join(f"<th>{w}</th>" for w in wk_used)
        + "<th>最強 / 最弱</th></tr></thead><tbody>"
        + "".join(
            f"<tr><td>{e(s['name'])}</td>"
            + "".join(
                f"<td>{yen(s['weekday'][w])}</td>" if w in s["weekday"] else "<td>-</td>"
                for w in wk_used
            )
            + (
                f"<td>{max(s['weekday'], key=s['weekday'].get)} / "
                f"{min(s['weekday'], key=s['weekday'].get)}　"
                f"{max(s['weekday'].values()) / min(s['weekday'].values()):.1f}倍</td>"
                if s["weekday"]
                else "<td>-</td>"
            )
            + "</tr>"
            for s in rep["stores"]
        )
        + "</tbody></table></div></section>"
    )

    parts.append(
        '<section><div class="head"><h2>日別推移</h2></div><div class="split" id="dailies"></div>'
        '<div class="legend"><span><b style="background:var(--accent)"></b>土日</span>'
        '<span><b style="background:var(--accent-soft)"></b>平日</span>'
        "<span>休業日は表示なし</span></div></section>"
    )

    losses = [s for s in rep["stores"] if s["closed"]]
    if losses:
        parts.append(
            '<section><div class="head"><h2>臨時休業と機会損失の目安</h2></div><div class="scroll"><table>'
            "<thead><tr><th>店舗</th><th>休業日</th><th>日数</th><th>推定損失</th></tr></thead><tbody>"
            + "".join(
                f"<tr><td>{e(s['name'])}</td><td>"
                + "、".join(f"{d.month}/{d.day}({WEEKDAYS[d.weekday()]})" for d in s["closed"])
                + f"</td><td>{len(s['closed'])}</td><td>{yen(s['loss'])}</td></tr>"
                for s in losses
            )
            + f"<tr><td>合計</td><td></td><td>{sum(len(s['closed']) for s in losses)}</td>"
            f"<td>{yen(sum(s['loss'] for s in losses))}</td></tr>"
            "</tbody></table></div>"
            '<p class="note">他店が営業していた日に休んでいた日を臨時休業として集計し、その店の同曜日平均で損失を試算しています。</p></section>'
        )

    parts.append(
        "<footer><p>Airレジ「日別売上」CSVから自動生成。金額は税込・値引後。</p></footer>"
    )

    data = [
        {
            "name": s["name"],
            "rows": [[f"{r['date'].month}/{r['date'].day}", r["dow"], r["sales"]] for r in s["rows"]],
        }
        for s in rep["stores"]
    ]
    return (
        HTML_TEMPLATE.replace("__TITLE__", html.escape(title))
        .replace("__BODY__", "\n".join(p for p in parts if p))
        .replace("__DATA__", json.dumps(data, ensure_ascii=False))
    )


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Airレジ日別売上CSVから月次レポートを作る")
    ap.add_argument("csv_dir", type=Path, help="対象月のCSVを入れたディレクトリ")
    ap.add_argument("--prev", type=Path, help="前月のCSVディレクトリ（前月比・増減分解に使用）")
    ap.add_argument("--out", type=Path, help="HTMLの出力先（省略時はコンソールのみ）")
    ap.add_argument("--month", help="タイトルに使う年月（例 2026-07／省略時はデータから判定）")
    args = ap.parse_args(argv)

    stores = load_dir(args.csv_dir)
    prev = load_dir(args.prev) if args.prev else None
    rep = build(stores, prev)
    month = args.month or rep["month"] or ""
    print_console(rep)

    if args.out:
        y, _, m = month.partition("-")
        title = f"{y}年{int(m)}月 売上レポート" if m.isdigit() else "売上レポート"
        args.out.write_text(render_html(rep, title), encoding="utf-8")
        print(f"\nHTMLを書き出しました: {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
