#!/usr/bin/env python3
"""メール添付用のコンパクト請求書PDFを生成するスクリプト。

reportlab + 非埋め込みCIDフォント（HeiseiKakuGo-W5）で生成するため、
1ページあたり数KBと非常に軽い。閲覧側ではビューアの日本語代替フォントで
表示される（Acrobat/Chrome/macOSプレビュー等の主要ビューアで表示可能）。

使い方:
    python3 invoice/generate_invoice_compact.py                     # 当月分
    python3 invoice/generate_invoice_compact.py --year 2026 --month 8
    python3 invoice/generate_invoice_compact.py --months 2025-09:2026-07 -o out.pdf  # 期間一括（複数ページ）
"""

import argparse
import calendar
import datetime
import json
import os
from zoneinfo import ZoneInfo

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfgen import canvas

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JST = ZoneInfo("Asia/Tokyo")
FONT = "HeiseiKakuGo-W5"

INK = HexColor("#1a1a18")
GRAY = HexColor("#555555")
BORDER = HexColor("#999999")
HEAD_BG = HexColor("#f0f0ee")
SEAL_RED = HexColor("#c93a2f")

W, H = A4  # 595.27 x 841.89 pt


TEMPLATE_FORM = "invoiceTemplate"

# 明細テーブルの座標（共通部分と可変部分で同じ値を使う）
TABLE_X = 20 * mm
TABLE_TOP = H - 140 * mm
ROW_H = 11 * mm


def draw_variable_text(c, config, year, month):
    """年月に依存する文言だけを描画する（請求書番号・発行日・摘要の対象月）。"""
    last_day = calendar.monthrange(year, month)[1]
    c.setFont(FONT, 9)
    c.setFillColor(GRAY)
    c.drawRightString(W - 20 * mm, H - 47 * mm, f"請求書番号：RY-{year:04d}{month:02d}")
    c.drawRightString(W - 20 * mm, H - 52 * mm, f"発行日：{year}年{month}月{last_day}日")
    c.setFillColor(INK)
    c.setFont(FONT, 10)
    c.drawString(TABLE_X + 4 * mm, TABLE_TOP - 2 * ROW_H + 4 * mm,
                 f"{config['line_item']}（{year}年{month}月分）")


def draw_template_form(c, config):
    """毎ページ共通の枠・固定文言をForm XObjectとして一度だけ定義する。

    複数ページ生成時、各ページはこのXObjectを参照するだけで済むため、
    ページ数が増えてもファイルサイズがほとんど増えない（メール添付時の
    base64サイズを抑える目的）。
    """
    c.beginForm(TEMPLATE_FORM)
    draw_invoice_page(c, config, None, None, static_only=True)
    c.endForm()


def draw_invoice_page(c, config, year, month, static_only=False, use_form=False):
    """請求書1ページ分を描画する。

    static_only=True: 年月に依存しない共通部分のみ描画（テンプレート定義用）。
    use_form=True: 共通部分はForm XObjectを参照し、年月依存の文言のみ描画。
    """
    issuer = config["issuer"]
    amount = config["amount"]
    amount_fmt = f"￥{amount:,}"
    tax_label = "税込" if config.get("tax_included", True) else "税抜"

    if use_form:
        c.doForm(TEMPLATE_FORM)
        draw_variable_text(c, config, year, month)
        c.showPage()
        return

    c.setFillColor(INK)
    c.setStrokeColor(INK)

    # タイトル
    c.setFont(FONT, 22)
    title = "請　求　書"
    c.drawCentredString(W / 2, H - 30 * mm, title)

    # 右上メタ（年月依存のため、テンプレート定義時は描かない）
    if not static_only:
        draw_variable_text(c, config, year, month)

    # 宛先（左）
    c.setFillColor(INK)
    c.setFont(FONT, 13)
    to_y = H - 65 * mm
    c.drawString(20 * mm, to_y, f"{config['recipient']} 御中")
    c.setLineWidth(1)
    c.line(20 * mm, to_y - 2 * mm, 100 * mm, to_y - 2 * mm)
    c.setFont(FONT, 9)
    c.setFillColor(GRAY)
    c.drawString(20 * mm, to_y - 8 * mm, "下記のとおりご請求申し上げます。")

    # 発行者（右）
    c.setFillColor(INK)
    fx = 115 * mm
    fy = H - 63 * mm
    c.setFont(FONT, 10)
    c.drawString(fx, fy, issuer["store"])
    c.setFont(FONT, 9)
    c.drawString(fx, fy - 5 * mm, f"{issuer['title']}　{issuer['name']}（{issuer['note']}）")
    c.setFillColor(GRAY)
    c.setFont(FONT, 8.5)
    c.drawString(fx, fy - 10 * mm, f"〒{issuer.get('postal', '')}")
    c.drawString(fx, fy - 14.5 * mm, issuer.get("address", ""))
    c.drawString(fx, fy - 19 * mm, f"TEL：{issuer.get('phone', '')}")

    # 印影
    seal = issuer.get("seal_text", "")
    if seal:
        cx, cy, r = W - 26 * mm, fy - 2 * mm, 7.5 * mm
        c.setStrokeColor(SEAL_RED)
        c.setLineWidth(1.4)
        c.circle(cx, cy, r, stroke=1, fill=0)
        c.setFillColor(SEAL_RED)
        c.setFont(FONT, 13)
        chars = list(seal[:2])
        if len(chars) == 2:
            c.drawCentredString(cx, cy + 0.8 * mm, chars[0])
            c.drawCentredString(cx, cy - 5.2 * mm, chars[1])
        else:
            c.drawCentredString(cx, cy - 2 * mm, seal)

    # 金額ボックス
    c.setStrokeColor(INK)
    bx, by, bw, bh = 20 * mm, H - 122 * mm, 115 * mm, 18 * mm
    c.setLineWidth(1.2)
    c.rect(bx, by, bw, bh)
    c.setFillColor(INK)
    c.setFont(FONT, 11)
    c.drawString(bx + 8 * mm, by + 6.5 * mm, "ご請求金額")
    c.setFont(FONT, 18)
    c.drawString(bx + 38 * mm, by + 6 * mm, amount_fmt)
    c.setFont(FONT, 9)
    c.setFillColor(GRAY)
    c.drawString(bx + 82 * mm, by + 6.5 * mm, tax_label)

    # 明細テーブル
    tx, tw = 20 * mm, W - 40 * mm
    ty = H - 140 * mm
    row_h = 11 * mm
    col_qty = tx + tw - 60 * mm
    col_amt = tx + tw - 40 * mm

    c.setStrokeColor(BORDER)
    c.setLineWidth(0.7)
    # ヘッダ行
    c.setFillColor(HEAD_BG)
    c.rect(tx, ty - row_h, tw, row_h, stroke=1, fill=1)
    c.setFillColor(INK)
    c.setFont(FONT, 9.5)
    c.drawCentredString(tx + (col_qty - tx) / 2, ty - row_h + 4 * mm, "摘要")
    c.drawCentredString(col_qty + 10 * mm, ty - row_h + 4 * mm, "数量")
    c.drawCentredString(col_amt + 20 * mm, ty - row_h + 4 * mm, "金額")
    # 明細行
    y1 = ty - 2 * row_h
    c.rect(tx, y1, tw, row_h, stroke=1, fill=0)
    c.setFont(FONT, 10)
    c.drawCentredString(col_qty + 10 * mm, y1 + 4 * mm, "1式")
    c.drawRightString(tx + tw - 4 * mm, y1 + 4 * mm, amount_fmt)
    # 合計行
    y2 = ty - 3 * row_h
    c.setFillColor(HexColor("#fafaf8"))
    c.rect(tx, y2, tw, row_h, stroke=1, fill=1)
    c.setFillColor(INK)
    c.drawString(tx + 4 * mm, y2 + 4 * mm, f"合計（{tax_label}）")
    c.drawRightString(tx + tw - 4 * mm, y2 + 4 * mm, amount_fmt)
    # 縦罫線
    c.line(col_qty, ty - row_h, col_qty, y1)
    c.line(col_amt, ty - row_h, col_amt, y1)

    # 備考
    ny = y2 - 12 * mm
    nh = 30 * mm
    c.setStrokeColor(HexColor("#cccccc"))
    c.rect(tx, ny - nh, tw, nh, stroke=1, fill=0)
    c.setFont(FONT, 9)
    c.drawString(tx + 5 * mm, ny - 7 * mm, "備考")
    c.setFillColor(GRAY)
    c.setFont(FONT, 8.5)
    c.drawString(tx + 5 * mm, ny - 13 * mm,
                 "※ 発行者は適格請求書発行事業者（インボイス制度）の登録を行っていないため、")
    c.drawString(tx + 5 * mm, ny - 17.5 * mm,
                 "　 本書は適格請求書には該当しません。")
    c.drawString(tx + 5 * mm, ny - 23 * mm,
                 f"※ 本書は支払側（{config['recipient']}）にて作成した支払明細を兼ねる書類です。")
    c.drawString(tx + 5 * mm, ny - 27.5 * mm,
                 "　 記載内容に相違がある場合はお申し出ください。")

    if not static_only:  # テンプレート定義中はページを閉じない
        c.showPage()


def parse_months(spec):
    """"2025-09:2026-07" → [(2025,9), ..., (2026,7)]"""
    start_s, end_s = spec.split(":")
    sy, sm = (int(x) for x in start_s.split("-"))
    ey, em = (int(x) for x in end_s.split("-"))
    months = []
    y, m = sy, sm
    while (y, m) <= (ey, em):
        months.append((y, m))
        m += 1
        if m > 12:
            y, m = y + 1, 1
    return months


def main():
    parser = argparse.ArgumentParser(description="コンパクト版 月次請求書PDFを生成")
    now = datetime.datetime.now(JST)
    parser.add_argument("--year", type=int, default=now.year)
    parser.add_argument("--month", type=int, default=now.month)
    parser.add_argument("--months", default=None, help="期間一括 例: 2025-09:2026-07")
    parser.add_argument("-o", "--output", default=None)
    args = parser.parse_args()

    with open(os.path.join(BASE_DIR, "config.json"), encoding="utf-8") as f:
        config = json.load(f)

    months = parse_months(args.months) if args.months else [(args.year, args.month)]
    if args.output:
        output = args.output
    elif len(months) == 1:
        output = os.path.join(BASE_DIR, "output",
                              f"seikyusho_{months[0][0]:04d}{months[0][1]:02d}_kono_c.pdf")
    else:
        output = os.path.join(BASE_DIR, "output",
                              f"seikyusho_{months[0][0]:04d}{months[0][1]:02d}-"
                              f"{months[-1][0]:04d}{months[-1][1]:02d}_kono_c.pdf")
    os.makedirs(os.path.dirname(output), exist_ok=True)

    pdfmetrics.registerFont(UnicodeCIDFont(FONT))
    c = canvas.Canvas(output, pagesize=A4)
    c.setTitle(f"請求書 {config['issuer']['store']}")
    if len(months) > 1:
        # 複数月はテンプレートを共有してファイルサイズを抑える
        draw_template_form(c, config)
        for y, m in months:
            draw_invoice_page(c, config, y, m, use_form=True)
    else:
        for y, m in months:
            draw_invoice_page(c, config, y, m)
    c.save()
    print(output)


if __name__ == "__main__":
    main()
