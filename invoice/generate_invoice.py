#!/usr/bin/env python3
"""請求書PDFを生成するスクリプト。

大衆酒場りゅうちゃん 店長 河野竜二氏（個人事業主）からProstyle株式会社への
月次業務委託料の請求書を、invoice/config.json と invoice/template.html を元に
PDFとして出力する。レンダリングにはヘッドレスChromiumを使用する。

使い方:
    python3 invoice/generate_invoice.py            # 当月（JST）分を生成
    python3 invoice/generate_invoice.py --year 2026 --month 8
    python3 invoice/generate_invoice.py -o /path/to/out.pdf
"""

import argparse
import calendar
import datetime
import json
import os
import shutil
import subprocess
import sys
import tempfile
from zoneinfo import ZoneInfo

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
JST = ZoneInfo("Asia/Tokyo")


def ensure_cjk_font():
    """日本語フォント（Noto Sans CJK）が無い環境ではPDFが肥大化するため、可能なら導入する。"""
    try:
        out = subprocess.run(["fc-list", ":lang=ja", "family"], capture_output=True, text=True, timeout=30)
        if "IPAexGothic" in out.stdout:
            return
        # 単体TTFのフォントはType0で埋め込まれPDFが小さくなる（TTC形式はType3化して肥大する）
        subprocess.run(["apt-get", "install", "-y", "fonts-ipaexfont-gothic"],
                       capture_output=True, timeout=300)
    except Exception:
        pass  # フォントが無くても生成自体は可能（サイズが大きくなるだけ）


def find_chromium():
    candidates = [os.environ.get("CHROME_BIN")]
    pw_root = os.environ.get("PLAYWRIGHT_BROWSERS_PATH", "/opt/pw-browsers")
    candidates.append(os.path.join(pw_root, "chromium"))
    if os.path.isdir(pw_root):
        for entry in sorted(os.listdir(pw_root), reverse=True):
            candidates.append(os.path.join(pw_root, entry, "chrome-linux", "chrome"))
            candidates.append(os.path.join(pw_root, entry, "chrome-linux", "headless_shell"))
    for name in ("chromium", "chromium-browser", "google-chrome", "google-chrome-stable"):
        candidates.append(shutil.which(name))
    for path in candidates:
        if path and os.path.isfile(path) and os.access(path, os.X_OK):
            return path
    raise FileNotFoundError("Chromiumが見つかりません。CHROME_BIN で実行ファイルを指定してください。")


def build_html(config, year, month):
    last_day = calendar.monthrange(year, month)[1]
    issue_date = datetime.date(year, month, last_day)
    template_path = os.path.join(BASE_DIR, "template.html")
    with open(template_path, encoding="utf-8") as f:
        html = f.read()
    issuer = config["issuer"]
    replacements = {
        "{{INVOICE_NO}}": f"RY-{year:04d}{month:02d}",
        "{{ISSUE_DATE}}": f"{issue_date.year}年{issue_date.month}月{issue_date.day}日",
        "{{RECIPIENT}}": config["recipient"],
        "{{ISSUER_STORE}}": issuer["store"],
        "{{ISSUER_TITLE}}": issuer["title"],
        "{{ISSUER_NAME}}": issuer["name"],
        "{{ISSUER_NOTE}}": issuer["note"],
        "{{ISSUER_POSTAL}}": issuer.get("postal", ""),
        "{{ISSUER_ADDRESS}}": issuer.get("address", ""),
        "{{ISSUER_PHONE}}": issuer.get("phone", ""),
        "{{SEAL_CHAR1}}": issuer.get("seal_text", "")[:1],
        "{{SEAL_CHAR2}}": issuer.get("seal_text", "")[1:2],
        "{{LINE_ITEM}}": config["line_item"],
        "{{PERIOD_LABEL}}": f"{year}年{month}月",
        "{{TOTAL_FMT}}": f"{config['amount']:,}",
        "{{TAX_LABEL}}": "税込" if config.get("tax_included", True) else "税抜",
    }
    for key, value in replacements.items():
        html = html.replace(key, value)
    return html


def render_pdf(html, output_path):
    ensure_cjk_font()
    chromium = find_chromium()
    with tempfile.TemporaryDirectory() as tmpdir:
        html_path = os.path.join(tmpdir, "invoice.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html)
        cmd = [
            chromium,
            "--headless",
            "--disable-gpu",
            "--no-sandbox",
            "--no-pdf-header-footer",
            f"--print-to-pdf={output_path}",
            f"file://{html_path}",
        ]
        subprocess.run(cmd, check=True, capture_output=True, timeout=120)
    if not os.path.isfile(output_path) or os.path.getsize(output_path) == 0:
        raise RuntimeError("PDFの生成に失敗しました。")
    try:  # pikepdfがあれば再圧縮してサイズを削減（無くても可）
        import pikepdf
        tmp = output_path + ".opt"
        with pikepdf.open(output_path) as pdf:
            pdf.save(tmp, compress_streams=True,
                     object_stream_mode=pikepdf.ObjectStreamMode.generate,
                     recompress_flate=True)
        os.replace(tmp, output_path)
    except Exception:
        pass


def main():
    parser = argparse.ArgumentParser(description="月次請求書PDFを生成")
    now = datetime.datetime.now(JST)
    parser.add_argument("--year", type=int, default=now.year)
    parser.add_argument("--month", type=int, default=now.month)
    parser.add_argument("-o", "--output", default=None)
    args = parser.parse_args()

    with open(os.path.join(BASE_DIR, "config.json"), encoding="utf-8") as f:
        config = json.load(f)

    output = args.output or os.path.join(
        BASE_DIR, "output", f"seikyusho_{args.year:04d}{args.month:02d}_kono.pdf"
    )
    os.makedirs(os.path.dirname(output), exist_ok=True)

    html = build_html(config, args.year, args.month)
    render_pdf(html, output)
    print(output)


if __name__ == "__main__":
    main()
