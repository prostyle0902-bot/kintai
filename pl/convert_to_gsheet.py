#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Drive上の .xlsx を、ファイルIDを保ったままGoogleスプレッドシートに変換する

    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 convert_to_gsheet.py 22期
    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 convert_to_gsheet.py 22期 --dry-run

--- なぜ要るか -----------------------------------------------------------
.xlsx のままDriveに置くと、ブラウザで開くたびにGoogleがその場で
スプレッドシートへ変換する（DOCS_EVERYWHERE_IMPORT）。13シート・3万セルの
損益計算書だとこれが重く、2026-08-21 に QUOTA_EXCEEDED / TRANSIENT_FAILURE で
開けなくなった。ネイティブのスプレッドシートにしてしまえば変換が要らなくなる。

--- 仕組み ---------------------------------------------------------------
files.update に
    body       = {"mimeType": "application/vnd.google-apps.spreadsheet"}  変換後
    media_body = xlsx                                                     送るもの
を渡すと、Driveが取り込み時に変換する。★ファイルIDは変わらない。
IDが変わらないので push.py の TARGETS もURLもそのまま使える。

--- 変換したあとの push.py ------------------------------------------------
ネイティブ化したあとも、同じやり方（body に mimeType を入れる）で
上書きし続けられる。push.py 側も同時に直すこと。
xlsxのまま上げ直すと、またExcelファイルに戻ってしまう。

--- 注意 -----------------------------------------------------------------
・変換は元に戻せない。戻したいときはローカルの .xlsx を上げ直す
・Driveで手入力した内容は消える（push.py と同じく丸ごと差し替えのため）
"""
import os
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GSHEET = "application/vnd.google-apps.spreadsheet"
FIELDS = "id,name,mimeType,size,modifiedTime,parents,webViewLink"

TARGETS = {
    "21期": ("損益計算書_21期テスト版.xlsx",
             os.environ.get("SHEET_ID_21", "1i7R-v75cigfHjTBCVJ2dzBWOJxAf2Iww")),
    "22期": ("損益計算書_22期.xlsx",
             os.environ.get("SHEET_ID_22", "1DGjxwNznNyXbRu1ovgMtca5zSKQyN0er")),
}


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dry = "--dry-run" in sys.argv
    if len(args) != 1 or args[0] not in TARGETS:
        sys.exit(f"使い方: python3 {sys.argv[0]} <{'|'.join(TARGETS)}> [--dry-run]")
    path, fid = TARGETS[args[0]]

    key = os.environ.get("GOOGLE_SA_KEY_FILE")
    if not key or not os.path.exists(key):
        sys.exit("GOOGLE_SA_KEY_FILE にサービスアカウントのJSONを指定してください")
    if not os.path.exists(path):
        sys.exit(f"{path} がありません。build22.py / fill2.py で作ってください")

    cred = service_account.Credentials.from_service_account_file(
        key, scopes=["https://www.googleapis.com/auth/drive"])
    svc = build("drive", "v3", credentials=cred)

    before = svc.files().get(fileId=fid, fields=FIELDS).execute()
    print("【変換前】")
    for k in ("name", "mimeType", "size", "modifiedTime"):
        print(f"   {k:<14}{before.get(k)}")
    if before["mimeType"] == GSHEET:
        print("\nすでにGoogleスプレッドシートです。何もしません。")
        return
    if dry:
        print("\n--dry-run なのでここまで。実行するとIDはそのままで形式だけ変わります。")
        return

    svc.files().update(
        fileId=fid,
        body={"mimeType": GSHEET},
        media_body=MediaFileUpload(path, mimetype=XLSX, resumable=True),
        fields=FIELDS,
    ).execute()

    after = svc.files().get(fileId=fid, fields=FIELDS).execute()
    print("\n【変換後】")
    for k in ("name", "mimeType", "size", "modifiedTime"):
        print(f"   {k:<14}{after.get(k)}")
    print(f"\n   ファイルID   {'変わっていない' if after['id'] == fid else '★変わった'}  {after['id']}")
    print(f"   親フォルダ    {'同じ' if before.get('parents') == after.get('parents') else '★変わった'}")
    print(f"   形式         {'ネイティブのスプレッドシートになった' if after['mimeType'] == GSHEET else '★変換されていない'}")
    print(f"\n{after.get('webViewLink')}")


if __name__ == "__main__":
    main()
