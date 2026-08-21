#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Driveの .xlsx をネイティブのスプレッドシートにする経路の調査記録

    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 convert_probe.py

--- なぜ調べたか ---------------------------------------------------------
.xlsx のままDriveに置くと、ブラウザで開くたびにGoogleがその場でスプレッドシートへ
変換する（DOCS_EVERYWHERE_IMPORT）。13シート3万セルの損益計算書だとこれが重く、
2026-08-21 に QUOTA_EXCEEDED / TRANSIENT_FAILURE で開けなくなった。
ネイティブ化すれば変換が要らなくなるので、自動でやれないか試した。

--- 結果（2026-08-21 実測）★どちらも駄目だった ---------------------------
① files.update に body={"mimeType": スプレッドシート} ＋ xlsx を渡して変換
     → 400 invalidContentType
       "Invalid MIME type provided for the uploaded content."
     v3 の update は形式変換に対応していない。
     ★「IDを保ったまま変換できる」という当初の見立ては誤り。

② files.copy に body={"mimeType": スプレッドシート} を渡して変換コピー
     → 403 storageQuotaExceeded
       "The user\'s Drive storage quota has been exceeded."
     サービスアカウント自身には保存容量が無い。既存ファイルの【変更】はできるが、
     【新規作成】はできない。共有ドライブでないと回避できない。

--- つまり ---------------------------------------------------------------
ネイティブ化は、利用者がDriveの画面で
「ファイル → Google スプレッドシートとして保存」するしかない。
そのとき新しいファイルができるので【ファイルIDは変わる】。
変えたら push.py の TARGETS を新しいIDに差し替えること。

--- まだ確かめていない ★ここが肝心 ---------------------------------------
ネイティブのスプレッドシートに対して、push.py の丸ごと差し替え
（files.update に xlsx を送る）が通るのかどうか。
通らないなら、ネイティブ化と引き換えに自動転記が止まる。
確かめるには、利用者にネイティブのシートを1枚作って共有してもらう必要がある。

このスクリプトは①②を再現するためのもの。作ったファイルは必ずゴミ箱へ入れる。
"""
import os
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GSHEET = "application/vnd.google-apps.spreadsheet"
SRC = "1DGjxwNznNyXbRu1ovgMtca5zSKQyN0er"      # 22期
LOCAL = "損益計算書_22期.xlsx"
F = "id,name,mimeType,modifiedTime,parents"


def main():
    key = os.environ.get("GOOGLE_SA_KEY_FILE")
    if not key or not os.path.exists(key):
        sys.exit("GOOGLE_SA_KEY_FILE を指定してください")
    cred = service_account.Credentials.from_service_account_file(
        key, scopes=["https://www.googleapis.com/auth/drive"])
    svc = build("drive", "v3", credentials=cred, cache_discovery=False)

    src = svc.files().get(fileId=SRC, fields=F).execute()
    print(f"元: {src['name']}  {src['mimeType']}")

    new_id = None
    try:
        # ① copy で変換できるか
        print("\n① files.copy で変換 …", end=" ")
        made = svc.files().copy(
            fileId=SRC,
            body={"name": "【変換テスト】損益計算書_22期",
                  "mimeType": GSHEET,
                  "parents": src.get("parents", [])},
            fields=F,
        ).execute()
        new_id = made["id"]
        ok1 = made["mimeType"] == GSHEET
        print("成功" if ok1 else "★変換されず")
        print(f"     新ID {new_id}")
        print(f"     形式 {made['mimeType']}")
        if not ok1:
            return

        # ② できたネイティブのシートに xlsx を丸ごと上書きできるか
        print("\n② ネイティブのシートに xlsx を files.update …", end=" ")
        try:
            after = svc.files().update(
                fileId=new_id,
                media_body=MediaFileUpload(LOCAL, mimetype=XLSX, resumable=True),
                fields=F,
            ).execute()
            kept = after["mimeType"] == GSHEET
            print("成功" if kept else "★成功したがExcelに戻った")
            print(f"     形式 {after['mimeType']}")
            print("\n→ 判定: ネイティブ化しても push.py の丸ごと差し替えを続けられる"
                  if kept else
                  "\n→ 判定: 上書きするたびExcelに戻る。運用が続かない")
        except HttpError as e:
            print("★失敗")
            print(f"     {e.status_code} {e.reason}")
            print("\n→ 判定: ネイティブ化すると push.py の丸ごと差し替えが使えなくなる")
    finally:
        if new_id:
            svc.files().update(fileId=new_id, body={"trashed": True}).execute()
            print(f"\nテスト用ファイル {new_id} はゴミ箱に入れました")


if __name__ == "__main__":
    main()
