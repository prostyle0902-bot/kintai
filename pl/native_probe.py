#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ネイティブのスプレッドシートに xlsx を丸ごと差し替えできるか試す（調査用）

    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 native_probe.py <ファイルID>

--- 何を確かめたいか -----------------------------------------------------
損益計算書を .xlsx のままDriveに置くと、開くたびにGoogleが変換するので重い
（convert_probe.py 参照）。ネイティブのスプレッドシートにすれば解決するが、
そのあとも push.py の丸ごと差し替えを続けられないと自動転記が止まる。
そこが分からないので、捨てて構わないネイティブのシートで先に試す。

--- 試すこと -------------------------------------------------------------
① files.update に xlsx のメディアだけ送る（push.py と同じやり方）
② ①が駄目なら body={"mimeType": スプレッドシート} も付けて送る
どちらかが通り、かつ形式がスプレッドシートのまま保たれれば運用が続けられる。

--- 結果（2026-08-21 実測。利用者が作った空のシートで検証）---------------
★①②とも 500 Internal Error。3回繰り返しても同じなので一時障害ではない。
  サービスアカウントの権限は writer で、権限の問題でもない（403ではなく500）。
  → ネイティブのスプレッドシートに xlsx を丸ごと差し替えることはできない。
    ネイティブ化すると push.py のやり方は使えなくなる。

一方で Sheets API なら全部できた:
    値の書き込み            ✅ 9セル
    数式                    ✅ =B2*2 が 24690 になった
    数値書式（赤字マイナス） ✅ "#,##0;[Red]-#,##0"
    背景色                  ✅
    タブの追加              ✅
  → ネイティブ化するなら、push.py を Sheets API 方式に作り替える必要がある。
    ローカルの .xlsx を openpyxl で読んで、セルと書式をシートへ写す形にすれば
    fill2.py も build2.py もそのまま使える。
"""
import os
import sys

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GSHEET = "application/vnd.google-apps.spreadsheet"
LOCAL = "損益計算書_22期.xlsx"
F = "id,name,mimeType,modifiedTime"


def _svc():
    key = os.environ.get("GOOGLE_SA_KEY_FILE")
    if not key or not os.path.exists(key):
        sys.exit("GOOGLE_SA_KEY_FILE を指定してください")
    cred = service_account.Credentials.from_service_account_file(
        key, scopes=["https://www.googleapis.com/auth/drive",
                     "https://www.googleapis.com/auth/spreadsheets"])
    return cred


def main():
    if len(sys.argv) != 2:
        sys.exit(f"使い方: python3 {sys.argv[0]} <ファイルID>")
    fid = sys.argv[1]
    cred = _svc()
    drive = build("drive", "v3", credentials=cred, cache_discovery=False)

    try:
        before = drive.files().get(fileId=fid, fields=F).execute()
    except HttpError as e:
        sys.exit(f"読めません（共有されていない可能性）: {e.status_code} {e.reason}")
    print("【差し替え前】")
    for k in ("name", "mimeType", "modifiedTime"):
        print(f"   {k:<14}{before.get(k)}")
    if before["mimeType"] != GSHEET:
        sys.exit("\nこのファイルはネイティブのスプレッドシートではありません")

    ok = False
    for label, body in (("① メディアだけ送る（push.py と同じ）", {}),
                        ("② body に mimeType も付ける", {"mimeType": GSHEET})):
        print(f"\n{label} …", end=" ")
        try:
            drive.files().update(
                fileId=fid, body=body,
                media_body=MediaFileUpload(LOCAL, mimetype=XLSX, resumable=True),
                fields=F,
            ).execute()
            print("成功")
            ok = True
            break
        except HttpError as e:
            print(f"失敗  {e.status_code} {getattr(e, 'reason', '')}")
    if not ok:
        print("\n→ 判定: ネイティブ化すると push.py の丸ごと差し替えが使えない")
        return

    after = drive.files().get(fileId=fid, fields=F).execute()
    print("\n【差し替え後】")
    for k in ("name", "mimeType", "modifiedTime"):
        print(f"   {k:<14}{after.get(k)}")
    kept = after["mimeType"] == GSHEET
    print(f"   形式          {'スプレッドシートのまま' if kept else '★Excelに戻った'}")
    if not kept:
        print("\n→ 判定: 上書きするたびExcelに戻る。運用が続かない")
        return

    # 中身が本当に入ったかを読み返す
    sheets = build("sheets", "v4", credentials=cred, cache_discovery=False)
    meta = sheets.spreadsheets().get(spreadsheetId=fid).execute()
    names = [s["properties"]["title"] for s in meta["sheets"]]
    print(f"\n   タブ {len(names)}枚: {names}")
    vals = sheets.spreadsheets().values().get(
        spreadsheetId=fid, range="神栖横丁!A1:D8").execute().get("values", [])
    print("   神栖横丁の左上:")
    for row in vals[:8]:
        print("     ", row)
    print("\n→ 判定: ネイティブ化しても push.py の丸ごと差し替えを続けられる")


if __name__ == "__main__":
    main()
