#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成した .xlsx を Google Drive のスプレッドシートへ直接反映する

これがあると「ダウンロード → Driveにアップロード → 変換」の手作業がなくなる。

    python3 push.py 21期
    python3 push.py 22期

--- 準備（1回だけ）-------------------------------------------------------
1. Google Cloud Console でプロジェクトを作る
2. Google Sheets API と Google Drive API を有効化
3. サービスアカウントを作成し、JSONキーを発行
4. JSONの client_email（〜@〜.iam.gserviceaccount.com）に、対象フォルダを
   「編集者」で共有する
5. JSONの中身を環境変数 GOOGLE_SA_KEY に入れる
   （ファイルで渡す場合は GOOGLE_SA_KEY_FILE にパスを入れる）

★鍵はリポジトリにコミットしないこと。.gitignore で sa*.json を除外している。

--- 仕組み ---------------------------------------------------------------
Drive の files.update に xlsx を投げると、既存のスプレッドシートの中身が
まるごと差し替わる（タブ・数式・書式ごと）。Python がファイルをディスクから
読んで送るので、サイズの制約を受けない。

★注意: 差し替えなので、Drive 側で手入力した内容は消える。
        手入力するなら別ファイルに分けること。
"""
import json, os, sys

# 必要: pip install google-api-python-client google-auth

TARGETS = {
    # 期: (ローカルの.xlsx, DriveのファイルID)
    #     ファイルIDは初回アップロード後に search_files などで調べて埋める
    "21期": ("損益計算書_21期テスト版.xlsx", os.environ.get("SHEET_ID_21")),
    "22期": ("損益計算書_22期.xlsx",        os.environ.get("SHEET_ID_22")),
}
SCOPES = ["https://www.googleapis.com/auth/drive"]
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def _creds():
    from google.oauth2 import service_account
    raw = os.environ.get("GOOGLE_SA_KEY")
    if raw:
        return service_account.Credentials.from_service_account_info(
            json.loads(raw), scopes=SCOPES)
    path = os.environ.get("GOOGLE_SA_KEY_FILE")
    if path and os.path.exists(path):
        return service_account.Credentials.from_service_account_file(path, scopes=SCOPES)
    sys.exit("GOOGLE_SA_KEY も GOOGLE_SA_KEY_FILE も設定されていません。"
             "push.py の冒頭の手順を参照してください。")


def push(period):
    local, file_id = TARGETS[period]
    if not file_id:
        sys.exit(f"{period} の DriveファイルIDが未設定です（環境変数 SHEET_ID_*）。"
                 f"初回だけ手でアップロードし、そのファイルIDを設定してください。")
    if not os.path.exists(local):
        sys.exit(f"{local} がありません。先に fill2.py / build22.py を実行してください。")

    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    svc = build("drive", "v3", credentials=_creds(), cache_discovery=False)
    before = svc.files().get(fileId=file_id, fields="name,modifiedTime").execute()
    print(f"対象: {before['name']}  (最終更新 {before['modifiedTime']})")

    media = MediaFileUpload(local, mimetype=XLSX, resumable=True)
    after = svc.files().update(fileId=file_id, media_body=media,
                               fields="name,modifiedTime,webViewLink").execute()
    print(f"反映しました: {after['name']}  → {after['modifiedTime']}")
    print(after["webViewLink"])


if __name__ == "__main__":
    if len(sys.argv) != 2 or sys.argv[1] not in TARGETS:
        sys.exit(f"使い方: python3 push.py [{' | '.join(TARGETS)}]")
    push(sys.argv[1])
