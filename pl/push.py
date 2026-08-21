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
5. JSONを Dropbox の /※請求書※ 直下に置く
   実物: sa-key.json.json （file_id: id:BKWmtQDznicAAAAAAAIh2g）
   サービスアカウント: pl-writer@prostyle-pl.iam.gserviceaccount.com

--- 疎通確認済み（2026-08-20）-------------------------------------------
21期・22期とも反映成功。Drive側のファイルサイズがローカルと完全一致した。
つまずいたのは1点だけ: プロジェクトで Google Drive API が無効だった
（403 accessNotConfigured）。Sheets API は使わないので有効化は不要。

--- 鍵の受け渡し（毎セッション）-----------------------------------------
コンテナはセッションごとに消えるので、毎回 Dropbox から取り直す。

    # Dropbox の download_link で一時URLを発行し、curl でディスクに落とす
    curl -sSL -o /tmp/sa.json "<一時URL>"
    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 push.py 21期

★この経路だと鍵の中身が会話ログを通らない。Driveコネクタの
  download_file_content は中身をアシスタントに返すので、この用途には使わない。
★クラウド環境変数にも入れないこと。公式ドキュメントが明確に禁じている
  （シークレット保管庫ではなく、環境を使う人は誰でも値を読める）。
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

# 期: (ローカルの.xlsx, DriveのファイルID)
# ファイルIDは秘密情報ではないので直接書いてよい。環境変数で上書きもできる。
TARGETS = {
    "21期": ("損益計算書_21期テスト版.xlsx",
            os.environ.get("SHEET_ID_21", "1i7R-v75cigfHjTBCVJ2dzBWOJxAf2Iww")),
    "22期": ("損益計算書_22期.xlsx",
            os.environ.get("SHEET_ID_22", "1DGjxwNznNyXbRu1ovgMtca5zSKQyN0er")),
}
SCOPES = ["https://www.googleapis.com/auth/drive"]
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
GSHEET = "application/vnd.google-apps.spreadsheet"
_KIND = {XLSX: "Excelファイル", GSHEET: "Googleスプレッドシート"}


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
    before = svc.files().get(fileId=file_id,
                             fields="name,mimeType,modifiedTime").execute()
    print(f"対象: {before['name']}  ({_KIND.get(before['mimeType'], before['mimeType'])}"
          f"／最終更新 {before['modifiedTime']})")

    # ★body に mimeType を入れて変換させることはできない。
    #   2026-08-21 に試したら 400 invalidContentType で弾かれた
    #   （"Invalid MIME type provided for the uploaded content."）。
    #   v3 の files.update は形式変換に対応していない。convert_probe.py に記録。
    #   なのでメディアだけ送る。上書きで形式が変わっていないことは下で確かめる。
    media = MediaFileUpload(local, mimetype=XLSX, resumable=True)
    after = svc.files().update(fileId=file_id, media_body=media,
                               fields="name,mimeType,modifiedTime,webViewLink").execute()
    assert after["mimeType"] == before["mimeType"], \
        f"形式が {_KIND.get(before['mimeType'], before['mimeType'])} から " \
        f"{_KIND.get(after['mimeType'], after['mimeType'])} に変わってしまった"
    print(f"反映しました: {after['name']}  → {after['modifiedTime']}")
    print(after["webViewLink"])


if __name__ == "__main__":
    if len(sys.argv) != 2 or sys.argv[1] not in TARGETS:
        sys.exit(f"使い方: python3 push.py [{' | '.join(TARGETS)}]")
    push(sys.argv[1])
