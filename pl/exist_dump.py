#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""既存の21期PLシート10個を Sheets API で読んで exist/<タブ>.json に落とす（読むだけ）

    GOOGLE_SA_KEY_FILE=/tmp/sa.json python3 exist_dump.py

★既存シートには絶対に書き込まない。values().get しか呼ばない。
"""
import json, os, sys

# 新シートのタブ名 → 既存スプシのファイルID
SHEETS = {
    "りゅうちゃん":     "1N5QF8UO_9z8MVZ0_Z1QBdAghfjjYtmvoXnpTTzLQs0E",
    # ★りゅうちゃんNEW（11J-zvL7…）はサービスアカウントに共有されていない（403）。
    #   2026-08-20更新でこちらのほうが新しい。必要なら利用者に共有してもらうこと。
    "もも焼きJAPAN":   "16a9y8ExDjpfn7K6iQg1-908GORIBZlrBJd71Qzvmw7g",
    "韓国酒場ハナ":     "1MbrqksGpf80EKzRAZggLBSDJ0rvRIZuEqOE7qUwjGow",
    "さわら十三里屋":   "1K1U0hqlJU5O0HJfg_Wijon3MeY7c5GuGKJ8KKfJVsms",
    "タコとハイボール": "10PZ-wKLXQArqEsmPRBN4pN6_ClYCs6MKShokT5w1rJ4",
    "焼きたて屋":       "1OB3yLta9y1d14DbN3jeqCoL9MKFknUbBnqR1WBOvLYQ",
    "神栖横丁":         "1aI4rq96FEHjS0GGhMoSTiMMPzHmrsR-VCqs9ipwObzA",
    "鳥害対策課":       "1q8C6SyuQ1-Iq3WKWfNlZHciDHkNyeS83S_2ANMEaluM",
    "業務課":           "1JKy9P6_nxhAHwoUlZEyuHGhrTZlySsEg6E6KzCaa5_g",
    "本部":             "1xO2MPtct0XSPVOC_d9TxRdZpTXGv9tqM6gewF8TxuYk",
}
OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "exist")


def main():
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
    cred = service_account.Credentials.from_service_account_file(
        os.environ["GOOGLE_SA_KEY_FILE"],
        scopes=["https://www.googleapis.com/auth/spreadsheets.readonly"])
    svc = build("sheets", "v4", credentials=cred, cache_discovery=False)
    os.makedirs(OUT, exist_ok=True)
    for tab, sid in SHEETS.items():
        try:
            meta = svc.spreadsheets().get(spreadsheetId=sid).execute()
        except HttpError as e:
            print(f"{tab:<16} 読めません: {e.status_code} {e.reason}")
            continue
        names = [s["properties"]["title"] for s in meta["sheets"]]
        got = svc.spreadsheets().values().batchGet(
            spreadsheetId=sid, ranges=[f"'{n}'!A1:Z1000" for n in names],
            valueRenderOption="UNFORMATTED_VALUE").execute()["valueRanges"]
        data = {n: g.get("values", []) for n, g in zip(names, got)}
        with open(os.path.join(OUT, f"{tab}.json"), "w") as f:
            json.dump({"title": meta["properties"]["title"], "sheets": data},
                      f, ensure_ascii=False)
        cells = sum(len(r) for v in data.values() for r in v)
        print(f"{tab:<16} {meta['properties']['title']:<36} タブ{len(names)}枚 {cells:>6}セル  {names}")


if __name__ == "__main__":
    main()
