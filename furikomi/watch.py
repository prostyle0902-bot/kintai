#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Dropboxの見張りフォルダに変化があったかを、1フォルダ1問い合わせで見分ける

    python3 watch.py --check sizes.json   # 変わったフォルダだけ出す（変化あれば exit 1）
    python3 watch.py --save sizes.json    # いまの値を基準として記録する
    python3 watch.py --list               # 見張っているフォルダの一覧（コピペ用）

sizes.json は Claude が Dropbox コネクタで集めた
    {"ns:7987961616//買掛/21期/2608月": 19461960, ...}
という「フォルダのパス → folder_size」の辞書。

--- なぜ要るか -----------------------------------------------------------
定期チェックは1日3回動く。毎回 list_folder で全ファイルを並べていたが、
買掛だけで70件超・銀行明細は117件あり、**変化が無い日でも毎回それを全部
読む**ことになっていた（2026-09-17 時点で1回およそ3万文字）。

Dropbox の get_file_metadata はフォルダに対して `folder_size` を返す。
これは**中に入っているファイルのサイズの合計**で、2026-09-17 に実測して
確かめた（買掛/21期/2608月: 73ファイルのサイズ合計 19,461,960 =
folder_size 19,461,960 でぴったり一致）。

なので「folder_size が前回と同じ＝中身は変わっていない」と見なせる。
変わったフォルダだけ list_folder すればよい。

--- ★この見分け方の穴 ---------------------------------------------------
**サイズがちょうど同じファイルに差し替わった場合だけ見逃す。**
（1件消えて同じバイト数の1件が増える、など）

そのため **1日1回、朝の回だけは従来どおり全フォルダを list_folder する**。
昼と夕方の回はこの見分け方で済ませる。手順は furikomi/README.md を参照。

--- 使い方（定期チェックの中で）------------------------------------------
1. Claude が見張りフォルダそれぞれに get_file_metadata を呼ぶ（--list の順）
2. 返ってきた folder_size を sizes.json に書く
   ★folder_size のキーが無いフォルダは【空】。0 と書く
3. python3 watch.py --check sizes.json
4. 「変わった」と出たフォルダだけ list_folder して、台帳と突き合わせる
5. 処理が終わったら python3 watch.py --save sizes.json
"""
import json
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
STORE = os.path.join(BASE, "watch_sizes.json")

NS = "ns:7987961616"

# 見張るフォルダ。増やすときはここに足す
WATCH = [
    # (A) 振込一覧 … 請求書
    f"{NS}//買掛/21期/2607月",
    f"{NS}//買掛/21期/2608月",
    f"{NS}//買掛/22期",
    # (B) 損益計算書 … 元データ
    f"{NS}//銀行明細/21期",
    f"{NS}//銀行明細/22期",
    f"{NS}//freeeカード明細/21期",
    f"{NS}//freeeカード明細/22期",
    f"{NS}//会計明細/21期",
    f"{NS}//会計明細/22期",
    f"{NS}//かめや（焼きたて屋本部）/21期",
    f"{NS}//かめや（焼きたて屋本部）/22期",
]


def load():
    if not os.path.exists(STORE):
        return {}
    with open(STORE, encoding="utf-8") as f:
        return json.load(f)


def check(path):
    with open(path, encoding="utf-8") as f:
        now = json.load(f)
    old = load()
    missing = [w for w in WATCH if w not in now]
    if missing:
        print("★取れていないフォルダがある。get_file_metadata を呼び直すこと:")
        for w in missing:
            print("   ", w)
        return 1
    if not old:
        print("★基準値がまだ無い（初回）。全フォルダを list_folder してから --save すること")
        return 1

    changed = [w for w in WATCH if now.get(w) != old.get(w)]
    for w in WATCH:
        a, b = old.get(w), now.get(w)
        mark = "★変わった" if a != b else "   同じ　"
        print(f"  {mark} {w.split('//')[1]:<28} {a if a is not None else '―':>12} → {b:>12}")
    if changed:
        print(f"\n変わったフォルダ {len(changed)}件。ここだけ list_folder すること:")
        for w in changed:
            print("   ", w)
        return 1
    print(f"\nどのフォルダも変わっていない（{len(WATCH)}フォルダ）。新着なしとして終了してよい")
    return 0


def save(path):
    with open(path, encoding="utf-8") as f:
        now = json.load(f)
    keep = {w: now[w] for w in WATCH if w in now}
    with open(STORE, "w", encoding="utf-8") as f:
        json.dump(keep, f, ensure_ascii=False, indent=1, sort_keys=True)
        f.write("\n")
    print(f"{len(keep)}フォルダの値を {os.path.basename(STORE)} に記録した")
    return 0


def main():
    a = sys.argv[1:]
    if a[:1] == ["--list"]:
        for w in WATCH:
            print(w)
        return 0
    if len(a) == 2 and a[0] == "--check":
        return check(a[1])
    if len(a) == 2 and a[0] == "--save":
        return save(a[1])
    sys.exit(__doc__)


if __name__ == "__main__":
    sys.exit(main())
