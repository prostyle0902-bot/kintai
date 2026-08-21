#!/bin/sh
# 損益計算書をDriveへ全部反映する
#
#   GOOGLE_SA_KEY_FILE=/tmp/sa.json sh push_all.sh
#
# 鍵の取り方は push.py の冒頭を参照（Dropboxから curl でディスクに落とす）。
#
# 反映先は3つ:
#   21期（スプレッドシート） … ふだん見るのはこちら。開くのが速い
#   21期（.xlsx）           … バックアップ。開くたびに変換が走るので重い
#   22期（スプレッドシート） … 2026年9月からの本番
set -e
python3 push_sheets.py 21期
echo
python3 push_sheets.py 22期
echo
python3 push.py 21期
