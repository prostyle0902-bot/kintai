#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""給料一覧表（個人別）PDF → {社員番号: (総支給額, 社会保険計)}

出典: Dropbox /※プロスタイル給与※/プロスタイル給与R7年・R8年/給料一覧表-YYYYMM.pdf
      ローカル kyuyo/YYYYMM.pdf

PDFの形
    1ページに12人ぶん、社員番号の列が横に並ぶ（5ページ）。最終ページの右端に総合計。
    社員番号は「部門4桁-連番4桁」。先頭4桁が部門コード。
    拾うのは2行だけ:
        総支給額     → PL「人件費（店長）」「人件費（アルバイト）」
        社会保険計   → PL「法定福利費」
    ★どちらも会社負担ぶんではなく本人ぶん。既存PLもこの値で作られている
      （payroll.py の4〜7月と1円まで一致することを確認済み）。

読み方
    pdftotext -bbox-layout でXMLにして、単語ごとの座標で拾う。
    ★-layout の固定幅テキストだと桁位置がページごとに最大16桁ずれて対応づけできない。
      社員番号の右端と値の右端が合うはず……という当たりを付けたが実測で合わなかった。
      座標なら「社員番号の右端x」と「値の右端x」が1pt以内で一致するので確実。
    総合計の列と突き合わせて検算する。合わなければ止める。
"""
import os
import re
import subprocess

BASE = os.path.dirname(os.path.abspath(__file__))
DIR = os.path.join(BASE, "kyuyo")
_EMP = re.compile(r"\d{4}-\d{4}")
_NUM = re.compile(r"[\d,]+")
# 拾う行のラベル（空白を除いた形）
ROWS = {"総支給額": "総支給額", "社会保険計": "社会保険計"}


def _words(path, page):
    """[(xMin, xMax, yMin, text)] を返す。"""
    import xml.etree.ElementTree as ET
    xml = subprocess.run(["pdftotext", "-bbox-layout", "-f", str(page),
                          "-l", str(page), path, "-"],
                         capture_output=True).stdout.decode("utf-8", "replace")
    root = ET.fromstring(xml)
    ns = "{http://www.w3.org/1999/xhtml}"
    out = []
    for w in root.iter(f"{ns}word"):
        out.append((float(w.get("xMin")), float(w.get("xMax")),
                    float(w.get("yMin")), (w.text or "").strip()))
    return out


def _pages(path):
    n = int(re.search(r"Pages:\s+(\d+)",
                      subprocess.run(["pdfinfo", path], capture_output=True)
                      .stdout.decode()).group(1))
    return n


def _emp_cols(words):
    """社員番号の列 [(右端x, 社員番号)]。
    ★単語が「0002-」「0011」のように割れていることがあるので、
      見出し行のyで文字を連結し直してから正規表現をかける。"""
    import collections
    by_y = collections.defaultdict(list)
    for x0, x1, y, t in words:
        by_y[round(y, 1)].append((x0, x1, t))
    # ★いちばん多く社員番号が並ぶ行を見出しとする。
    #   最終ページは社員が1人だけのことがある（202601がそう）ので「2人以上」で
    #   探すと見出しを見つけられず、総合計と社員の値を混ぜてしまう。
    best, n_best = None, 0
    for y, ws in by_y.items():
        ws.sort()
        joined = "".join(t for _a, _b, t in ws)
        n = len(_EMP.findall(joined))
        if n > n_best:
            best, n_best = ws, n
    if best is None or n_best == 0:
        return []
    # 連結文字列の各文字が、どの単語の右端に対応するかを持っておく
    joined, ends = "", []
    for _x0, x1, t in best:
        joined += t
        ends += [x1] * len(t)
    return [(ends[m.end() - 1], m.group()) for m in _EMP.finditer(joined)]


def _read_total(words, total):
    """社員のいないページから総合計だけ拾う。"""
    for key, idx in (("総支給額", 0), ("社会保険計", 1)):
        y = _row_y(words, key)
        if y is None:
            continue
        near = sorted((x1, t) for x0, x1, yy, t in words
                      if abs(yy - y) < 2.5 and _NUM.fullmatch(t))
        joined = "".join(t for _x, t in near).replace(",", "")
        if joined.isdigit():
            total[idx] = int(joined)


def _row_y(words, label):
    """ラベルの文字が並ぶ行のy座標。1文字ずつバラけているので連結して探す。"""
    by_y = {}
    for x0, x1, y, t in words:
        by_y.setdefault(round(y, 1), []).append((x0, t))
    for y, ws in sorted(by_y.items()):
        joined = "".join(t for _x, t in sorted(ws))
        if joined.replace(" ", "").startswith(label):
            return y
    return None


def parse(ym):
    """{社員番号: (総支給額, 社会保険計)} と 総合計(総支給額, 社会保険計)。"""
    path = os.path.join(DIR, f"{ym}.pdf")
    if not os.path.exists(path):
        return None, None
    out, total = {}, [0, 0]
    for p in range(1, _pages(path) + 1):
        words = _words(path, p)
        emps = _emp_cols(words)
        # ★総合計が社員のいない最終ページに単独で出ることがある（202601がそう）。
        #   社員列が無いからと飛ばすと検算ができなくなる。
        if not emps:
            _read_total(words, total)
            continue
        # 総合計の列（見出しの「総合計」）。最終ページだけにある
        sogo = next((x1 for x0, x1, y, t in words if t == "総合計"), None)
        for key, idx in (("総支給額", 0), ("社会保険計", 1)):
            y = _row_y(words, key)
            if y is None:
                continue
            # ★ラベルと値でyが1.2ptほどずれる（実測）。行間は約8ptなので2.5で切る
            vals = [(x1, t.replace(",", "")) for x0, x1, yy, t in words
                    if abs(yy - y) < 2.5 and _NUM.fullmatch(t)]
            # 値は社員番号より右端が約4pt右に出る（実測）。列の間隔は約72ptなので
            # 「いちばん近い社員の列」に割り当て、20pt以上離れていたら捨てる。
            for x1, t in vals:
                if not t.isdigit():
                    continue
                right, emp = min(emps, key=lambda e: abs(x1 - 4 - e[0]))
                if abs(x1 - 4 - right) > 20:
                    continue                      # 総合計の列など
                a, c = out.get(emp, (0, 0))
                out[emp] = (a + int(t), c) if idx == 0 else (a, c + int(t))
            if sogo is not None and sogo - 4 - emps[-1][0] > 20:
                # 総合計は1文字ずつ分かれていることがあるので、右端が近いものを連結。
                # 最後の社員の列より 20pt 以上右にあるものだけを拾う。
                near = sorted((x1, t) for x0, x1, yy, t in words
                              if abs(yy - y) < 2.5 and x1 - 4 - emps[-1][0] > 20)
                joined = "".join(t for _x, t in near).replace(",", "")
                if joined.isdigit():
                    total[idx] = int(joined)
    out = {k: v for k, v in out.items() if v != (0, 0)}
    return out, tuple(total)


def by_bumon(ym):
    """{部門コード4桁: (総支給額計, 社会保険計)}"""
    emp, _ = parse(ym)
    if emp is None:
        return None
    out = {}
    for k, (pay, ins) in emp.items():
        b = k[:4]
        a, c = out.get(b, (0, 0))
        out[b] = (a + pay, c + ins)
    return out


if __name__ == "__main__":
    import sys
    for ym in (sys.argv[1:] or ["202607"]):
        emp, total = parse(ym)
        if emp is None:
            print(f"{ym}: PDFがありません")
            continue
        s = sum(v[0] for v in emp.values())
        t = sum(v[1] for v in emp.values())
        ok = "一致" if (s, t) == total else f"★不一致 総合計={total}"
        assert (s, t) == total, f"{ym}: 拾った合計 {(s, t)} と総合計 {total} が合わない"
        print(f"{ym}  {len(emp)}人  総支給額 {s:,} ／ 社会保険計 {t:,}   {ok}")
        for b, (a, c) in sorted(by_bumon(ym).items()):
            print(f"      部門{b}  総支給額 {a:>10,}  社会保険 {c:>9,}")
