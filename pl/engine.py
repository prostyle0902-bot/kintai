#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""freeeカード明細 → PL転記エンジン（仕様【4】）"""
import pandas as pd, unicodedata, math, re
from decimal import Decimal, ROUND_FLOOR

# ---- 取引先マスタ（カード明細用）: キーワード, 税率, 分類, PL行 ----
MASTER = [
 ("タイヨー",8,"スーパー","仕入（freeeカード）"),
 ("カスミ",8,"スーパー","仕入（freeeカード）"),
 ("ｷﾞﾖｳﾑｽ-ﾊﾟ-",8,"スーパー","仕入（freeeカード）"),
 ("業務スーパー",8,"スーパー","仕入（freeeカード）"),
 ("セイミヤ",8,"スーパー","仕入（freeeカード）"),
 ("ビッグハウス",8,"スーパー","仕入（freeeカード）"),
 ("さえき",8,"スーパー","仕入（freeeカード）"),
 ("道の駅",8,"直売所","仕入（freeeカード）"),
 ("MEGAドン・キホーテ",8,"ディスカウント","仕入（freeeカード）"),
 ("ローソン",8,"コンビニ","仕入（freeeカード）"),
 ("丸や",8,"食品","仕入（freeeカード）"),
 ("ﾀﾍﾞﾁﾖｸ",8,"ネット食材","仕入（ネット食材）"),
 ("イーアスつくば",8,"食品","仕入（freeeカード）"),
 ("ウエルシア",8,"ドラッグ","消耗品費（freeeカード）"),
 ("マツモトキヨシ",8,"ドラッグ","消耗品費（freeeカード）"),
 ("クスリのアオキ",8,"ドラッグ","消耗品費（freeeカード）"),
 ("酒蔵やまなか",10,"酒","仕入（やまなか）"),
 ("ベストリカー",10,"酒","仕入（freeeカード）"),
 ("リカー＆フーズひしや",10,"酒","仕入（freeeカード）"),
 ("ミトエクセル",10,"酒","仕入（freeeカード）"),
 ("コダマインリョウ",10,"飲料","仕入（freeeカード）"),
 ("ヒーロー佐原店",10,"酒","仕入（freeeカード）"),
 ("AMAZON",10,"通販","仕入（freeeカード）"),
 ("(ﾗｸﾃﾝ",10,"楽天","仕入（freeeカード）"),
 ("MonotaRO",10,"資材","消耗品費（freeeカード）"),
 ("シモジマ",10,"包材","消耗品費（freeeカード）"),
 ("ミクリード",10,"業務用","仕入（freeeカード）"),
 ("ラクスル",10,"印刷","広告宣伝費（共通宣伝費）"),
 ("プリントパック",10,"印刷","広告宣伝費（共通宣伝費）"),
 ("カインズ",10,"ホームセンター","消耗品費（freeeカード）"),
 ("DCM",10,"ホームセンター","消耗品費（freeeカード）"),
 ("コメリ",10,"ホームセンター","消耗品費（freeeカード）"),
 ("KOMERI",10,"ホームセンター","消耗品費（freeeカード）"),
 ("セリア",10,"100均","消耗品費（freeeカード）"),
 ("ダイソー",10,"100均","消耗品費（freeeカード）"),
 ("ワークマン",10,"作業着","消耗品費（freeeカード）"),
 ("オフハウス",10,"備品","消耗品費（freeeカード）"),
 ("ドコモ",10,"通信","通信費"),
 ("USEN",10,"BGM","その他経費"),
 ("ダスキン",10,"清掃","その他経費"),
 ("LINE公式アカウント",10,"広告","広告宣伝費（共通宣伝費）"),
 ("ダゾーン",10,"放映","その他経費"),
 ("OPENAI",10,"ソフト","その他経費"),
 ("ﾗｸﾃﾝﾄﾗﾍﾞﾙ",10,"宿泊","旅費・交通費"),
 ("すき家",10,"外食","接待交際費"),
 ("大阪王将",10,"外食","接待交際費"),
 ("ｱﾎﾟﾛｽﾃ-ｼﾖﾝ",10,"燃料","旅費・交通費"),
 ("ｴﾈｵｽ",10,"燃料","旅費・交通費"),
 ("コスモ石油",10,"燃料","旅費・交通費"),
 ("カ)ﾐﾔｺ",10,"資材","仕入（その他）"),
 ("戸田産業",10,"資材","仕入（その他）"),
 ("ミズノ",10,"資材","仕入（その他）"),
 ("CHIMNEY TOWN",10,"物販","その他経費"),
 ("古河ＳＳ",10,"燃料","旅費・交通費"),   # 利用者確認済 2026-08-17
]

# ---- カード名(下4桁) → 店舗タブ ----
CARD2TAB = {
 "4188":"タコとハイボール", "0538":"さわら十三里屋", "9782":"さわら十三里屋",
 "7655":"韓国酒場ハナ", "0207":"もも焼きJAPAN", "2739":"りゅうちゃん",
 "0585":"焼きたて屋", "8772":"神栖横丁", "3870":"神栖横丁",
 "0321":"業務課", "0819":"鳥害対策課",
}
# 除外: 十三里屋 固定費カード0538 の ドコモ/USEN/ダスキン（請求書側で計上）
EXCLUDE_0538 = ("ドコモ","USEN","ダスキン")
# メモ判定
MEMO10 = ("酒","資材","炭")
MEMO8  = ("食材",)

# 税率が矛盾したときに勝つキーワード（利用者確認済 2026-08-17）
# 例: 「セリア　カスミ神栖店」→セリア10% / 「タイヨーベストリカー神栖店」→ベストリカー10%
OVERRIDE_WINNERS = ("セリア", "ベストリカー")


def norm(s):
    """全角/半角・大小文字を揃えて比較用に正規化"""
    if s is None or (isinstance(s, float) and math.isnan(s)):
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    return s.replace("　", " ").replace(" ", "").upper()


def merchant_only(riyouten, cardname):
    """利用店名は「<店名> <カード名>」形式。末尾のカード名を除いて店名部分だけ返す。
    カード名にマスタ語(コメリ/横丁など)が含まれ誤ヒットするのを防ぐ。"""
    s = str(riyouten)
    c = str(cardname)
    if c and s.endswith(c):
        return s[: -len(c)].strip()
    # 正規化しても末尾一致するか確認（全角スペース等の揺れ対策）
    ns, nc = norm(s), norm(c)
    if nc and ns.endswith(nc):
        # 元文字列側で対応する長さを削る
        cut = len(s)
        while cut > 0 and norm(s[:cut]).endswith(nc):
            cut -= 1
        return s[: cut + 1].strip()
    return s.strip()


def match_master(name):
    """マスタ一致（複数ヒットは全部返す）"""
    n = norm(name)
    hits = [m for m in MASTER if norm(m[0]) and norm(m[0]) in n]
    if not hits:
        return []
    # 最長キーワード優先で並べる
    hits.sort(key=lambda m: len(norm(m[0])), reverse=True)
    return hits


def memo_rate(memo):
    n = norm(memo)
    if not n:
        return None, None
    if any(norm(k) in n for k in MEMO10):
        return 10, "メモ"
    if any(norm(k) in n for k in MEMO8):
        return 8, "メモ"
    return None, None


def load(path, riyou_month):
    d = pd.read_csv(path, encoding="utf-8-sig", dtype=str)
    d["利用額"] = d["利用額"].astype(int)
    d["利用日"] = pd.to_datetime(d["利用日"])
    d["_riyou_month"] = riyou_month
    return d


def classify(d):
    rows, hold = [], []
    for _, r in d.iterrows():
        card4 = str(r["カード下4桁"]).zfill(4)
        tab = CARD2TAB.get(card4)
        name_raw = r["利用店名"]
        name = merchant_only(name_raw, r["カード名"])
        memo = r.get("領収書メモ")
        amt = int(r["利用額"])
        day = r["利用日"]
        base = dict(利用日=day.strftime("%Y-%m-%d"), 店舗=tab, 取引先=name,
                    摘要=("" if pd.isna(memo) else memo), 税込=amt,
                    カード=f'{r["カード名"]}({card4})', 元ファイル=r["_srcfile"])

        if tab is None:
            hold.append({**base, 理由: "カード下4桁が未定義"} if False else
                        {**base, "理由": f"カード下4桁 {card4} がタブ対応表にない"})
            continue

        # 除外ルール
        if card4 == "0538" and any(norm(k) in norm(name) for k in EXCLUDE_0538):
            base.update(判定="除外", 税率="", 税抜="", 消費税="", PL行="（除外）二重計上防止",
                        計上月="", 備考="十三里屋固定費カード0538: 請求書側で計上")
            rows.append(base); continue

        hits = match_master(name)
        rate, src = memo_rate(memo)

        if rate is None:                      # 優先1で決まらない → 優先2
            if not hits:
                hold.append({**base, "理由": "マスタにキーワード一致なし・メモもなし"})
                continue
            distinct = {h[1] for h in hits}
            if len(distinct) > 1:
                win = [h for h in hits if h[0] in OVERRIDE_WINNERS]
                if len(win) == 1:
                    hits = win + [h for h in hits if h not in win]
                    src_note = "マスタ(優先ルール)"
                else:
                    hold.append({**base, "理由": "マスタが複数ヒットし税率が矛盾: " +
                                 " / ".join(f"{h[0]}={h[1]}%" for h in hits)})
                    continue
            else:
                src_note = "マスタ"
            rate, src = hits[0][1], src_note

        if not hits:
            hold.append({**base, "理由": f"メモから税率{rate}%は判るが、PL行を決めるマスタ一致がない"})
            continue

        kw, mrate, cat, plrow = hits[0]

        # 酒蔵やまなか 特別ルール: 利用日1〜5日は前月分へ
        month = day.month
        biko = ""
        if "やまなか" in norm(kw) or "ヤマナカ" in norm(name):
            if 1 <= day.day <= 5:
                month = (day - pd.offsets.MonthBegin(1)).month if day.day else month
                month = (day.replace(day=1) - pd.Timedelta(days=1)).month
                biko = "前月分に振替"

        # 円未満切り捨て。floatだと 30800/1.1=27999.99… で1円ずれるため Decimal を使う
        tax_ex = int((Decimal(amt) / (1 + Decimal(rate) / 100)).to_integral_value(ROUND_FLOOR))
        # ★1件5,000円以下（税抜）の接待交際費は会議費へ（利用者指示 2026-08-22）。
        #   cards.py（JCB/三井住友）と同じ扱い。判定はこの損益計算書に合わせて税抜。
        if plrow == "接待交際費" and tax_ex <= 5000:
            plrow = "会議費"
            biko = (biko + "／" if biko else "") + "1件5,000円以下なので会議費"
        base.update(判定=src, 税率=rate, 税抜=tax_ex, 消費税=amt - tax_ex,
                    PL行=plrow, 計上月=f"{month}月", 備考=biko, 一致キーワード=kw, 分類=cat)
        rows.append(base)
    return pd.DataFrame(rows), pd.DataFrame(hold)


if __name__ == "__main__":
    frames = []
    for path, m, src in [("csv/statement-2026-07.csv", 6, "statement-2026-07.csv"),
                         ("csv/statement-2026-08.csv", 7, "statement-2026-08.csv")]:
        d = load(path, m); d["_srcfile"] = src; frames.append(d)
    d = pd.concat(frames, ignore_index=True)
    ok, hold = classify(d)
    ok.to_csv("out_meisai.csv", index=False, encoding="utf-8-sig")
    hold.to_csv("out_hold.csv", index=False, encoding="utf-8-sig")
    print("判定OK", len(ok), "／ 保留", len(hold))
