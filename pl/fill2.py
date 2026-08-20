#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""確定データを v2 ワークブックへ転記"""
import pandas as pd
from openpyxl.styles import Font, PatternFill, Border, Side
import build2, sales, inv6, inv7, payroll, cards, board

STAMP = "2026-08-20 05:30"

# 店舗ごとに行名が違うものの読み替え（既存スプシに合わせる）
REMAP = {("韓国酒場ハナ", "仕入（やまなか）"): "仕入（山中ストアー）"}
F_POST = PatternFill("solid", fgColor="D9F2D0")
THIN = Side(style="thin", color="B4C6E7")
BORD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# ---- 請求書 2607月（＝7月計上）: (タブ, 取引先, PL行, 税抜, 消費税, 元ファイル, 備考) ----
INV = [
 ("りゅうちゃん","片野商店","水道光熱費（ガス料金）",19583,1958,"りゅうちゃん/片野商店　りゅうちゃん.pdf",
  "PDF2ページ＝請求書2枚を合算（⑩17,691＋⑨1,892）／既存スプシ19,583と一致"),
 ("韓国酒場ハナ","片野商店","水道光熱費（ガス料金）",5518,551,"ハナ/片野商店　ハナ.pdf",""),
 ("もも焼きJAPAN","片野商店","水道光熱費（ガス料金）",7738,773,"もも焼き/片野商店　もも焼き.pdf",""),
 ("タコとハイボール","片野商店","水道光熱費（ガス料金）",8145,814,"タコハイ/片野商店　タコハイ.pdf",
  "既存スプシ8,145と一致"),
 ("神栖横丁","片野商店","水道光熱費（ガス料金）",19245,1924,"横丁/片野商店　横丁.pdf",""),
 ("神栖横丁","水道料金","水道光熱費（水道料金）",76810,7681,"横丁/水道料金　横丁振替.pdf",
  "水道50,655＋下水道33,836（7月分）"),
 ("神栖横丁","NTTファイナンス","通信費",7534,753,"横丁/NTTファイアンス　横丁振替.pdf",
  "ファイル名「ファイアンス」は誤記。マスタのNTTファイナンスと同一（確認済）"),
 ("神栖横丁","USEN","通信費（USEN）",1280,128,"横丁/USEN 横丁振替.pdf","請求書No F-1-20260804-104412-01"),
 ("神栖横丁","USEN","通信費（USEN）",1280,128,"横丁/USEN　横丁振替.pdf","請求書No F-1-20260804-105743-01（別契約）"),
 ("神栖横丁","アルソック","ALSOK",20500,2050,"横丁/アルソック　横丁振替.pdf","既存スプシ20,500と一致"),
 ("神栖横丁","ウゴーク","ウゴーク",30000,3000,"横丁/ウゴーク　横丁8月末.pdf","既存スプシ30,000と一致"),
 ("神栖横丁","業務（グリスト清掃）","グリスト清掃業務課",40000,4000,"横丁/業務　横丁.pdf",
  "既存スプシ グリスト清掃業務課40,000と一致"),
 ("焼きたて屋","門倉石油","水道光熱費（ガス料金）",21122,2112,"焼きたて屋/門倉石油　焼きたて屋.pdf",
  "既存スプシ21,122と一致"),
 ("焼きたて屋","丸善エコアース","ごみ処分費",3600,360,"焼きたて屋/丸善エコアース　焼きたて屋.pdf",
  "既存スプシ ごみ処分費3,600と一致"),
 ("さわら十三里屋","椎名環境整備","廃棄物処分",12000,1200,"十三里屋/椎名環境整備　十三里屋8月末.pdf",
  "既存スプシ12,000と一致"),
 # なめがたしろはとふぁーむ（1枚を既存スプシと同じ行に分解）
 ("さわら十三里屋","なめがたしろはとふぁーむ","仕入（なめファ8％）",316948,25355,
  "十三里屋/なめがたしろはとふぁーむ　十三里屋8月末.pdf","仕入食材7月（8%軽減）／既存スプシ316,948と一致"),
 ("さわら十三里屋","なめがたしろはとふぁーむ","仕入（なめファ10％）",4600,460,
  "十三里屋/なめがたしろはとふぁーむ　十三里屋8月末.pdf","仕入れ資材7月／既存スプシ4,600と一致"),
 ("さわら十三里屋","なめがたしろはとふぁーむ","地代家賃（賃料）",200000,20000,
  "十三里屋/なめがたしろはとふぁーむ　十三里屋8月末.pdf","家賃"),
 ("さわら十三里屋","なめがたしろはとふぁーむ","水道光熱費（水道料金）",1100,110,
  "十三里屋/なめがたしろはとふぁーむ　十三里屋8月末.pdf","水道代7月分"),
 ("さわら十三里屋","なめがたしろはとふぁーむ","水道光熱費（電気料金）",54436,5443,
  "十三里屋/なめがたしろはとふぁーむ　十三里屋8月末.pdf","佐原従量電灯C 22,851＋佐原低圧電力 31,585"),
 ("さわら十三里屋","なめがたしろはとふぁーむ","その他経費（ロイヤリティ）",98943,9894,
  "十三里屋/なめがたしろはとふぁーむ　十三里屋8月末.pdf",
  "ブランド使用3.0% 31,245＋本部管理経費6.5% 67,698／既存スプシ98,943と一致"),
 # ヴィーナスダイニング（タコハイで確定）
 ("タコとハイボール","ヴィーナスダイニング","仕入（インフォマート8％）",126455,10116,
  "タコハイ/ヴィーナスダイニング　タコハイ8月末.pdf","材料代8%軽減／既存スプシ126,455と一致"),
 ("タコとハイボール","ヴィーナスダイニング","仕入（インフォマート10％）",21400,2140,
  "タコハイ/ヴィーナスダイニング　タコハイ8月末.pdf","材料代10%／既存スプシ21,400と一致"),
 ("タコとハイボール","ヴィーナスダイニング","仕入（インフォマート利用料）",2500,250,
  "タコハイ/ヴィーナスダイニング　タコハイ8月末.pdf","既存スプシ2,500と一致"),
 ("タコとハイボール","ヴィーナスダイニング","ロイヤリティ",45311,4531,
  "タコハイ/ヴィーナスダイニング　タコハイ8月末.pdf","ロイヤリティ5%／既存スプシ45,311と一致"),
 # 日本食研（1枚を3店舗に分割）
 ("韓国酒場ハナ","日本食研","仕入（日本食研）",31515,2521,
  "ハナ/日本食研　ハナ・もも焼き・タコハイ8月末.pdf","内訳(8795207)韓国酒場ハナ／全品8%軽減"),
 ("もも焼きJAPAN","日本食研","仕入（日本食研）",6087,487,
  "ハナ/日本食研　ハナ・もも焼き・タコハイ8月末.pdf","内訳(8489343)もも焼JAPAN／全品8%軽減"),
 ("タコとハイボール","日本食研","仕入（日本食研）",5056,404,
  "ハナ/日本食研　ハナ・もも焼き・タコハイ8月末.pdf",
  "内訳(9278338)神栖横丁5丁目＝タコハイ店／既存スプシ5,056と一致"),
 # 鳥害対策課
 ("鳥害対策課","日本鳩対策センター","仕入（日本鳩対策センター）",1003398,100340,
  "鳥害対策/日本鳩対策センター.pdf","税抜御買上額／既存スプシ1,003,398と一致"),
 ("鳥害対策課","西尾レントオール","外注費（西尾レントオール）",247350,24735,
  "鳥害対策/西尾レントオール.pdf","10%対象計"),
 # 業務課
 ("業務課","ビーエム","仕入（ビーエム）",129580,12958,"業務/ビーエム　.pdf",
  "当月御買上額（繰越330,922は含めない）／既存スプシ129,580と一致"),
 ("業務課","ニッセーデリカ","仕入（ニッセーデリカ）",2080,208,"業務/ニッセーデリカ.pdf",
  "既存スプシは税込2,288。新シートは税抜2,080で統一"),
 ("業務課","江東微生物研究所","仕入（江東微生物研究所）",600,60,"業務/江東微生物研究所.pdf",
  "既存スプシは税込660。新シートは税抜600で統一"),
 ("業務課","ケミカルテクノロジー","仕入（ケミカルテクノロジー）",72900,7290,"業務/ケミカルテクノロジー.pdf",
  "既存スプシ72,900と一致"),
 ("業務課","エス・アイ・ビー・エス","外注費（SIBS）",66000,6600,"業務/エスアイビーエス.pdf",
  "既存スプシ66,000と一致"),
 ("業務課","SUN-X","外注費（SUN-X）",36000,3600,"業務/SUN-X.pdf","既存スプシ36,000と一致"),
 ("業務課","ピカピカ","外注費（ピカピカ）",45455,4545,"業務/ピカピカ　7月末振替.pdf",
  "税込50,000・うち消費税4,545／既存スプシ45,455と一致"),
 ("業務課","ミノアカ","外注費（ミノアカ）",275000,27500,"業務/ミノアカ.xlsx",
  "7月分12件（信金各支店・ハウスクリーニング等）／既存スプシ275,000と一致"),
 ("業務課","日本ビルメン経営品質協議会","諸会費（日本ビルメン）",20000,0,
  "業務/日本ビルメン経営品質協議会.pdf",
  "令和8年8月分会費。消費税の記載なし・適格請求書ではない旨の注記あり（新規項目）"),
 ("業務課","リペアボックス","修繕費（リペアボックス）",15000,1500,"業務/リペアボックス.pdf",
  "東京食堂（神栖市）修理／新規項目"),
 ("本部","帝国データバンク","帝国データバンク",3000,300,"業務/帝国データバンク.pdf",
  "COSMOSNET 2026.07利用分／新規項目"),
]

INV += inv7.INV7_ADD   # 2026-08-20 追加投入分（本部フォルダ8件を含む28ファイル）

INV_HOLD = []   # 2026-08-18 利用者指示により全件解決（行名＝取引先名）

# 金額のない実績データ（明細ログにのみ記録）
JISSEKI = [("焼きたて屋","丸善エコアース","焼きたて屋/丸善エコアース実績.pdf",
            "2026年7月 回収実績: 月間回収量80kg／回収袋数7袋（可燃・生ごみ）。"
            "金額は『ごみ処分費』3,600円として別途計上済み")]


def main(dst="損益計算書_21期テスト版.xlsx"):
    wb = build2.new_wb()
    ok = pd.read_csv("out_meisai.csv", encoding="utf-8-sig")
    try:
        hold = pd.read_csv("out_hold.csv", encoding="utf-8-sig")
    except pd.errors.EmptyDataError:          # 保留ゼロ件のとき
        hold = pd.DataFrame(columns=["利用日", "店舗", "取引先", "税込", "理由"])
    live = ok[ok["判定"] != "除外"].copy()
    live["税抜"] = live["税抜"].astype(int)

    posted, missing = 0, []
    for (tab, plrow, month), val in live.groupby(["店舗", "PL行", "計上月"])["税抜"].sum().items():
        plrow = REMAP.get((tab, plrow), plrow)
        if plrow not in build2.RIDX[tab]:
            missing.append((tab, plrow, "カード")); continue
        c = wb[tab][f"{build2.MCOL[month]}{build2.RIDX[tab][plrow]}"]
        c.value = int(c.value or 0) + int(val); c.fill = F_POST; c.number_format = "#,##0"
        posted += 1
    for tab, vendor, plrow, ex, tax, src, biko in INV:
        if plrow not in build2.RIDX[tab]:
            missing.append((tab, plrow, "請求書")); continue
        c = wb[tab][f"{build2.MCOL['7月']}{build2.RIDX[tab][plrow]}"]
        c.value = int(c.value or 0) + ex; c.fill = F_POST; c.number_format = "#,##0"
        posted += 1
    # 2606月 → 6月列（仕様どおり）
    for tab, vendor, plrow, ex, tax, src, biko in inv6.INV6:
        if plrow not in build2.RIDX[tab]:
            missing.append((tab, plrow, "請求書6月")); continue
        c = wb[tab][f"{build2.MCOL['6月']}{build2.RIDX[tab][plrow]}"]
        c.value = int(c.value or 0) + ex; c.fill = F_POST; c.number_format = "#,##0"
        posted += 1

    # ===== JCB / 三井住友カード（支払日ベース） =====
    F_CARD = PatternFill("solid", fgColor="FFF2CC")
    card_cells, card_hold = 0, []
    agg = {}
    for tab, merch, plrow, ex, tax, src, m, used, iss in cards.rows():
        if plrow is None:
            card_hold.append((iss, merch, ex, used)); continue
        agg[(tab, plrow, m)] = agg.get((tab, plrow, m), 0) + ex
    for (tab, plrow, m), val in agg.items():
        if plrow not in build2.RIDX[tab]:
            missing.append((tab, plrow, "カード(JCB/三井住友)")); continue
        c = wb[tab][f"{build2.MCOL[m]}{build2.RIDX[tab][plrow]}"]
        c.value = int(c.value or 0) + val; c.fill = F_CARD; c.number_format = "#,##0"
        card_cells += 1
    print("JCB/三井住友セル", card_cells, "／保留", len(card_hold))

    # ===== board（売掛）→ 業務課・鳥害対策課・神栖横丁の売上 =====
    F_BOARD = PatternFill("solid", fgColor="FCE4EC")
    board_rows = list(board.rows())
    for tab, plrow, m, ex, tax, n, src, detail in board_rows:
        c = wb[tab][f"{build2.MCOL[m]}{build2.RIDX[tab][plrow]}"]
        c.value = int(c.value or 0) + ex; c.fill = F_BOARD; c.number_format = "#,##0"
    print("boardセル", len(board_rows))
    # boardが正になる (タブ, 行, 月) — 既存スプシの売上転記から除外する
    board_owned = set(board.SUPPRESS) | {(t, r, m) for t, r, m, *_ in board_rows}

    # ===== 人件費・法定福利費（給与データ 4月〜7月） =====
    F_PAY = PatternFill("solid", fgColor="E2D9F2")
    pay_cells = 0
    for tab, plrow, m, val, note in payroll.rows():
        if plrow not in build2.RIDX[tab]:
            missing.append((tab, plrow, "給与")); continue
        c = wb[tab][f"{build2.MCOL[m]}{build2.RIDX[tab][plrow]}"]
        c.value = int(c.value or 0) + val; c.fill = F_PAY; c.number_format = "#,##0"
        pay_cells += 1
    print("人件費セル", pay_cells)

    # ===== 売上（既存スプシから転記） =====
    F_SALES = PatternFill("solid", fgColor="FDE9D9")
    sales_cells = 0
    for tab, rows in sales.SALES.items():
        for rowname, vals in rows.items():
            r = build2.RIDX[tab][rowname]
            for i, m in enumerate(build2.MONTHS):
                if vals[i] is None:
                    continue
                if (tab, rowname, m) in board_owned:
                    continue          # boardを正とするので既存スプシ値は使わない
                c = wb[tab][f"{build2.MCOL[m]}{r}"]
                c.value = int(vals[i]); c.fill = F_SALES; c.number_format = "#,##0"
                sales_cells += 1
    print("売上セル", sales_cells)

    # 明細ログ
    ws = wb.create_sheet("明細ログ")
    ws.cell(1, 1, "明細ログ（転記の元データ／検証用）").font = Font(bold=True, size=13)
    hdr = ["日付","店舗","取引先","摘要","税込","税率","税抜","消費税","転記先PL行","転記先の月",
           "データ元ファイル名","処理日時","備考"]
    for j, h in enumerate(hdr, start=1):
        c = ws.cell(2, j, h); c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="305496"); c.border = BORD
    for _, x in live.iterrows():
        ws.append([x["利用日"], x["店舗"], x["取引先"], x["摘要"], int(x["税込"]), int(x["税率"]),
                   int(x["税抜"]), int(x["消費税"]),
                   REMAP.get((x["店舗"], x["PL行"]), x["PL行"]), x["計上月"], x["元ファイル"], STAMP,
                   "" if pd.isna(x.get("備考")) else x.get("備考")])
    for _, x in ok[ok["判定"] == "除外"].iterrows():
        ws.append([x["利用日"], x["店舗"], x["取引先"], x["摘要"], int(x["税込"]), "", "", "",
                   "（除外）二重計上防止", "", x["元ファイル"], STAMP, x["備考"]])
    for tab, vendor, plrow, ex, tax, src, biko in INV:
        ws.append(["2026-07", tab, vendor, "請求書", ex + tax, "", ex, tax, plrow, "7月",
                   f"買掛/21期/2607月/{src}", STAMP, biko])
    for tab, vendor, plrow, ex, tax, src, biko in inv6.INV6:
        ws.append(["2026-06", tab, vendor, "請求書", ex + tax, "", ex, tax, plrow, "6月",
                   f"買掛/21期/2606月/確認済/{src}", STAMP, biko])
    for tab, vendor, src, biko in JISSEKI:
        ws.append(["2026-07", tab, vendor, "実績", "", "", "", "", "（金額なし・実績のみ）", "7月",
                   f"買掛/21期/2607月/{src}", STAMP, biko])
    for tab, merch, plrow, ex, tax, src, m, used, iss in cards.rows():
        ws.append([used, tab, merch, f"{iss}カード", ex + tax, 10, ex, tax,
                   plrow or "（保留）", m, src, STAMP,
                   "支払日ベースで計上（明細は利用日をそのまま記載）。"
                   "カード明細に税額の記載がないため10%で逆算"])
    for tab, plrow, m, ex, tax, n, src, detail in board_rows:
        for d, cust, v in detail:
            ws.append([d, tab, cust, "board請求", "", 10, v, "", plrow, m, src, STAMP,
                       f"boardの請求一覧CSVより（グループ列で部門判定）。この行を含む{n}件を合計{ex:,}円として転記"])
    for tab, plrow, m, val, note in payroll.rows():
        ws.append([f"2026-{ {'4月':'04','5月':'05','6月':'06','7月':'07'}[m] }", tab, "（給与）",
                   "人件費" if plrow.startswith("人件費") else "社会保険料",
                   val, "", val, 0, plrow, m, payroll.SRC, STAMP, note])
    for col, w in zip("ABCDEFGHIJKLM", [11,14,26,12,10,6,10,9,26,10,42,18,60]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A3"

    # 保留リスト
    hs = wb.create_sheet("保留リスト")
    hs.cell(1, 1, "保留リスト（判定できなかったもの／推測で埋めていません）"
            ).font = Font(bold=True, size=13, color="C00000")
    for j, h in enumerate(["区分","日付/月","店舗","取引先","金額(税込)","理由・メモ"], start=1):
        c = hs.cell(2, j, h); c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="C00000"); c.border = BORD
    for _, x in hold.iterrows():
        hs.append(["カード明細", x["利用日"], x["店舗"], x["取引先"], int(x["税込"]), x["理由"]])
    for iss, merch, ex, used in card_hold:
        hs.append([f"{iss}カード", used, "本部", merch, ex, "取引先マスタ（cards.py）に未登録。行名を要指示"])
    for tab, vendor, src, reason in INV_HOLD:
        hs.append(["請求書", "", tab, vendor, "", f"{src} — {reason}"])
    if hs.max_row == 2:
        hs.cell(3, 1, "保留はありません（2026-08-20 時点／カード明細6月・7月＋請求書2606月・2607月＋給与4〜7月）"
                ).font = Font(bold=True, color="008000")
    for col, w in zip("ABCDEF", [12,12,16,26,12,95]):
        hs.column_dimensions[col].width = w
    hs.freeze_panes = "A3"

    wb.save(dst)
    print("転記セル数", posted, "／ 明細ログ", ws.max_row - 2, "行 ／ 保留", hs.max_row - 2, "件")
    if missing:
        print("!! 行が見つからない:", missing)


if __name__ == "__main__":
    main()
