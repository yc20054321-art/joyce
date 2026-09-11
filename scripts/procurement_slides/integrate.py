# -*- coding: utf-8 -*-
"""主副料採購整合說明 — 為什麼工時省得下來，人力卻減不掉。

Built from the two source decks, in the same visual language:
  * 主採：《主採業務範疇說明》2026-09-07 G（14 類、1,993 封、AI 可協助比例逐類 5-25%）
  * 副採：《V6副採工作範疇簡略說明》2026-09-07 Q1（11 類、ERP PO 2,656 張、每項 10-30 分）

Every figure here is lifted from those decks. The only derived numbers are two
weighted sums, both shown with their working on the slide:
  * 主採 AI 可協助比例的占比加權平均 = 12.6%（落在該簡報自述的「約 10-15%」內）
  * 副採 四個潛在介入點所在類別的占比合計 = 40%（20+5+10+5）
"""
import os, re, shutil, zipfile
from gen import (tb, shape, run, header, takeaway, source, wrap, emu, _id,
                 MARGIN_X, CONTENT_W, NAVY, BLUE, INK, BODY, MUTED,
                 PANEL, TINT, RED, REDBG, WHITE, FONT)

BASE, UN, OUT = 'q1.pptx', 'int_unpacked', 'integrated_out.pptx'


# ============================================================== 1. 封面
def cover():
    _id[0] = 1
    s = shape(0, 0, 13.333, 7.5, NAVY) + shape(0, 7.4, 13.333, 0.1, BLUE)
    s += tb(1.05, 2.30, 11.2, 1.40, [
        [run("主副料採購 — AI 導入評估", 4400, 1, WHITE)],
        [run("為什麼工時省得下來，人力卻減不掉？", 2700, 1, "C9DFF7")],
    ], lnspc=112)
    s += shape(1.05, 4.10, 3.0, 0.04, BLUE)
    s += tb(1.05, 4.38, 11.2, 1.75, [
        [run("一、主採／副採在做什麼 —— 量體與結構", 1500, 1, "D8E6F7")],
        [run("二、六個共通原因 —— 為什麼 AI 只能協助", 1500, 1, "D8E6F7")],
        [run("三、AI 已經接手到哪裡", 1500, 1, "D8E6F7")],
        [run("四、結論與人力建議", 1500, 1, "D8E6F7")],
    ], lnspc=145)
    s += tb(1.05, 6.30, 11.2, 0.50, [
        [run("彙整：主採／副採　│　2026-09-08　│　人力編制：主採 3 位＋副採 3 位", 1200, 1, WHITE)],
        [run("內部討論用，請勿外傳。部分內容由 AI 協助整理，數字需相關單位／本人複核後方可對外引用。",
             950, 0, "9EC4EF")],
    ], lnspc=150)
    return wrap(s)


# ============================================================== 2. 量體與結構
def scale():
    _id[0] = 1
    s = header("一、量體與結構", "主採與副採的共同形狀 — 類別多、單一類別占比低、每項只有 10–30 分鐘")

    lx, lw = MARGIN_X, 5.85
    rx, rw = MARGIN_X + 6.24, 5.85
    for x, w, name, col, stats, note in [
        (lx, lw, "主採（編制 3 位）", NAVY,
         [("14 類", "已盤點的工作類型"),
          ("1,993 封", "樣本期間信件總數"),
          ("15%", "最大單一類別：重大品質異常"),
          ("12.6%", "AI 可協助比例的占比加權平均")],
         "十四類的 AI 可協助比例逐類落在 5–25%，加權後 12.6% —— 與該簡報自述的「約 10–15%」一致。"),
        (rx, rw, "副採（編制 3 位）", BLUE,
         [("11 類", "已盤點的工作類型"),
          ("2,656 張", "ERP PO（2026 年 1–9 月）"),
          ("20%", "最大單一類別：by 款採購／品質 ISSUE"),
          ("10–30 分", "每項作業的典型耗時")],
         "Shipping Monitor key-in 約 4,000 次（推算）；十一類佔比加總 100%，分布平均、無單一大宗。"),
    ]:
        s += tb(x, 1.70, w, 0.28, [[run(name, 1250, 1, col)]])
        s += shape(x, 2.04, w, 3.36, PANEL, "roundRect")
        s += shape(x, 2.04, 0.06, 3.36, col)
        y = 2.20
        for num, lab in stats:
            s += tb(x + 0.26, y, 1.85, 0.42, [[run(num, 1700, 1, col)]], anchor="ctr")
            s += tb(x + 2.24, y, w - 2.50, 0.42, [[run(lab, 950, 0, BODY)]], anchor="ctr", lnspc=120)
            y += 0.50
        s += tb(x + 0.26, y + 0.14, w - 0.52, 0.62, [[run(note, 900, 0, MUTED)]], lnspc=128)

    s += shape(MARGIN_X, 5.58, CONTENT_W, 0.62, REDBG, "roundRect")
    s += shape(MARGIN_X, 5.58, 0.06, 0.62, RED)
    s += tb(MARGIN_X + 0.24, 5.66, 2.05, 0.24, [[run("★ 關鍵的形狀", 1050, 1, RED)]])
    s += tb(MARGIN_X + 2.40, 5.66, CONTENT_W - 2.64, 0.48,
            [[run("兩邊都沒有一塊「大到值得整批自動化」的工作 —— 主採最大的一類占 15%，副採最大的兩類各占 20%。"
                  "就算把最大的一類完全自動化，也只動到整體的一到兩成，而那一類偏偏是最不能自動化的品質異常。",
                  1000, 1, BODY)]], lnspc=125)

    s += takeaway(6.40, [
        run("→ ", 1300, 1, NAVY),
        run("工作是「碎的」而不是「大的」—— 自動化任何單一類別的邊際效益都有限，", 1300, 1, INK),
        run("這是人力減不掉的第一個結構原因", 1300, 1, RED),
        run("。", 1300, 1, INK),
    ], h=0.58)
    s += source("資料來源：主採業務範疇說明（2026-09-07）十四大類表；V6副採工作範疇簡略說明（2026-09-07）"
                "十一大類與 ERP／Shipping Monitor 實績。12.6% 為依占比加權平均之推算值。")
    return wrap(s)


# ============================================================== 3. 六個共通原因
def reasons():
    _id[0] = 1
    s = header("二、共通原因", "六個共通原因 — 為什麼 AI 只能協助，不能按比例替換人力")
    s += tb(MARGIN_X, 1.66, CONTENT_W, 0.26,
            [[run("主採談布廠、副採談副料供應商，對象不同，但卡住的地方一模一樣", 1250, 1, NAVY)]])

    cards = [
        ("① 是「議定」，不是「查詢」",
         "交期、價格、責任都要多輪來回談定 —— 主採還含人民幣匯率商談，副採要與供應商／成衣廠反覆議定。"
         "AI 產得出初稿，談不出共識。"),
        ("② 等待不會因為 AI 而變快",
         "工時花在等布廠、成衣廠、客戶回覆。處理速度從來不是瓶頸 —— 就算報表 3 秒生成，對方明天才回。"),
        ("③ 要看實物才能拍板",
         "色差、移染、白汙、看色核色，都要目視實物與色板比對，並送樣、複測、打樣驗證。"
         "這些動作本身無法由 AI 執行。"),
        ("④ 溝通載體不在系統裡",
         "缸差核色靠微信與上海布組即時互動；重大異常多為電話、現場會議、口頭確認。"
         "零散對話與圖片，AI 讀不到也追不了。"),
        ("⑤ 承諾與責任要人來扛",
         "下單即產生商業約束力；異常究責、換不換標、CAP 報告都要有人簽名負責。"
         "責任無法交給工具，這與工具好不好無關。"),
        ("⑥ 資料會回頭改寫上一輪的答案",
         "訂單切割、分廠重算、TA 變動、PH3 填補 —— 任何時間點算出的答案，都可能在下一刻被工廠或供應商的回覆推翻。"),
    ]
    cols, gap = 3, 0.24
    cw = (CONTENT_W - gap * (cols - 1)) / cols
    ch, vgap, y0 = 1.50, 0.16, 2.02
    for i, (t, d) in enumerate(cards):
        r, c = divmod(i, cols)
        x = MARGIN_X + c * (cw + gap)
        y = y0 + r * (ch + vgap)
        s += shape(x, y, cw, ch, PANEL, "roundRect")
        s += shape(x, y, 0.06, ch, BLUE)
        s += tb(x + 0.22, y + 0.13, cw - 0.44, 0.24, [[run(t, 1050, 1, BLUE)]])
        s += tb(x + 0.22, y + 0.44, cw - 0.44, ch - 0.58, [[run(d, 900, 0, BODY)]], lnspc=128)

    s += shape(MARGIN_X, 5.38, CONTENT_W, 0.62, REDBG, "roundRect")
    s += shape(MARGIN_X, 5.38, 0.06, 0.62, RED)
    s += tb(MARGIN_X + 0.24, 5.46, 2.05, 0.24, [[run("★ 六項的共同點", 1050, 1, RED)]])
    s += tb(MARGIN_X + 2.40, 5.46, CONTENT_W - 2.64, 0.48,
            [[run("沒有一項是「人做得比較慢」，全部都是「不能由工具代為決定」。"
                  "AI 加速的是前置整理，決定權從頭到尾沒有移轉過。", 1000, 1, BODY)]], lnspc=125)

    s += takeaway(6.24, [
        run("→ ", 1300, 1, NAVY),
        run("能被壓縮的是產出速度；主副採的工時花在等別人回覆、和別人談定、對實物拍板 —— ", 1300, 1, INK),
        run("這三件事 AI 幫得上忙，卻沒有一件能代替人完成", 1300, 1, RED),
        run("。", 1300, 1, INK),
    ], h=0.60)
    s += source("資料來源：主採業務範疇說明（2026-09-07）「為什麼」欄之歸納；V6副採工作範疇簡略說明"
                "（2026-09-07）代表性工作情境與工時特性頁。")
    return wrap(s)


# ============================================================== 4. AI 接手到哪裡
def coverage():
    _id[0] = 1
    s = header("三、AI 導入現況", "AI 已經接手到哪裡 — 全部落在輸入端，沒有一項能跨進判斷端")

    lx, lw = MARGIN_X, 5.85
    rx, rw = MARGIN_X + 6.24, 5.85

    s += tb(lx, 1.70, lw, 0.28, [[run("主採｜已導入 5 支 skill", 1250, 1, NAVY)]])
    s += shape(lx, 2.04, lw, 2.92, TINT, "roundRect")
    s += shape(lx, 2.04, 0.06, 2.92, NAVY)
    tools = [
        ("GU TA 用量自動核對填入", "填入 TA 的 Order Qty 欄"),
        ("Invoice 核對填 TA", "產出可填／衝突／模糊清單"),
        ("B/L 缸差比對", "產出核可狀態報告"),
        ("CBD → TA 產出工具", "轉換為 TA 格式"),
        ("海空運費比較 產出工具", "查詢海／空運報價，產出比較 Excel"),
    ]
    y = 2.20
    for a, b in tools:
        s += tb(lx + 0.26, y, 2.70, 0.30, [[run("· " + a, 1000, 1, NAVY)]], anchor="ctr")
        s += tb(lx + 3.04, y, lw - 3.30, 0.30, [[run(b, 950, 0, BODY)]], anchor="ctr")
        y += 0.36
    s += tb(lx + 0.26, 4.04, lw - 0.52, 0.76,
            [[run("對應的類別合計約占 25% 的信件量；十四類加權後的 AI 可協助比例為 12.6%"
                  "（該簡報自述約 10–15%）。", 950, 0, BODY)]], lnspc=128)

    s += tb(rx, 1.70, rw, 0.28, [[run("副採｜4 個潛在可介入點", 1250, 1, BLUE)]])
    s += shape(rx, 2.04, rw, 2.92, PANEL, "roundRect")
    s += shape(rx, 2.04, 0.06, 2.92, BLUE)
    pts = [
        ("PO／BOM 系統性資料 key-in", "by 款採購作業　20%"),
        ("空運費計算（by flow）", "空運費計算　5%"),
        ("出貨 TA 表格更新", "出貨 TA 更新　10%"),
        ("分廠數量計算", "中國產區分廠計算　5%"),
    ]
    y = 2.20
    for a, b in pts:
        s += tb(rx + 0.26, y, 3.10, 0.30, [[run("· " + a, 1000, 1, BLUE)]], anchor="ctr")
        s += tb(rx + 3.44, y, rw - 3.70, 0.30, [[run(b, 950, 0, BODY)]], algn="r", anchor="ctr")
        y += 0.36
    s += tb(rx + 0.26, 4.04, rw - 0.52, 0.76,
            [[run("這四個介入點分布在合計約 40% 的類別中（20＋5＋10＋5）——"
                  "但可自動化的只是這些類別裡的登錄與計算步驟。", 950, 0, BODY)]], lnspc=128)

    s += shape(MARGIN_X, 5.08, CONTENT_W, 0.76, REDBG, "roundRect")
    s += shape(MARGIN_X, 5.08, 0.06, 0.76, RED)
    s += tb(MARGIN_X + 0.24, 5.16, 2.05, 0.24, [[run("⚠ 別讀錯這兩個數字", 1050, 1, RED)]])
    s += tb(MARGIN_X + 2.40, 5.16, CONTENT_W - 2.64, 0.62,
            [[run("「25%」與「40%」講的是這些工具「碰得到的類別」占多少，不是「可以被取代的工時」。"
                  "工具接手的是類別裡的登錄與比對步驟；同一類別裡的議定、催辦、判讀、拍板，仍然逐件由人處理。",
                  950, 0, BODY)]], lnspc=125)

    s += takeaway(6.24, [
        run("→ ", 1300, 1, NAVY),
        run("九個項目全部落在「比對與填入」；要不要覆蓋、算不算核可、報告能不能寄出，", 1300, 1, INK),
        run("仍然是人按下確認鍵", 1300, 1, RED),
        run("。", 1300, 1, INK),
    ], h=0.58)
    s += source("資料來源：主採團隊現行 Claude Code skill 清單（bl-dyelot-check／ta-invoice-fill／"
                "gu-ta-orderqty-fill／gu-cbd-to-ta-tool，另含海空運費比較產出工具）；V6副採工作範疇簡略說明（2026-09-07）工時特性頁。"
                "25%、40%、12.6% 均為依表列占比推算之估算值。")
    return wrap(s)


# ============================================================== 5. 結論
def conclusion():
    _id[0] = 1
    s = header("四、結論與建議", "AI 是協作者，不是替代人力的方案 — 建議維持現有配置")

    s += shape(MARGIN_X, 1.70, CONTENT_W, 0.86, NAVY, "roundRect")
    s += tb(MARGIN_X + 0.34, 1.80, CONTENT_W - 0.68, 0.68,
            [[run("省下的 10–15% 不是集中在某一個人、某一段時間，而是散在六個人每天的 10–30 分鐘切片裡 —— "
                  "省下的是等待中的空隙，不是可以整併掉的一個編制。", 1300, 1, WHITE)]],
            anchor="ctr", lnspc=128)

    items = [
        ("① 繼續擴大，但目標放對", NAVY,
         "持續擴大結構化比對類工具的覆蓋（TA／Invoice／B/L 比對、分廠數量、空運費計算）。"
         "目標是把這 10–15% 的行政步驟做穩、降低 key-in 錯誤，而不是拿它折抵編制。"),
        ("② 這幾類不納入 AI 取代規劃", RED,
         "重大品質異常、議價協商（含匯率商談）、看色與核色判斷、跨廠催辦、跨國交涉 —— "
         "維持現有人力配置。這些不是工具成熟度問題，是責任歸屬問題。"),
        ("③ 之後要調整編制，看量體不看工具數", NAVY,
         "若日後評估人力，建議以款數／PO 張數／異常件數／製樣階段數等量體指標判斷，"
         "而不是以導入了幾支 AI 工具判斷 —— 兩者不成比例。"),
    ]
    y = 2.72
    for t, col, d in items:
        s += shape(MARGIN_X, y, CONTENT_W, 0.86, REDBG if col == RED else PANEL, "roundRect")
        s += shape(MARGIN_X, y, 0.06, 0.86, col)
        s += tb(MARGIN_X + 0.24, y + 0.09, 3.05, 0.24, [[run(t, 1050, 1, col)]])
        s += tb(MARGIN_X + 3.42, y + 0.09, CONTENT_W - 3.66, 0.68, [[run(d, 950, 0, BODY)]], lnspc=128)
        y += 0.94

    s += takeaway(5.66, [
        run("→ ", 1300, 1, NAVY),
        run("AI 讓同樣的六個人接得住愈來愈多的款數與異常 —— ", 1300, 1, INK),
        run("它撐住的是產能，不是騰出的編制", 1300, 1, RED),
        run("。", 1300, 1, INK),
    ], h=0.60)

    s += shape(MARGIN_X, 6.42, CONTENT_W, 0.62, PANEL, "roundRect")
    s += shape(MARGIN_X, 6.42, 0.06, 0.62, MUTED)
    s += tb(MARGIN_X + 0.24, 6.50, 1.70, 0.24, [[run("⚠ 數字口徑", 1000, 1, RED)]])
    s += tb(MARGIN_X + 2.05, 6.50, CONTENT_W - 2.29, 0.48,
            [[run("主採占比取自信箱關鍵字分類估算，樣本僅一人信箱可見範圍、非三人完整信箱、不含微信；"
                  "副採占比為體感估計。兩者皆非工時實測，正式引用前請主採／副採本人及主管複核。",
                  900, 0, BODY)]], lnspc=125)
    s += source("資料來源：主採業務範疇說明（2026-09-07）結論與建議頁；V6副採工作範疇簡略說明（2026-09-07）。")
    return wrap(s)


# ============================================================== package
SLIDES = [cover, scale, reasons, coverage, conclusion]


def main():
    if os.path.isdir(UN):
        shutil.rmtree(UN)
    zipfile.ZipFile(BASE).extractall(UN)

    # start from the base package but keep none of its slides or notes
    shutil.rmtree(os.path.join(UN, 'ppt/slides'))
    shutil.rmtree(os.path.join(UN, 'ppt/notesSlides'), ignore_errors=True)
    # the base deck's notes master and pasted bitmap have no referrer left
    shutil.rmtree(os.path.join(UN, 'ppt/notesMasters'), ignore_errors=True)
    shutil.rmtree(os.path.join(UN, 'ppt/media'), ignore_errors=True)
    os.remove(os.path.join(UN, 'ppt/theme/theme2.xml'))   # was the notes master's theme
    os.makedirs(os.path.join(UN, 'ppt/slides/_rels'))

    names = []
    for i, fn in enumerate(SLIDES, 1):
        n = 'slide%d.xml' % i
        names.append(n)
        open(os.path.join(UN, 'ppt/slides', n), 'w', encoding='utf-8').write(fn())
        open(os.path.join(UN, 'ppt/slides/_rels', n + '.rels'), 'w', encoding='utf-8').write(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')

    ctp = os.path.join(UN, '[Content_Types].xml')
    ct = open(ctp, encoding='utf-8').read()
    ct = re.sub(r'<Override PartName="/ppt/(slides|notesSlides)/[^"]*"[^>]*/>', '', ct)
    ct = re.sub(r'<Override PartName="/ppt/notesMasters/[^"]*"[^>]*/>', '', ct)
    ct = ct.replace('<Default Extension="png" ContentType="image/png"/>', '')
    ct = ct.replace('<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.'
                    'openxmlformats-officedocument.theme+xml"/>', '')
    ct = ct.replace('</Types>', ''.join(
        '<Override PartName="/ppt/slides/%s" ContentType="application/vnd.openxmlformats-'
        'officedocument.presentationml.slide+xml"/>' % n for n in names) + '</Types>')
    open(ctp, 'w', encoding='utf-8').write(ct)

    prp = os.path.join(UN, 'ppt/_rels/presentation.xml.rels')
    pr = open(prp, encoding='utf-8').read()
    pr = re.sub(r'<Relationship [^>]*Target="(slides/|notesMasters/)[^"]*"[^>]*/>', '', pr)
    used = {int(m) for m in re.findall(r'Id="rId(\d+)"', pr)}
    nxt = (max(used) if used else 0) + 1
    rids = []
    add = ''
    for n in names:
        rid = 'rId%d' % nxt; nxt += 1
        rids.append(rid)
        add += ('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/slide" Target="slides/%s"/>' % (rid, n))
    open(prp, 'w', encoding='utf-8').write(pr.replace('</Relationships>', add + '</Relationships>'))

    pxp = os.path.join(UN, 'ppt/presentation.xml')
    px = open(pxp, encoding='utf-8').read()
    px = re.sub(r'<p:notesMasterIdLst>.*?</p:notesMasterIdLst>', '', px, flags=re.S)
    lst = re.search(r'<p:sldIdLst>.*?</p:sldIdLst>', px, re.S).group(0)
    px = px.replace(lst, '<p:sldIdLst>' + ''.join(
        '<p:sldId id="%d" r:id="%s"/>' % (256 + i, r) for i, r in enumerate(rids)) + '</p:sldIdLst>')
    open(pxp, 'w', encoding='utf-8').write(px)

    if os.path.exists(OUT):
        os.remove(OUT)
    zf = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
    for root, _, files in os.walk(UN):
        for f in files:
            full = os.path.join(root, f)
            zf.write(full, os.path.relpath(full, UN))
    zf.close()
    print('wrote %s (%d slides)' % (OUT, len(names)))


main()
