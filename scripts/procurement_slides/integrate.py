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



def source2(text):
    """Two-line source line — this 2-pager carries both decks' 口徑 disclaimers,
    which do not fit the family's single 0.26" line."""
    return tb(MARGIN_X, 7.04, CONTENT_W, 0.38, [[run(text, 850, 0, MUTED)]], lnspc=122)


# ============================================================== 1. 現況
def status():
    _id[0] = 1
    s = header("主副料採購 ‧ AI 導入評估　①",
               "AI 接手的是輸入端 — 量體、已導入的工具，以及它們接手到哪裡")

    stats = [("25 類", "主副採已盤點的工作類型（14＋11）", NAVY),
             ("6 位", "人力編制（主採 3＋副採 3）", NAVY),
             ("12.6%", "主採 AI 可協助比例（占比加權）", BLUE),
             ("9 項", "已導入 skill ＋ 潛在可介入點", BLUE)]
    gap = 0.24
    sw = (CONTENT_W - gap * 3) / 4
    for i, (num, lab, col) in enumerate(stats):
        x0 = MARGIN_X + i * (sw + gap)
        s += shape(x0, 1.68, sw, 0.86, PANEL, "roundRect")
        s += shape(x0, 1.68, 0.06, 0.86, col)
        s += tb(x0 + 0.24, 1.77, sw - 0.44, 0.40, [[run(num, 2100, 1, col)]])
        s += tb(x0 + 0.24, 2.19, sw - 0.44, 0.28, [[run(lab, 900, 0, BODY)]], lnspc=120)

    lx, lw = MARGIN_X, 5.85
    rx, rw = MARGIN_X + 6.24, 5.85
    panel_y, panel_h, note_y = 3.02, 2.72, 4.98

    s += tb(lx, 2.68, lw, 0.26, [[run("主採｜已導入 5 支 skill", 1250, 1, NAVY)]])
    s += shape(lx, panel_y, lw, panel_h, TINT, "roundRect")
    s += shape(lx, panel_y, 0.06, panel_h, NAVY)
    tools = [("GU TA 用量自動核對填入", "填入 TA 的 Order Qty 欄"),
             ("Invoice 核對填 TA", "產出可填／衝突／模糊清單"),
             ("B/L 缸差比對", "產出核可狀態報告"),
             ("CBD → TA 產出工具", "轉換為 TA 格式"),
             ("海空運費比較 產出工具", "查詢海／空運報價，產出比較 Excel")]
    y = panel_y + 0.14
    for a, b in tools:
        s += tb(lx + 0.26, y, 2.70, 0.30, [[run("· " + a, 1000, 1, NAVY)]], anchor="ctr")
        s += tb(lx + 3.04, y, lw - 3.30, 0.30, [[run(b, 900, 0, BODY)]], anchor="ctr")
        y += 0.34
    s += tb(lx + 0.26, note_y, lw - 0.52, 0.62,
            [[run("對應的類別合計約占 25% 的信件量；十四類加權後的 AI 可協助比例為 12.6%"
                  "（該簡報自述約 10–15%）。", 900, 0, BODY)]], lnspc=128)

    s += tb(rx, 2.68, rw, 0.26, [[run("副採｜4 個潛在可介入點", 1250, 1, BLUE)]])
    s += shape(rx, panel_y, rw, panel_h, PANEL, "roundRect")
    s += shape(rx, panel_y, 0.06, panel_h, BLUE)
    pts = [("PO／BOM 系統性資料 key-in", "by 款採購作業　20%"),
           ("空運費計算（by flow）", "空運費計算　5%"),
           ("出貨 TA 表格更新", "出貨 TA 更新　10%"),
           ("分廠數量計算", "中國產區分廠計算　5%")]
    y = panel_y + 0.14
    for a, b in pts:
        s += tb(rx + 0.26, y, 3.10, 0.30, [[run("· " + a, 1000, 1, BLUE)]], anchor="ctr")
        s += tb(rx + 3.44, y, rw - 3.70, 0.30, [[run(b, 900, 0, BODY)]], algn="r", anchor="ctr")
        y += 0.34
    s += tb(rx + 0.26, note_y, rw - 0.52, 0.62,
            [[run("分布在合計約 40% 的類別中（20＋5＋10＋5）；可自動化的只是這些類別裡的"
                  "登錄與計算步驟，不是整類工作。", 900, 0, BODY)]], lnspc=128)

    s += shape(MARGIN_X, 5.92, CONTENT_W, 0.86, REDBG, "roundRect")
    s += shape(MARGIN_X, 5.92, 0.06, 0.86, RED)
    s += tb(MARGIN_X + 0.24, 6.02, 2.30, 0.24, [[run("⚠ 別讀錯這兩個數字", 1050, 1, RED)]])
    s += tb(MARGIN_X + 2.66, 6.02, CONTENT_W - 2.90, 0.70,
            [[run("「25%」與「40%」講的是這些工具「碰得到的類別」占多少，不是「可以被取代的工時」。", 1000, 1, BODY)],
             [run("九個項目全部落在「比對與填入」—— 要不要覆蓋、算不算核可、報告能不能寄出，仍然是人按下確認鍵。",
                  1000, 0, BODY)]], lnspc=128)

    s += source2("資料來源：主採業務範疇說明（2026-09-07）；V6副採工作範疇簡略說明（2026-09-07）。"
                "12.6%／25%／40% 均為依表列占比推算之估算值，非工時實測。"
                "主採占比取自信箱關鍵字分類（樣本僅一人信箱可見範圍、不含微信），副採占比為體感估計。")
    return wrap(s)


# ============================================================== 2. 原因與建議
def why():
    _id[0] = 1
    s = header("主副料採購 ‧ AI 導入評估　②",
               "為什麼工時省得下來，人力卻減不掉 — 六個共通原因")
    s += tb(MARGIN_X, 1.62, CONTENT_W, 0.26,
            [[run("主採談布廠、副採談副料供應商，對象不同，但卡住的地方一模一樣", 1250, 1, NAVY)]])

    cards = [
        ("① 是「議定」，不是「查詢」",
         "交期、價格、責任都要多輪來回談定，主採還含人民幣匯率商談。AI 產得出初稿，談不出共識。"),
        ("② 等待不會因為 AI 而變快",
         "工時花在等布廠、成衣廠、客戶回覆。就算報表 3 秒生成，對方明天才回 —— 處理速度從來不是瓶頸。"),
        ("③ 要看實物才能拍板",
         "色差、移染、白汙、看色核色，都要目視實物與色板比對，並送樣、複測、打樣驗證。"),
        ("④ 溝通載體不在系統裡",
         "缸差核色靠微信、重大異常靠電話與現場會議。零散對話與圖片，AI 讀不到也追不了。"),
        ("⑤ 承諾與責任要人來扛",
         "下單即產生商業約束力；異常究責、換不換標、CAP 報告都要有人簽名。責任無法交給工具。"),
        ("⑥ 資料會回頭改寫上一輪的答案",
         "訂單切割、分廠重算、TA 變動、PH3 填補 —— 前一刻算出的答案，隨時被工廠或供應商的回覆推翻。"),
    ]
    cols, gap = 3, 0.24
    cw = (CONTENT_W - gap * (cols - 1)) / cols
    ch, vgap, y0 = 1.16, 0.12, 1.96
    for i, (t, d) in enumerate(cards):
        r, c = divmod(i, cols)
        x0 = MARGIN_X + c * (cw + gap)
        y = y0 + r * (ch + vgap)
        s += shape(x0, y, cw, ch, PANEL, "roundRect")
        s += shape(x0, y, 0.06, ch, BLUE)
        s += tb(x0 + 0.22, y + 0.11, cw - 0.44, 0.22, [[run(t, 1050, 1, BLUE)]])
        s += tb(x0 + 0.22, y + 0.39, cw - 0.44, ch - 0.52, [[run(d, 900, 0, BODY)]], lnspc=128)

    s += shape(MARGIN_X, 4.52, CONTENT_W, 0.74, NAVY, "roundRect")
    s += tb(MARGIN_X + 0.32, 4.60, CONTENT_W - 0.64, 0.58,
            [[run("主採最大的一類占 15%、副採最大的兩類各占 20% —— 沒有一塊大到值得整批自動化；"
                  "而省下的 10–15% 散在六個人每天的 10–30 分鐘切片裡，", 1200, 1, WHITE),
              run("是等待中的空隙，不是可以整併掉的一個編制。", 1200, 1, "FFD9D9")]],
            anchor="ctr", lnspc=128)

    advice = [
        ("① 繼續擴大，但目標放對", NAVY,
         "擴大結構化比對類工具的覆蓋；目標是把這 10–15% 的行政步驟做穩、降低 key-in 錯誤，而不是拿它折抵編制。"),
        ("② 這幾類不納入 AI 取代規劃", RED,
         "重大品質異常、議價協商（含匯率商談）、看色與核色判斷、跨廠催辦、跨國交涉 —— 維持現有人力配置。"),
        ("③ 評估編制看量體，不看工具數", NAVY,
         "日後若要評估人力，以款數／PO 張數／異常件數等量體指標判斷，而不是以導入了幾支 AI 工具判斷。"),
    ]
    y = 5.40
    for t, col, d in advice:
        s += shape(MARGIN_X, y, CONTENT_W, 0.44, REDBG if col == RED else PANEL, "roundRect")
        s += shape(MARGIN_X, y, 0.06, 0.44, col)
        s += tb(MARGIN_X + 0.24, y, 3.05, 0.44, [[run(t, 1000, 1, col)]], anchor="ctr")
        s += tb(MARGIN_X + 3.42, y, CONTENT_W - 3.66, 0.44, [[run(d, 950, 0, BODY)]], anchor="ctr")
        y += 0.50

    s += source2("結論：AI 是協作者，不是替代人力的方案 —— 它撐住的是產能，不是騰出的編制。｜"
                "資料來源：主採業務範疇說明（2026-09-07）、V6副採工作範疇簡略說明（2026-09-07）。"
                "占比皆為估算值、非工時實測，正式引用前請主採／副採本人及主管複核。")
    return wrap(s)


# ============================================================== package
SLIDES = [status, why]


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
