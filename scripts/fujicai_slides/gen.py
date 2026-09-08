# -*- coding: utf-8 -*-
"""Generate 副採 slides matching the existing V6 deck's visual language."""
import os, re, shutil, zipfile

E = 914400
def emu(v): return str(int(round(v * E)))

# ---- deck design tokens (read off slides 1/2/8/14) ----
NAVY   = "1C4F8F"
BLUE   = "2A78D6"
INK    = "0B0B0B"
BODY   = "454441"
MUTED  = "898781"
PANEL  = "F4F6F9"
TINT   = "E8F1FC"
RED    = "D03B3B"
REDBG  = "FDEFEF"
WHITE  = "FFFFFF"
FONT   = "微軟正黑體"

MARGIN_X, CONTENT_W = 0.62, 12.09

_id = [1]
def nid():
    _id[0] += 1
    return _id[0]

def esc(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))

def run(text, sz=1000, b=0, col=BODY, i=0):
    return dict(t=text, sz=sz, b=b, col=col, i=i)

def _runs_xml(runs):
    out = []
    for r in runs:
        sp = ' xml:space="preserve"' if r["t"] != r["t"].strip() else ""
        out.append(
            '<a:r><a:rPr lang="zh-TW" altLang="en-US" sz="%d" b="%d"%s dirty="0">'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:latin typeface="%s"/><a:ea typeface="%s"/><a:cs typeface="%s"/></a:rPr>'
            '<a:t%s>%s</a:t></a:r>'
            % (r["sz"], r["b"], ' i="1"' if r["i"] else "", r["col"], FONT, FONT, FONT,
               sp, esc(r["t"]))
        )
    return "".join(out)

def tb(x, y, w, h, paras, algn="l", anchor="t", lnspc=118):
    """paras: list of list-of-runs (each inner list is one paragraph)."""
    ps = []
    for runs in paras:
        ps.append(
            '<a:p><a:pPr algn="%s"><a:lnSpc><a:spcPct val="%d000"/></a:lnSpc>'
            '<a:spcBef><a:spcPts val="0"/></a:spcBef>'
            '<a:spcAft><a:spcPts val="0"/></a:spcAft></a:pPr>%s</a:p>'
            % (algn, lnspc, _runs_xml(runs))
        )
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="TextBox %d"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="%s"/>'
        '<a:lstStyle/>%s</p:txBody></p:sp>'
        % (nid(), _id[0], emu(x), emu(y), emu(w), emu(h), anchor, "".join(ps))
    )

def shape(x, y, w, h, fill, prst="rect"):
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="%d" name="Shape %d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
        '<a:prstGeom prst="%s"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>'
        '<a:p><a:pPr algn="ctr"/><a:endParaRPr/></a:p></p:txBody></p:sp>'
        % (nid(), _id[0], emu(x), emu(y), emu(w), emu(h), prst, fill)
    )

def header(eyebrow, title, eyecol=BLUE):
    s = shape(0, 0, 0.16, 7.5, NAVY)
    s += tb(MARGIN_X, 0.42, CONTENT_W, 0.32, [[run(eyebrow, 1250, 1, eyecol)]])
    s += tb(MARGIN_X, 0.76, CONTENT_W, 0.43, [[run(title, 2500, 1, INK)]], lnspc=110)
    s += shape(MARGIN_X, 1.46, CONTENT_W, 0.03, NAVY)
    return s

def takeaway(y, runs, h=0.62):
    s = shape(MARGIN_X, y, CONTENT_W, h, TINT, "roundRect")
    s += shape(MARGIN_X, y, 0.06, h, NAVY)
    s += tb(MARGIN_X + 0.24, y + 0.03, CONTENT_W - 0.48, h - 0.06, [runs], anchor="ctr")
    return s

def source(text):
    return tb(MARGIN_X, 7.18, CONTENT_W, 0.26, [[run(text, 850, 0, MUTED)]])

SLIDE_OPEN = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
    '<a:effectLst/></p:bgPr></p:bg><p:spTree>'
    '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
)
SLIDE_CLOSE = ('</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')

def wrap(body):
    _id[0] = 1
    return SLIDE_OPEN + body + SLIDE_CLOSE

EYEBROW = "三、副採"

# ==================== SLIDE A : 系統操作 ====================
def slide_systems():
    _id[0] = 1
    s = header(EYEBROW, "副採使用到的「系統」 — 不是大量集中，是零碎時間的持續累積")

    s += tb(MARGIN_X, 1.70, CONTENT_W, 0.28,
            [[run("逐款、逐品項、逐張 PO、逐個顏色的重複操作 —— 單次都不長，但每一款都要從頭跑一次", 1250, 1, NAVY)]])

    # table
    c1x, c1w = 0.76, 2.55
    c2x, c2w = 3.45, 1.45
    c3x, c3w = 5.08, 7.55
    hy, hh = 2.06, 0.44
    s += shape(MARGIN_X, hy, CONTENT_W, hh, NAVY)
    s += tb(c1x, hy, c1w, hh, [[run("系統", 1100, 1, WHITE)]], anchor="ctr")
    s += tb(c2x, hy, c2w, hh, [[run("單次耗時", 1100, 1, WHITE)]], anchor="ctr")
    s += tb(c3x, hy, c3w, hh, [[run("實際操作內容 / 需要人判斷的地方", 1100, 1, WHITE)]], anchor="ctr")

    rows = [
        ("ORDER SYSTEM", "視情況",
         "BOM 的正確性 —— 需請大貨業務儘快 BOM 轉檔或更正細節；BOM 不對，後段開立的 PO 全部要重來", False),
        ("ERP｜開立 PO", "10–15 分／款",
         "BY 款 BY 品項逐筆開立；一款 3～5～7 張 PO 不等，張數隨品項數而增加", False),
        ("ERP｜印 PO", "5–10 分／款",
         "BY PO 逐張列印，張數愈多耗時愈長", False),
        ("ERP｜檔案提供供應商", "30 分／款",
         "每次進單的全流程作業；另含挪料，收料極少", False),
        ("SHIPPING MONITOR", "3 分／筆",
         "每張 PO 每有出貨都要 Key；一款多張 PO、多次出貨即多筆，累加後不可忽視", False),
        ("電子表單｜異損申請", "視情況",
         "依異常狀況逐案填報，無固定頻率", False),
        ("實驗室品管", "逐色逐品項",
         "新款各品項各色的色牢度；含各測試的時效安排、申請書填寫與報告管理", False),
        ("BV 送測", "年度性",
         "各品項年度測試：有毒物質、貨櫃模擬、防鏽、寸法變化、拉力物性、色牢度＋PH 值＋游離甲醛…等；同樣含時效安排、申請書與報告管理", True),
    ]
    y = hy + hh
    rh = 0.44
    for i, (name, cost, desc, hot) in enumerate(rows):
        bg = WHITE if i % 2 == 0 else PANEL
        s += shape(MARGIN_X, y, CONTENT_W, rh, bg)
        s += tb(c1x, y, c1w, rh, [[run(name, 1050, 1, INK)]], anchor="ctr")
        s += tb(c2x, y, c2w, rh, [[run(cost, 1050, 1, RED if hot else BLUE)]], anchor="ctr")
        s += tb(c3x, y + 0.04, c3w, rh - 0.08, [[run(desc, 950, 0, BODY)]], anchor="ctr")
        y += rh

    s += takeaway(6.26, [
        run("→ ", 1300, 1, NAVY),
        run("這些系統操作都不「大」 —— 真正吃時間的是 ", 1300, 1, INK),
        run("次數", 1300, 1, RED),
        run("：一款要跑 3～7 張 PO、每張 PO 每次出貨再各 Key 一次。", 1300, 1, INK),
    ], h=0.60)
    s += source("資料來源：副採業務工作範疇彙整（2026-09-02 內部盤點）。耗時為實務估算區間，非系統紀錄值，對外引用前請人工複核。")
    return wrap(s)

# ==================== SLIDE B : 人力為主的協調作業 ====================
def slide_manual():
    _id[0] = 1
    s = header(EYEBROW, "人力為主＋微 AI 協助為輔 — 十一項無法交給系統的協調作業")
    s += tb(MARGIN_X, 1.66, CONTENT_W, 0.26,
            [[run("以下每一項都要「對人」—— 供應商、成衣廠、駐廠、開發業務、大貨業務、樣品中心，沒有一項是把資料丟進系統就會有答案", 1250, 1, NAVY)]])

    items = [
        ("① 各製樣階段 樣品調料安排",
         "初樣／各色／RE 各色／MONITOR／產前樣，逐階段與各供應商互動調料；單號通知樣品中心＆開發；報價往返於供應商＆開發"),
        ("② TA 計劃可行性互動",
         "與各開發業務多次互動、確認 TA 計劃（含變化）的可行性，並持續追蹤 CBD DATE 是否按時提供 CBD"),
        ("③ by flow 空運費計算",
         "BY 品項 × 件數 × 重量，逐段推算「從哪到哪」的預計空運費；flow 一改就要重算"),
        ("④ 打色／送核／核色",
         "各品項打色、送核、核色與紀錄管理；整理核可副料卡，再寄送至二級產區"),
        ("⑤ 採購與交期管理",
         "採購數量計算／製表／推算交期／下單／交期管理／看色／出貨安排，一條龍由同一人接完"),
        ("⑥ 出貨 TA 更新",
         "各自表格分別管理與更新；供料 TA 沒有 color breakdown，視覺上不易判讀，只能逐筆核對"),
        ("⑦ 訂單切割重整",
         "因產能調度而切割訂單時，各品項數量要切割重計、開新 PO、交期整批重整"),
        ("⑧ 中國產區分廠計算",
         "每次待上線前 PO 出來才能分廠：按配色／用量／損耗／差數重新分配整理分廠數量，再逐一更新給各供應商"),
        ("⑨ 補料安排",
         "多裁多出的補料安排，以及用量爭議的補料協調 —— 責任歸屬談定前無法下單"),
        ("⑩ 品質 ISSUE 釐清與處理",
         "與成衣廠／駐廠／供應商／大貨四方互動，釐清狀況並收斂到可執行的處理方式"),
        ("⑪ 例行會議與供應商來訪",
         "參與各類 routine 會議，以及供應商來訪的接待與議題對應"),
    ]
    cols, gap = 3, 0.24
    cw = (CONTENT_W - gap * (cols - 1)) / cols
    ch, vgap = 0.96, 0.09
    y0 = 2.04
    for idx, (title, desc) in enumerate(items):
        r, c = divmod(idx, cols)
        x = MARGIN_X + c * (cw + gap)
        y = y0 + r * (ch + vgap)
        s += shape(x, y, cw, ch, PANEL, "roundRect")
        s += shape(x, y, 0.06, ch, BLUE)
        s += tb(x + 0.22, y + 0.09, cw - 0.44, 0.21, [[run(title, 1050, 1, BLUE)]])
        s += tb(x + 0.22, y + 0.33, cw - 0.44, ch - 0.43, [[run(desc, 900, 0, BODY)]], lnspc=125)

    # 12th cell: red accent note (matches the deck's ⑦ 例外應變 treatment)
    r, c = divmod(len(items), cols)
    x = MARGIN_X + c * (cw + gap)
    y = y0 + r * (ch + vgap)
    s += shape(x, y, cw, ch, REDBG, "roundRect")
    s += shape(x, y, 0.06, ch, RED)
    s += tb(x + 0.22, y + 0.09, cw - 0.44, 0.21, [[run("★ 共同點", 1050, 1, RED)]])
    s += tb(x + 0.22, y + 0.33, cw - 0.44, ch - 0.43,
            [[run("全部是逐款、逐品項、逐色的個別協調，沒有固定公式可套；AI 目前只能在制式表單上微量協助。", 900, 0, BODY)]], lnspc=125)

    s += takeaway(6.38, [
        run("→ ", 1300, 1, NAVY),
        run("系統只負責「登錄結果」—— 結果長什麼樣，是這十一項人工協調談出來的。", 1300, 1, INK),
    ], h=0.60)
    s += source("資料來源：副採業務工作範疇彙整（2026-09-02 內部盤點）。")
    return wrap(s)

# ==================== SLIDE C : 工時樣態 ====================
def slide_hours():
    _id[0] = 1
    s = header(EYEBROW, "這些作業的工時投入 — 不是大塊時間，是零碎累積且併行發生")

    stats = [
        ("10–30 分", "每項作業的典型耗時", BLUE),
        ("3～7 張", "一款需開立的 PO 張數", NAVY),
        ("45–55 分", "一款進單、單就 ERP 系統面的合計", BLUE),
        ("3 分 / 筆", "每個 shipment 的 Key-in", NAVY),
    ]
    gap = 0.24
    sw = (CONTENT_W - gap * 3) / 4
    for i, (num, lab, col) in enumerate(stats):
        x = MARGIN_X + i * (sw + gap)
        s += shape(x, 1.70, sw, 0.92, PANEL, "roundRect")
        s += shape(x, 1.70, 0.06, 0.92, col)
        s += tb(x + 0.24, 1.80, sw - 0.44, 0.42, [[run(num, 2200, 1, col)]])
        s += tb(x + 0.24, 2.24, sw - 0.44, 0.30, [[run(lab, 950, 0, BODY)]], lnspc=120)

    # left: quantifiable
    lx, lw = MARGIN_X, 5.85
    rx, rw = MARGIN_X + 6.24, 5.85
    s += tb(lx, 2.82, lw, 0.28, [[run("① 可量化的部分 —— 系統操作", 1250, 1, NAVY)]])
    s += shape(lx, 3.16, lw, 2.96, PANEL, "roundRect")
    s += shape(lx, 3.16, 0.06, 2.96, BLUE)
    rows = [
        ("開立 PO（BY 款 BY 品項）", "10–15 分／款"),
        ("印 PO（BY PO 逐張）", "5–10 分／款"),
        ("提供檔案給各供應商（每次進單全流程）", "30 分／款"),
        ("SHIPPING MONITOR 每筆 shipment Key-in", "3 分／筆"),
        ("實驗室品管／BV 送測 的申請書與報告管理", "視測試安排"),
        ("電子表單 異損申請", "視情況"),
    ]
    y = 3.30
    for i, (a, b) in enumerate(rows):
        s += tb(lx + 0.26, y, lw - 2.05, 0.30, [[run("· " + a, 1000, 0, BODY)]], anchor="ctr", lnspc=120)
        s += tb(lx + lw - 1.75, y, 1.50, 0.30, [[run(b, 1000, 1, BLUE)]], algn="r", anchor="ctr")
        y += 0.36
    s += tb(lx + 0.26, y + 0.06, lw - 0.52, 0.52,
            [[run("合計：一款進單，單就 ERP 三段作業即約 45–55 分（依上述區間相加），尚未計入 shipment Key-in 與後續異動重跑。", 950, 1, RED)]], lnspc=125)

    # right: non-quantifiable
    s += tb(rx, 2.82, rw, 0.28, [[run("② 不可量化的部分 —— 為什麼「加總」會失真", 1250, 1, NAVY)]])
    s += shape(rx, 3.16, rw, 2.96, REDBG, "roundRect")
    s += shape(rx, 3.16, 0.06, 2.96, RED)
    notes = [
        "其間會有各項事務同時發生中 —— emails／Teams／電話／會議／供應商來訪，作業經常被中斷後再接回",
        "這些細節不像做報告那種「跑資料＋分析」可以一次做完，而是零碎時間的持續累積",
        "同一款的工時會隨 PO 張數、品項數、顏色數倍增 —— 款數不能代表 loading",
        "訂單切割、分廠重算、TA 變動時，前面已完成的計算與 PO 要整批重跑一次",
        "被中斷後重新接回原本的判斷脈絡，本身就是成本，且不會出現在任何工時紀錄上",
    ]
    y = 3.32
    for n in notes:
        s += tb(rx + 0.26, y, rw - 0.52, 0.52, [[run("· " + n, 1000, 0, BODY)]], lnspc=125)
        y += 0.56

    s += takeaway(6.36, [
        run("→ ", 1300, 1, NAVY),
        run("副採的工時不會集中出現在任何一張報表上 —— 它分散在每一款、每一張 PO、每一個顏色的重複操作與即時協調裡。", 1300, 1, INK),
    ], h=0.60)
    s += source("資料來源：副採業務工作範疇彙整（2026-09-02 內部盤點）。耗時為實務估算區間，合計為區間相加之推算值，非系統紀錄，對外引用前請人工複核。")
    return wrap(s)

# ==================== package assembly ====================
def main():
    root = "unpacked"
    slides = {"slide16.xml": slide_systems(), "slide17.xml": slide_manual(), "slide18.xml": slide_hours()}
    for name, xml in slides.items():
        open(os.path.join(root, "ppt/slides", name), "w", encoding="utf-8").write(xml)
        open(os.path.join(root, "ppt/slides/_rels", name + ".rels"), "w", encoding="utf-8").write(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
            'Target="../slideLayouts/slideLayout7.xml"/></Relationships>')

    # content types
    ctp = os.path.join(root, "[Content_Types].xml")
    ct = open(ctp, encoding="utf-8").read()
    add = "".join(
        '<Override PartName="/ppt/slides/%s" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' % n
        for n in slides)
    ct = ct.replace("</Types>", add + "</Types>")
    open(ctp, "w", encoding="utf-8").write(ct)

    # presentation rels
    prp = os.path.join(root, "ppt/_rels/presentation.xml.rels")
    pr = open(prp, encoding="utf-8").read()
    used = {int(m) for m in re.findall(r'Id="rId(\d+)"', pr)}
    newr = {}
    nxt = max(used) + 1
    for n in slides:
        newr[n] = "rId%d" % nxt
        nxt += 1
    add = "".join(
        '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/%s"/>' % (newr[n], n)
        for n in slides)
    pr = pr.replace("</Relationships>", add + "</Relationships>")
    open(prp, "w", encoding="utf-8").write(pr)

    # presentation.xml: insert after the 10th slide entry (大貨 section end)
    pxp = os.path.join(root, "ppt/presentation.xml")
    px = open(pxp, encoding="utf-8").read()
    lst = re.search(r"<p:sldIdLst>(.*?)</p:sldIdLst>", px, re.S).group(1)
    entries = re.findall(r"<p:sldId [^/]*/>", lst)
    maxid = max(int(m) for m in re.findall(r'id="(\d+)"', lst))
    ins = []
    for i, n in enumerate(slides):
        maxid += 1
        ins.append('<p:sldId id="%d" r:id="%s"/>' % (maxid, newr[n]))
    entries = entries[:10] + ins + entries[10:]
    px = px.replace("<p:sldIdLst>" + lst + "</p:sldIdLst>", "<p:sldIdLst>" + "".join(entries) + "</p:sldIdLst>")
    open(pxp, "w", encoding="utf-8").write(px)
    print("inserted", list(slides), "->", newr)

main()
