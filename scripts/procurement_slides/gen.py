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

EYEBROW = "四、副採"
EYEBROW_M = "三、主採"

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


# ==================== SLIDE M1 : 主採工作範疇與量體 ====================
def slide_main_scope():
    _id[0] = 1
    s = header(EYEBROW_M, "主採業務的工作範疇 — 14 類工作，AI 可代勞程度全部評為「低」")

    stats = [
        ("1,993 封", "樣本期間信件總數", BLUE),
        ("14 類", "已盤點的工作類型", NAVY),
        ("100%", "全數評為「AI 可代勞程度：低」", RED),
        ("871 封", "最大宗：EC・樣品・APP 調料調布", NAVY),
    ]
    gap = 0.24
    sw = (CONTENT_W - gap * 3) / 4
    for i, (num, lab, col) in enumerate(stats):
        x = MARGIN_X + i * (sw + gap)
        s += shape(x, 1.70, sw, 0.88, REDBG if col == RED else PANEL, "roundRect")
        s += shape(x, 1.70, 0.06, 0.88, col)
        s += tb(x + 0.24, 1.79, sw - 0.44, 0.40, [[run(num, 2100, 1, col)]])
        s += tb(x + 0.24, 2.21, sw - 0.44, 0.30, [[run(lab, 900, 0, BODY)]], lnspc=120)

    # keep this label short: the 信件數／占比 column headers sit on the same baseline
    s += tb(MARGIN_X, 2.70, 4.00, 0.26,
            [[run("十四類工作的信件量與占比", 1250, 1, NAVY)]])

    items = [
        ("前端面料開發事宜", "350", "10%", False),
        ("EC／樣品／APP 調料調布需求", "871", "10%", False),
        ("進單前交期 TA", "307", "10%", False),
        ("FBO／CBD／主料 TA 確認", "200", "10%", False),
        ("重大品質異常", "82", "15%", True),
        ("換片率／短碼／補布／庫存／色差判定", "80", "7%", False),
        ("物流／出貨追蹤", "24", "5%", False),
        ("內部固定會議・報表", "—", "10%", False),
        ("驗貨／測試", "20", "5%", False),
        ("開訂單 PO", "16", "3%", False),
        ("收料／結帳", "393", "3%", False),
        ("缸差核可進度", "—", "10%", False),
        ("寄件", "—", "1%", False),
        ("測試帳單申請", "—", "1%", False),
    ]
    cw = (CONTENT_W - 0.30) / 2
    rh, vg, y0 = 0.38, 0.045, 3.02
    for idx, (name, cnt, pct, hot) in enumerate(items):
        col, row = divmod(idx, 7)
        x = MARGIN_X + col * (cw + 0.30)
        y = y0 + row * (rh + vg)
        s += shape(x, y, cw, rh, REDBG if hot else PANEL, "roundRect")
        s += shape(x, y, 0.06, rh, RED if hot else BLUE)
        s += tb(x + 0.22, y, cw - 2.05, rh, [[run(name, 1000, 1, INK)]], anchor="ctr")
        s += tb(x + cw - 1.72, y, 0.80, rh, [[run(cnt, 1000, 1, BODY)]], algn="r", anchor="ctr")
        s += tb(x + cw - 0.82, y, 0.60, rh, [[run(pct, 1000, 1, RED if hot else BLUE)]], algn="r", anchor="ctr")
    s += tb(MARGIN_X + cw - 1.72, y0 - 0.24, 0.80, 0.20, [[run("信件數", 850, 1, MUTED)]], algn="r")
    s += tb(MARGIN_X + cw - 0.82, y0 - 0.24, 0.60, 0.20, [[run("占比", 850, 1, MUTED)]], algn="r")
    s += tb(MARGIN_X + cw + 0.30 + cw - 1.72, y0 - 0.24, 0.80, 0.20, [[run("信件數", 850, 1, MUTED)]], algn="r")
    s += tb(MARGIN_X + cw + 0.30 + cw - 0.82, y0 - 0.24, 0.60, 0.20, [[run("占比", 850, 1, MUTED)]], algn="r")

    s += takeaway(6.19, [
        run("→ ", 1300, 1, NAVY),
        run("信件數最少的「重大品質異常」（82 封）占比卻最高（15%）—— ", 1300, 1, INK),
        run("工時不跟信件數走，跟事情的難度走", 1300, 1, RED),
        run("。", 1300, 1, INK),
    ], h=0.60)
    s += source("資料來源：主採工作內容分類與 AI 可行性分析（綜合分析表）。信件數為樣本期間統計值；占比為關鍵字分類之估算值，樣本與口徑限制見本章第三頁。")
    return wrap(s)

# ==================== SLIDE M2 : 為什麼 AI 代勞不了 ====================
def slide_main_why():
    _id[0] = 1
    s = header(EYEBROW_M, "為什麼十四類全評為「低」 — 六個共通的結構性原因")
    s += tb(MARGIN_X, 1.66, CONTENT_W, 0.26,
            [[run("不是工作內容單純，而是瓶頸都落在系統之外 —— 等待、協商、實物判斷", 1250, 1, NAVY)]])

    cards = [
        ("① 多輪來回議定，不是單向查詢",
         "TA 交期不是布廠給一個數字就定案：布廠產能、原料到位時間都要反覆協調，過程還包含人民幣匯率商談，屬多輪議價。"),
        ("② 一對多的跨廠協調",
         "同一款式通常同步在多家布廠開發，須持續追蹤各廠回覆與進度，再即時整合資訊回覆業務端，確保開發時程推進。"),
        ("③ 三方條件要同時滿足",
         "最終交期須同時滿足業務訂單需求、布廠產能與交期、成衣廠產線排程 —— 是多重限制下的權衡取捨，不是查表。"),
        ("④ 需實物比對與染整專業判斷",
         "移染、白汙、色牢度、色點及髒污比例的成因判定，仰賴布料／染整專業與目視實物、色板比對，並需送樣、複測、打樣驗證。"),
        ("⑤ 溝通載體不在系統裡",
         "缸差核色進度主要靠微信與上海布組即時互動，是零散對話與圖片核對；重大異常也多為電話、現場會議、口頭確認。"),
        ("⑥ 對外承諾要有人拍板",
         "下單即產生商業約束力；CBD／FBO 數字會隨業務端分次確認而滾動更新，需有人判斷「這次是否已是最終版本」，避免未定案就搶先下單。"),
    ]
    cols, gap = 3, 0.24
    cw = (CONTENT_W - gap * (cols - 1)) / cols
    ch, vgap, y0 = 1.50, 0.16, 2.02
    for idx, (title, desc) in enumerate(cards):
        r, c = divmod(idx, cols)
        x = MARGIN_X + c * (cw + gap)
        y = y0 + r * (ch + vgap)
        s += shape(x, y, cw, ch, PANEL, "roundRect")
        s += shape(x, y, 0.06, ch, BLUE)
        s += tb(x + 0.22, y + 0.13, cw - 0.44, 0.24, [[run(title, 1050, 1, BLUE)]])
        s += tb(x + 0.22, y + 0.44, cw - 0.44, ch - 0.58, [[run(desc, 900, 0, BODY)]], lnspc=128)

    s += shape(MARGIN_X, 5.38, CONTENT_W, 0.62, REDBG, "roundRect")
    s += shape(MARGIN_X, 5.38, 0.06, 0.62, RED)
    s += tb(MARGIN_X + 0.24, 5.45, 1.75, 0.24, [[run("★ 最關鍵的一點", 1050, 1, RED)]])
    s += tb(MARGIN_X + 2.10, 5.45, CONTENT_W - 2.34, 0.48,
            [[run("耗時的原因是「等待與多輪確認」，不是處理速度 —— AI 即使能快速產出初稿排程，也縮短不了人對人溝通確認的等待時間。", 1000, 1, BODY)]], lnspc=125)

    s += takeaway(6.24, [
        run("→ ", 1300, 1, NAVY),
        run("能被壓縮的是「產出速度」，但主採的工時花在", 1300, 1, INK),
        run("等別人回覆、和別人談定、對實物拍板", 1300, 1, RED),
        run(" —— 這三件事沒有一件會因為 AI 而變快。", 1300, 1, INK),
    ], h=0.60)
    s += source("資料來源：主採工作內容分類與 AI 可行性分析（綜合分析表）「為什麼」欄之歸納。")
    return wrap(s)

# ==================== SLIDE M3 : 已接手的與接不了的 ====================
def slide_main_ai():
    _id[0] = 1
    s = header(EYEBROW_M, "AI 已接手的與接不了的 — 以及這份數字的口徑")

    stats = [
        ("55 款 / 64 塊主布", "27SS 1st sample 前端開發量體", BLUE),
        ("28 款 / 24 塊布", "27FW 1st sample（目前）", NAVY),
        ("7 種", "須固定參與的例行會議與報表", NAVY),
    ]
    gap = 0.30
    sw = (CONTENT_W - gap * 2) / 3
    for i, (num, lab, col) in enumerate(stats):
        x = MARGIN_X + i * (sw + gap)
        s += shape(x, 1.70, sw, 0.82, PANEL, "roundRect")
        s += shape(x, 1.70, 0.06, 0.82, col)
        s += tb(x + 0.24, 1.79, sw - 0.44, 0.34, [[run(num, 1700, 1, col)]])
        s += tb(x + 0.24, 2.17, sw - 0.44, 0.28, [[run(lab, 900, 0, BODY)]], lnspc=120)

    lx, lw = MARGIN_X, 5.85
    rx, rw = MARGIN_X + 6.24, 5.85
    s += tb(lx, 2.68, lw, 0.26, [[run("✔ 已有現成 AI skill 輔助 —— 清單中可行性最高的三項", 1250, 1, NAVY)]])
    s += shape(lx, 3.00, lw, 2.20, TINT, "roundRect")
    s += shape(lx, 3.00, 0.06, 2.20, NAVY)
    tools = [
        ("GU TA 用量自動填入", "對應「FBO／CBD／主料 TA 確認」"),
        ("Invoice 核對填 TA", "對應「進單前交期 TA」"),
        ("B/L 缸差比對", "對應「物流／出貨追蹤」"),
    ]
    y = 3.14
    for a, b in tools:
        s += tb(lx + 0.26, y, 2.35, 0.28, [[run("· " + a, 1000, 1, NAVY)]], anchor="ctr")
        s += tb(lx + 2.68, y, lw - 2.94, 0.28, [[run(b, 950, 0, BODY)]], anchor="ctr")
        y += 0.34
    s += tb(lx + 0.26, y + 0.14, lw - 0.52, 0.72,
            [[run("這三類合計約占 25% 的信件量（依表列占比相加），但 AI 輔助的是「填表與核對」——"
                  "　交期議定、催促布廠訂艙、船班異動後重排產線，仍然逐件由人處理。", 950, 0, BODY)]], lnspc=128)

    s += tb(rx, 2.68, rw, 0.26, [[run("✖ 接不了的實例 —— 重大品質異常與判定", 1250, 1, RED)]])
    s += shape(rx, 3.00, rw, 2.20, REDBG, "roundRect")
    s += shape(rx, 3.00, 0.06, 2.20, RED)
    cases = [
        ("移染", "SANBEN 60236F014A #37 Brown、60226F034A #34 Brown；凱信 60146N118／60226N024B"),
        ("白汙", "GISTEX 226N009A／236N011A／226N033A；CENTEX 226N020B／236F007A"),
        ("色點及髒污比例", "176F036／F037／F038／F040A 換片率與色差判定"),
        ("後續責任", "須撰寫品質改善 CAP，對客戶說明成因與對策並承擔索賠、延誤出貨的商業風險"),
    ]
    y = 3.12
    for a, b in cases:
        s += tb(rx + 0.26, y, 1.25, 0.24, [[run(a, 1000, 1, RED)]])
        s += tb(rx + 1.58, y, rw - 1.84, 0.48, [[run(b, 950, 0, BODY)]], lnspc=125)
        y += 0.52

    s += shape(MARGIN_X, 5.40, CONTENT_W, 0.76, PANEL, "roundRect")
    s += shape(MARGIN_X, 5.40, 0.06, 0.76, MUTED)
    s += tb(MARGIN_X + 0.24, 5.48, 1.60, 0.24, [[run("⚠ 數字口徑", 1000, 1, RED)]])
    s += tb(MARGIN_X + 1.94, 5.48, CONTENT_W - 2.18, 0.62,
            [[run("樣本僅涵蓋本人信箱可見範圍（列為收件人／CC 或寄件人），非三人完整信箱，且不含微信與布廠溝通內容；"
                  "占比採關鍵字比對分類，同一封信可能同時符合多個類別，本表以優先序取單一類別歸類，與人工逐封判讀可能有落差。"
                  "實際工作分工與 AI 導入範圍請與三位本人及主管討論確認。", 900, 0, BODY)]], lnspc=125)

    s += takeaway(6.36, [
        run("→ ", 1300, 1, NAVY),
        run("已被 AI 接走的是「填表與核對」；真正的工時瓶頸 —— 交期議定、異常判定、跨廠催辦 —— ", 1300, 1, INK),
        run("一項都沒被接走", 1300, 1, RED),
        run("。", 1300, 1, INK),
    ], h=0.58)
    s += source("資料來源：主採工作內容分類與 AI 可行性分析（綜合分析表）之實際信件主旨舉例與備註欄。")
    return wrap(s)


# ==================== 副採: import the user's own 2026-09-07 deck ====================
FUCAI_SRC = "fucai_unpacked"          # unpacked copy of the 副採 deck
FUCAI_SLIDES = [2, 3, 4, 5]           # its content slides (1 is that deck's own cover)

# eyebrow rewritten so each page reads as part of this deck's 四、副採 section
FUCAI_EYEBROW = {
    2: "四、副採｜系統操作",
    3: "四、副採｜人力處理・判斷範疇",
    4: "四、副採｜代表性工作情境",
    5: "四、副採｜工時特性",
}
WHITE_BG = ('<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
            '<a:effectLst/></p:bgPr></p:bg>')

def _first_text_runs(xml):
    """Span of the eyebrow shape: the 2nd <p:sp> in the tree (1st is the left rail)."""
    i = xml.find('<p:sp>'); j = xml.find('</p:sp>', i)
    k = xml.find('<p:sp>', j); l = xml.find('</p:sp>', k)
    return k, l


def _fix_stat_row(xml):
    """The source deck's 系統操作 stat row is misaligned: card 1 and card 2 overlap by
    0.17", the three gaps differ (-0.17 / 0.37 / 0.10), and card 2's label is indented
    0.37" instead of 0.10". Re-lay the four cards on an even grid spanning the content
    column, and put every label back at card + 0.10"."""
    import re
    E = 914400.0
    pat = re.compile(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>')
    hits = [m for m in pat.finditer(xml)
            if 1.60 <= int(m.group(2)) / E <= 1.90 and 0.60 <= int(m.group(4)) / E <= 1.00]
    cards = [m for m in hits if int(m.group(4)) / E > 0.85]
    labels = [m for m in hits if int(m.group(4)) / E <= 0.85]
    if len(cards) != 4 or len(labels) != 4:
        return xml                      # source deck changed shape; leave it alone

    card_w = int(cards[0].group(3)) / E                      # 2.87"
    gap = (CONTENT_W - card_w * 4) / 3
    edits = {}
    for i, m in enumerate(cards):
        cx = MARGIN_X + i * (card_w + gap)
        edits[m.span()] = '<a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/>' % (
            emu(cx), emu(1.70), m.group(3), m.group(4))
    for i, m in enumerate(labels):
        cx = MARGIN_X + i * (card_w + gap) + 0.10
        edits[m.span()] = '<a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/>' % (
            emu(cx), emu(1.80), m.group(3), m.group(4))
    for (s, e) in sorted(edits, reverse=True):
        xml = xml[:s] + edits[(s, e)] + xml[e:]
    return xml

def fucai_slide(n):
    import re
    xml = open(os.path.join(FUCAI_SRC, "ppt/slides/slide%d.xml" % n), encoding="utf-8").read()

    # this deck's slides inherit their background from the layout; pin it white like the rest
    if "<p:bg>" not in xml:
        xml = xml.replace("<p:spTree>", WHITE_BG + "<p:spTree>", 1)
        # <p:bg> must precede <p:spTree> inside <p:cSld>
        xml = xml.replace("<p:cSld name=", "<p:cSld name=", 1)

    # retitle the eyebrow, collapsing its runs into one so the label reads cleanly
    k, l = _first_text_runs(xml)
    head = xml[k:l]
    runs = re.findall(r"<a:r>.*?</a:r>", head, re.S)
    if runs:
        first = runs[0]
        relabelled = re.sub(r"<a:t>.*?</a:t>", "<a:t>%s</a:t>" % FUCAI_EYEBROW[n], first, count=1, flags=re.S)
        newhead = head.replace("".join(runs), relabelled)
        xml = xml[:k] + newhead + xml[l:]

    # "上一頁" pointed at the source deck's own page order
    xml = xml.replace("見上一頁對照表", "見本章「系統操作」頁對照表")

    if n == 2:
        xml = _fix_stat_row(xml)
    return xml

# ==================== package assembly ====================
def main():
    root = "unpacked"
    slides = {"slide16.xml": slide_main_scope(), "slide17.xml": slide_main_why(),
              "slide18.xml": slide_main_ai()}
    for i, n in enumerate(FUCAI_SLIDES):
        slides["slide%d.xml" % (19 + i)] = fucai_slide(n)
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
