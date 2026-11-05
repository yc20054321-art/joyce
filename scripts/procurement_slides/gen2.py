# -*- coding: utf-8 -*-
"""Append 05/06 slides (職能影響・職等分工・跨單位反饋) to the AI 導入評估 deck.

Content source: Joyce's mail「ai 與人力議題」(.msg), split by 主採／副採／駐廠業務.
駐廠's 06 row comes from「AI提案 by 駐廠業務 20261104.xlsx」(that mail carries no 06
for 駐廠); it is marked on the slide.
"""
import os, re, shutil, zipfile

E = 914400
def emu(v): return str(int(round(v * E)))

NAVY, BLUE, INK, BODY, MUTED = "1C4F8F", "2A78D6", "0B0B0B", "454441", "898781"
PANEL, TINT, RED, REDBG, WHITE = "F4F6F9", "E8F1FC", "D03B3B", "FDEFEF", "FFFFFF"
FONT = "微軟正黑體"
MARGIN_X, CONTENT_W = 0.62, 12.09
SRC_Y, SRC_H = 7.04, 0.38


import unicodedata
def _wide(c): return unicodedata.east_asian_width(c) in ('W', 'F')
def _wpt(s, sz): return sum(sz if _wide(c) else sz * 0.52 for c in s)

def nlines(text, sz_pt, w_in):
    """How many lines `text` needs at sz_pt inside w_in inches."""
    avail = w_in * 72.0
    return max(1, int(_wpt(text, sz_pt) / avail) + (1 if _wpt(text, sz_pt) % avail else 0))

def th(text, sz_pt, w_in, lnspc=126):
    """Height in inches that `text` needs."""
    return nlines(text, sz_pt, w_in) * sz_pt * (lnspc / 100.0) / 72.0

_id = [1]
def nid():
    _id[0] += 1
    return _id[0]

def esc(t): return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def run(text, sz=1000, b=0, col=BODY):
    return dict(t=text, sz=sz, b=b, col=col)

def _runs(rs):
    out = []
    for r in rs:
        sp = ' xml:space="preserve"' if r["t"] != r["t"].strip() else ""
        out.append(
            '<a:r><a:rPr lang="zh-TW" altLang="en-US" sz="%d" b="%d" dirty="0">'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:latin typeface="%s"/>'
            '<a:ea typeface="%s"/><a:cs typeface="%s"/></a:rPr><a:t%s>%s</a:t></a:r>'
            % (r["sz"], r["b"], r["col"], FONT, FONT, FONT, sp, esc(r["t"])))
    return "".join(out)

def tb(x, y, w, h, paras, algn="l", anchor="t", lnspc=118):
    ps = "".join(
        '<a:p><a:pPr algn="%s"><a:lnSpc><a:spcPct val="%d000"/></a:lnSpc>'
        '<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft>'
        '</a:pPr>%s</a:p>' % (algn, lnspc, _runs(rs)) for rs in paras)
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="TextBox %d"/><p:cNvSpPr txBox="1"/>'
            '<p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/>'
            '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="%s"/>'
            '<a:lstStyle/>%s</p:txBody></p:sp>'
            % (nid(), _id[0], emu(x), emu(y), emu(w), emu(h), anchor, ps))

def shape(x, y, w, h, fill, prst="rect"):
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="Shape %d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/></a:xfrm>'
            '<a:prstGeom prst="%s"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="%s"/>'
            '</a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></p:spPr><p:txBody>'
            '<a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            '<a:p><a:pPr algn="ctr"/><a:endParaRPr/></a:p></p:txBody></p:sp>'
            % (nid(), _id[0], emu(x), emu(y), emu(w), emu(h), prst, fill))

def header(eyebrow, title, tsz=2500):
    s = shape(0, 0, 0.16, 7.5, NAVY)
    s += tb(MARGIN_X, 0.42, CONTENT_W, 0.32, [[run(eyebrow, 1250, 1, BLUE)]])
    s += tb(MARGIN_X, 0.76, CONTENT_W, 0.46, [[run(title, tsz, 1, INK)]], lnspc=110)
    s += shape(MARGIN_X, 1.46, CONTENT_W, 0.03, NAVY)
    return s

def takeaway(y, rs, h=0.58):
    s = shape(MARGIN_X, y, CONTENT_W, h, TINT, "roundRect")
    s += shape(MARGIN_X, y, 0.06, h, NAVY)
    s += tb(MARGIN_X + 0.24, y + 0.03, CONTENT_W - 0.48, h - 0.06, [rs], anchor="ctr")
    return s

def source(text, lnspc=122):
    return tb(MARGIN_X, SRC_Y, CONTENT_W, SRC_H, [[run(text, 850, 0, MUTED)]], lnspc=lnspc)

OPEN = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:effectLst/></p:bgPr></p:bg><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>')
CLOSE = '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'

def wrap(body):
    return OPEN + body + CLOSE


EYE = "主副料採購 ‧ AI 導入評估　"

# ====================================================== ③ 職能影響與工時釋出
def slide_impact():
    _id[0] = 1
    s = header(EYE + "③", "AI 導入後的職能影響 — 主採・副採・駐廠")
    s += tb(MARGIN_X, 1.64, CONTENT_W, 0.26,
            [[run("三個單位的工時釋出幅度都落在 15% 上下；釋出的時間轉入議價、預防與異常判斷", 1250, 1, NAVY)]])

    units = [
        ("主採", NAVY, "10–15%", "預估可釋出工時", [
            ("減少的工作", "單據核對、交期催詢、報表彙整、手動 key in 等重複性工作。"),
            ("加重的工作", "數據分析、供應商合作與議價、前端品質控管、面料開發、異常判斷。"),
            ("職能轉型", "從「下單跟單」轉為「供應鏈管理、全流程品質控管、異常處理與成本優化」。"),
        ]),
        ("副採", BLUE, "15–20%", "約每月 26–35 小時（1 人 × 8 小時 × 22 天推估）", [
            ("角色轉變", "從時間有限之下的執行者，提升為副料的成本、交期與品質負責人。"),
            ("成本", "詢價、比價、議價力道加大。"),
            ("交期", "密切追蹤前端開發及核可進度，及時啟動生產或預防交期延誤。"),
            ("品質", "異常從單件處理改為追源頭，事前預警，減少上線前／中／後的急況釐清與損失。"),
        ]),
        ("駐廠業務", NAVY, "15%", "預估可釋出工時", [
            ("工作型態", "主要工作範圍仍以現場為主。"),
            ("組織", "員額與組織架構暫不變動。"),
            ("AI 著力點", "巡廠報告彙整翻譯、樣品費 invoice 檢查、KAKEN 布測檢查、缸差表與到料追蹤、"
                        "全流程時間點提醒、packing list 自動比對。"),
        ]),
    ]
    gap = 0.24
    cw = (CONTENT_W - gap * 2) / 3
    for i, (name, col, pct, pct_lab, rows) in enumerate(units):
        x = MARGIN_X + i * (cw + gap)
        s += tb(x, 1.98, cw, 0.26, [[run(name, 1250, 1, col)]])
        s += shape(x, 2.30, cw, 3.78, PANEL, "roundRect")
        s += shape(x, 2.30, 0.06, 3.78, col)
        s += tb(x + 0.24, 2.42, cw - 0.48, 0.42, [[run(pct, 2200, 1, col)]])
        s += tb(x + 0.24, 2.86, cw - 0.48, 0.34, [[run(pct_lab, 900, 0, MUTED)]], lnspc=122)
        y, iw = 3.30, cw - 0.48
        for lab, txt in rows:
            lh = th(lab, 9.5, iw, 124)
            bh = th(txt, 9.0, iw, 126)
            s += tb(x + 0.24, y, iw, lh, [[run(lab, 950, 1, col)]], lnspc=124)
            s += tb(x + 0.24, y + lh + 0.04, iw, bh, [[run(txt, 900, 0, BODY)]], lnspc=126)
            y += lh + bh + 0.18

    s += takeaway(6.28, [
        run("→ ", 1300, 1, NAVY),
        run("工時佔比最高的仍是「與人溝通」——  釋出的 15% 上下是轉入議價、預防與判斷，", 1300, 1, INK),
        run("屬於職能升級，而非可縮減的編制", 1300, 1, RED),
        run("。", 1300, 1, INK),
    ])
    s += source("資料來源：主採／副採／駐廠業務 05 回饋彙整（2026-11）。工時釋出為各單位自行評估之推估值，非工時實測。")
    return wrap(s)


def table(x, y, cols, rows, head_h=0.40, pad=0.12, fill_to=None):
    """cols: [(width, title)]; rows: [[cell, ...], ...] with cell = str or (str, bold, colour).
    Row heights are measured from content, never guessed."""
    total_w = sum(w for w, _ in cols)
    s = shape(x, y, total_w, head_h, NAVY)
    cx = x
    for w, t in cols:
        s += tb(cx + 0.14, y, w - 0.28, head_h, [[run(t, 1050, 1, WHITE)]], anchor="ctr")
        cx += w
    heights = [pad + max(th(c[0] if isinstance(c, tuple) else c,
                            9.5 if (isinstance(c, tuple) and c[1]) else 9.0, w - 0.28)
                         for (w, _), c in zip(cols, cells)) for cells in rows]
    if fill_to:                       # spread the slack so the table fills the page
        slack = fill_to - (y + head_h + sum(heights))
        if slack > 0:
            heights = [h + slack / len(heights) for h in heights]
    yy = y + head_h
    for i, (cells, h) in enumerate(zip(rows, heights)):
        s += shape(x, yy, total_w, h, WHITE if i % 2 == 0 else PANEL)
        cx = x
        for (w, _), cell in zip(cols, cells):
            txt, b, col = cell if isinstance(cell, tuple) else (cell, 0, BODY)
            s += tb(cx + 0.14, yy + pad / 2, w - 0.28, h - pad,
                    [[run(txt, 950 if b else 900, b, col)]], anchor="ctr", lnspc=126)
            cx += w
        yy += h
    return s, yy


# ====================================================== ④ 主採職等轉變
def slide_grades_main():
    _id[0] = 1
    s = header(EYE + "④", "主採｜AI 導入前後的職等轉變（五～二職等）", 2400)
    s += tb(MARGIN_X, 1.64, CONTENT_W, 0.26,
            [[run("共同方向：例行事務交給 AI，人往上做判斷 —— 每一階的重心都從「輸入」移到「核對、判斷與對外協商」",
                  1250, 1, NAVY)]])

    cols = [(1.15, "職等"), (2.35, "導入前主要在做"), (2.10, "導入後交給 AI"),
            (4.80, "導入後的重心"), (1.69, "要補強的能力")]
    rows = [
        [("5 職等 副理", 1, INK), "跟單、回覆群組訊息、親自處理異常細節、彙整報表",
                "資料彙整、報表、信件整理與擬稿",
                "看 AI 整理好的結果做判斷；用價格資料庫議價、執行供應商評核、靠品質平台事前預防異常。"
                "另訂出交期表、異常反映等標準格式讓 AI 能直接讀取，並驗收 AI 成效",
                ("數據判讀、流程設計、帶人", 0, BODY)],
        [("4 職等", 1, INK), "大量信件往來與單據處理",
                "郵件歸檔、擬信、PO、TA 等例行輸入",
                "管從下單到出貨的全流程；一般異常第一線判斷；比價與數據分析。"
                "也是 AI 工具開發主力，用 Claude Code 把組內需求做成工具",
                "分析能力、工具開發"],
        [("3 職等", 1, INK), "手動 key in PO、收料、TA、Ship monitor，以及文件歸檔",
                "上述輸入作業幾乎都交給 AI 或人機協作",
                "從輸入資料的人變成核對 AI 結果的人：檢查 AI 做得對不對、處理 AI 做不了的例外，"
                "並蒐集完整的異常資料（照片、缸號、數量）",
                "異常判斷、與供應商溝通，為升 4 職等做準備"],
        [("2 職等", 1, INK), "支援作業 —— 核色、測試、開單結帳、快遞寄送",
                "郵件和文件歸檔、表格更新提醒",
                "把核色、測試、快遞與文件作業標準化，確保交給 AI 和同事的資料完整，"
                "並初步檢查 AI 的輸出",
                "面料基礎知識與 AI 工具操作，為升 3 職等做準備"],
    ]
    tbl, bottom = table(MARGIN_X, 1.96, cols, rows, fill_to=6.02)
    s += tbl

    s += takeaway(bottom + 0.18, [
        run("→ ", 1300, 1, NAVY),
        run("AI 取代的是每一階的「輸入」，", 1300, 1, INK),
        run("空出來的位置由上一階的判斷工作往下填", 1300, 1, RED),
        run(" —— 這是升級路徑，不是減人路徑。", 1300, 1, INK),
    ])
    s += source("資料來源：主採 05 回饋（2026-11）。主採職等分六／五／四／三／二（中國籍）；六職等經理之角色定位見次頁。")
    return wrap(s)


# ====================================================== ⑤ 三單位職等角色
def slide_grades_all():
    _id[0] = 1
    s = header(EYE + "⑤", "職等角色定位與職能 — 主採・副採・駐廠", 2400)

    s += shape(MARGIN_X, 1.62, CONTENT_W, 0.50, TINT, "roundRect")
    s += shape(MARGIN_X, 1.62, 0.06, 0.50, NAVY)
    s += tb(MARGIN_X + 0.24, 1.68, CONTENT_W - 0.48, 0.38,
            [[run("主採分六／五／四／三／二（中國籍）；副採分五／三；駐廠分七／五／四與三 —— "
                  "職等愈低，被 AI 接手的輸入工作愈多；職等愈高，判斷、協商與跨部門整合的比重愈大。",
                  1100, 1, BODY)]], anchor="ctr", lnspc=124)

    cols = [(1.00, "單位"), (1.25, "職等"), (2.45, "角色定位"), (7.39, "主要職能")]
    R = lambda h, u, g, role, fn, col: [(u, 1, col), (g, 1, INK), role, fn]
    rows = [
        R(0.40, "主採", "6 職等 經理", "訂方向、做決策",
          "核心供應商佈局、年度量價、重大異常裁定、AI 導入方向、人員配置與考核", NAVY),
        R(0.40, "", "5 職等 副理", "對 account 的交期、成本、品質結果負責",
          "主導議價與 core item 降價、供應商評核、複雜異常對外協商、訂流程與格式標準、帶人", NAVY),
        R(0.40, "", "4 職等", "獨立完成指定 account",
          "下單到出貨全流程、一般異常第一線判斷、比價與數據分析、AI 工具開發主力", NAVY),
        R(0.40, "", "3 職等", "依標準流程執行",
          "PO 開立、收料、TA 與 Ship monitor（人機協作）、交期更新、樣布寄送、蒐集回報異常資料", NAVY),
        R(0.56, "", "2 職等", "完成支援作業，讓 3 職等以上專心本業",
          "缸差核色、測試、快遞寄送、缸差／文件歸檔、表格資料更新、異常布樣收發、"
          "China 經銷訂單主料開單與結帳、操作 AI 歸檔與提醒工具並初步檢查輸出", NAVY),
        R(0.40, "副採", "5 職等", "監督、戰略與成本管理",
          "除三職等全部職能外，另含監督協助組員工作異況、供應商關係與戰略開發、報價及成本管理、"
          "關注原物料價格波動及市場取向、異常防範、參與會議", BLUE),
        R(0.56, "", "3 職等", "獨立負責所屬部門副料全流程",
          "各自負責不同部門之副料相關全流程：款式品項／供應商／樣品／報價／測試／送核／"
          "採購計算與下單／交期管理／出貨安排／品質異常判斷處理／補料／各方溝通", BLUE),
        R(0.40, "駐廠業務", "7 職等", "策略整合面",
          "全產區流程前期風險討論與對策制定、異常狀況策略制定與跨部門整合、組員執行風險控管、"
          "專案項目訂定、對外客戶資訊整合", NAVY),
        R(0.40, "", "5 職等", "指導檢視面",
          "例行流程中的風險評估、異常狀況的部門間串聯、組員執行內容的檢視、專案項目主導", NAVY),
        R(0.40, "", "4 與 3 職等", "執行跟催面",
          "例行流程跟催檢查、異常狀況整理與回報、執行主管單位下達的解決方案", NAVY),
    ]
    tbl, _ = table(MARGIN_X, 2.26, cols, rows, head_h=0.38, fill_to=6.80)
    s += tbl
    s += source("資料來源：主採／副採／駐廠業務 05 回饋彙整（2026-11）。此為職能分工架構，人數編制另見第 ① 頁。")
    return wrap(s)


# ====================================================== ⑥ 跨單位作業反饋
def slide_crossunit():
    _id[0] = 1
    s = header(EYE + "⑥", "跨單位作業反饋 — 主採・副採・駐廠", 2400)

    units = [
        ("主採", NAVY, "僅就處外單位提出", [
            ("品管｜前端品管常無法提前發現大規模瑕疵",
             "異常到裁剪或成衣階段才發現 → 訂驗布抽驗比例與重點項目、異常建檔、高風險布種加嚴檢驗", 0),
            ("成衣廠・品管｜異常舉證不足或與實際不符",
             "供應商質疑判斷、影響信任 → 訂固定格式：款號、布號／缸號、數量、照片、樣布", 0),
            ("整合部｜PO 簽核後仍需人工、請款畫面無法操作、收料單需印紙本",
             "AI 省下的時間又花回去 → 簽核自動通知並產生寄廠草稿；開放請款畫面；改電子檔結帳", 0),
            ("（處內）異常窗口不清、交接資料不一致",
             "屬 6A／6B 處內問題，於處內解決", 1),
        ]),
        ("副採", BLUE, "核心卡點：太多系統要重覆輸入", [
            ("互動範圍",
             "除開發業務、大貨業務、各供應商、測試公司外，內部跨單位尚有成衣廠、樣品室、船務、"
             "會計、物流、材質研發課、二樓實驗室等", 0),
            ("卡點",
             "一般溝通沒有太大問題；惟公司內部系統或功能不少且有些不見得串聯，或使用習慣不同，"
             "仍會重工，不知如何收斂或共有資訊", 0),
            ("改善方向",
             "此正符合公司導入 AI 的方向 —— 現階段先用 AI 搬運這些資訊，卡點可消弭不少", 0),
        ]),
        ("駐廠業務", NAVY, "核心卡點：共有行事曆", [
            ("現況",
             "面料到布、缸差卡、乾濕摩擦測試、簽核狀態都記在駐廠業務個人 Excel", 0),
            ("無法連動",
             "工廠要另用 Teams／郵件回報、滬辦缸差卡無回傳確認、台北採購缺件要另外寫信、"
             "台北大貨業務看不到即時狀態只能來問駐廠", 0),
            ("影響與建議",
             "資訊靠人轉傳、版本不一致、通知延遲，異常往往到開裁前才發現 → "
             "轉為跨單位共用平台＋AI 自動比對與通知", 0),
        ]),
    ]

    gap = 0.24
    cw = (CONTENT_W - gap * 2) / 3
    iw = cw - 0.48
    top, bottom = 1.68, 6.04

    # measure every column first, so all three panels share one height
    need = 0.0
    for _, _, lead, rows in units:
        n = 0.14 + th("X", 12.5, iw, 118) + 0.08 + th(lead, 10.0, iw, 124) + 0.20
        for lab, txt, _g in rows:
            n += th(lab, 9.5, iw, 124) + 0.04 + th(txt, 9.0, iw, 126) + 0.16
        need = max(need, n + 0.12)
    panel_h = min(bottom - top, max(need, 3.0))

    for i, (name, col, lead, rows) in enumerate(units):
        x = MARGIN_X + i * (cw + gap)
        s += shape(x, top, cw, panel_h, PANEL, "roundRect")
        s += shape(x, top, 0.06, panel_h, col)
        y = top + 0.14
        hname = th(name, 12.5, iw, 118)
        s += tb(x + 0.24, y, iw, hname, [[run(name, 1250, 1, col)]], lnspc=118)
        y += hname + 0.08
        hlead = th(lead, 10.0, iw, 124)
        s += tb(x + 0.24, y, iw, hlead, [[run(lead, 1000, 1, col)]], lnspc=124)
        y += hlead + 0.20
        for lab, txt, grey in rows:
            lc = MUTED if grey else INK
            bc = MUTED if grey else BODY
            lh = th("· " + lab, 9.5, iw, 124)
            bh = th(txt, 9.0, iw, 126)
            s += tb(x + 0.24, y, iw, lh, [[run("· " + lab, 950, 1, lc)]], lnspc=124)
            s += tb(x + 0.24, y + lh + 0.04, iw, bh, [[run(txt, 900, 0, bc)]], lnspc=126)
            y += lh + bh + 0.16

    s += takeaway(6.24, [
        run("→ ", 1300, 1, NAVY),
        run("現階段先用 AI 搬運資訊、降低重覆輸入；中期需要的是", 1300, 1, INK),
        run("跨單位共用的狀態平台", 1300, 1, RED),
        run(" —— 讓工廠、滬辦、採購與大貨業務看到同一份即時狀態。", 1300, 1, INK),
    ])
    s += source("資料來源：主採／副採 06 回饋（2026-11）；駐廠業務部分取自「AI提案 by 駐廠業務 20261104」"
                "（該次 06 回饋未含駐廠）。")
    return wrap(s)


# ====================================================== package
NEW = [slide_impact, slide_grades_main, slide_grades_all, slide_crossunit]
BASE, UN, OUT = 'in_nov/base.pptx', 'out_un', 'nov_out.pptx'


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    os.chdir(here)
    if os.path.isdir(UN):
        shutil.rmtree(UN)
    zipfile.ZipFile(BASE).extractall(UN)

    existing = sorted(int(m) for f in os.listdir(os.path.join(UN, 'ppt/slides'))
                      for m in re.findall(r'slide(\d+)\.xml$', f))
    n0 = max(existing)
    names = []
    for i, fn in enumerate(NEW):
        name = 'slide%d.xml' % (n0 + 1 + i)
        names.append(name)
        open(os.path.join(UN, 'ppt/slides', name), 'w', encoding='utf-8').write(fn())
        open(os.path.join(UN, 'ppt/slides/_rels', name + '.rels'), 'w', encoding='utf-8').write(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')

    ctp = os.path.join(UN, '[Content_Types].xml')
    ct = open(ctp, encoding='utf-8').read().replace('</Types>', ''.join(
        '<Override PartName="/ppt/slides/%s" ContentType="application/vnd.openxmlformats-'
        'officedocument.presentationml.slide+xml"/>' % n for n in names) + '</Types>')
    open(ctp, 'w', encoding='utf-8').write(ct)

    prp = os.path.join(UN, 'ppt/_rels/presentation.xml.rels')
    pr = open(prp, encoding='utf-8').read()
    nxt = max(int(m) for m in re.findall(r'Id="rId(\d+)"', pr)) + 1
    rids, add = [], ''
    for n in names:
        rid = 'rId%d' % nxt; nxt += 1
        rids.append(rid)
        add += ('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/slide" Target="slides/%s"/>' % (rid, n))
    open(prp, 'w', encoding='utf-8').write(pr.replace('</Relationships>', add + '</Relationships>'))

    pxp = os.path.join(UN, 'ppt/presentation.xml')
    px = open(pxp, encoding='utf-8').read()
    lst = re.search(r'<p:sldIdLst>(.*?)</p:sldIdLst>', px, re.S).group(1)
    sid = max(int(m) for m in re.findall(r'id="(\d+)"', lst)) + 1
    extra = ''.join('<p:sldId id="%d" r:id="%s"/>' % (sid + i, r) for i, r in enumerate(rids))
    open(pxp, 'w', encoding='utf-8').write(px.replace(lst, lst + extra))

    if os.path.exists(OUT):
        os.remove(OUT)
    zf = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
    for root, _, files in os.walk(UN):
        for f in files:
            full = os.path.join(root, f)
            zf.write(full, os.path.relpath(full, UN))
    zf.close()
    print('wrote %s (%d + %d slides)' % (OUT, n0, len(names)))


main()
