# -*- coding: utf-8 -*-
"""Integrate the six unit decks for the 執行長與處長季度會議 into one deck for Arthur,
organised by the three functional lines: 開發 / 大貨 / 主副採.

Sources (all read from in_all/):
  Olga.pptx   日本線開發策略（01–03 / 執行機制 / 費用與獲利）
  Sofia.pptx  23／21 部門 2027 高標業績成長策略（01–04 / 06 / 07 / 結語）
  Bruce.pptx  童裝部 六大主題（01 / 02 / 03 / 05 / 06）
  A_Rene.pptx V6-A 大貨 05／06
  B_Peggy.pptx V6-B 大貨跟單 05／06
  C_Joyce.pptx V6-C 台北主採＋副採 05／06（含駐廠），04 OH 另見 oh04 deck
"""
import os, re, shutil, zipfile, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gen2 import (E, emu, NAVY, BLUE, INK, BODY, MUTED, PANEL, TINT, RED, REDBG, WHITE,
                  FONT, MARGIN_X, CONTENT_W, SRC_Y, SRC_H, th, nlines, _id, nid, esc,
                  run, tb, shape, header, takeaway, source, wrap, table)
from gen2 import _runs

EYE = "執行長與處長季度會議　"
SRC_ALL = ("資料來源：Olga 日本線開發策略、Sofia 23／21 部門 2027 高標業績成長策略、Bruce 童裝部六大主題（2026-11-04）、"
           "Rene V6-A 大貨 05／06、Peggy V6-B 大貨跟單 05／06、Joyce V6-C 台北主副採 04–06。")


# ---------------------------------------------------------------- panels
HANG = 0.17          # hanging indent for enumerated body lines


def tbx(x, y, w, h, paras, lnspc=126, marL=0.0, indent=0.0, anchor="t"):
    """Like gen2.tb but with a paragraph indent, so a wrapped enumerated item
    keeps its number hanging in the left margin instead of running on."""
    ps = "".join(
        '<a:p><a:pPr algn="l" marL="%s" indent="%s"><a:lnSpc><a:spcPct val="%d000"/></a:lnSpc>'
        '<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft>'
        '</a:pPr>%s</a:p>' % (emu(marL), emu(indent), lnspc, _runs(rs)) for rs in paras)
    return ('<p:sp><p:nvSpPr><p:cNvPr id="%d" name="TextBox %d"/><p:cNvSpPr txBox="1"/>'
            '<p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/>'
            '</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
            '<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="%s"/>'
            '<a:lstStyle/>%s</p:txBody></p:sp>'
            % (nid(), _id[0], emu(x), emu(y), emu(w), emu(h), anchor, ps))


def _lines(txt):
    """A body is either one string, or a list of items; each item is a string or
    a (head, rest) pair whose head is set in bold navy."""
    if isinstance(txt, (list, tuple)):
        return [it if isinstance(it, (list, tuple)) else (None, it) for it in txt]
    return [(None, txt)]


def _body_h(txt, bsz, iw):
    """Measured height; enumerated items wrap into iw - HANG."""
    items = _lines(txt)
    hang = HANG if len(items) > 1 else 0.0
    return sum(th((hd or "") + rest, bsz / 100.0, iw - hang, 126) for hd, rest in items)


def _body_xml(x, y, iw, h, txt, bsz):
    items = _lines(txt)
    if len(items) == 1 and items[0][0] is None:
        return tbx(x, y, iw, h, [[run(items[0][1], bsz, 0, BODY)]], lnspc=126)
    paras = []
    for hd, rest in items:
        rs = []
        if hd:
            rs.append(run(hd, bsz, 1, NAVY))
        rs.append(run(rest, bsz, 0, BODY))
        paras.append(rs)
    return tbx(x, y, iw, h, paras, lnspc=126, marL=HANG, indent=-HANG)


def pneed(w, name, lead, rows, name_sz=1250, lead_sz=1000, lsz=950, bsz=900, pad=0.16):
    iw = w - 0.48
    n = 0.14 + th(name, name_sz / 100.0, iw, 118) + 0.08
    if lead:
        n += th(lead, lead_sz / 100.0, iw, 124) + 0.20
    for lab, txt in rows:
        n += th("· " + lab, lsz / 100.0, iw, 124) + 0.04 + _body_h(txt, bsz, iw) + pad
    return n + 0.12


def prender(x, top, w, h, name, col, lead, rows, name_sz=1250, lead_sz=1000,
            lsz=950, bsz=900, pad=0.16):
    iw = w - 0.48
    s = shape(x, top, w, h, PANEL, "roundRect")
    s += shape(x, top, 0.06, h, col)
    y = top + 0.14
    hn = th(name, name_sz / 100.0, iw, 118)
    s += tb(x + 0.24, y, iw, hn, [[run(name, name_sz, 1, col)]], lnspc=118)
    y += hn + 0.08
    if lead:
        hl = th(lead, lead_sz / 100.0, iw, 124)
        s += tb(x + 0.24, y, iw, hl, [[run(lead, lead_sz, 1, col)]], lnspc=124)
        y += hl + 0.20
    for lab, txt in rows:
        lh = th("· " + lab, lsz / 100.0, iw, 124)
        bh = _body_h(txt, bsz, iw)
        s += tb(x + 0.24, y, iw, lh, [[run("· " + lab, lsz, 1, INK)]], lnspc=124)
        s += _body_xml(x + 0.24, y + lh + 0.04, iw, bh, txt, bsz)
        y += lh + bh + pad
    return s


def prow(items, top, bottom, gap=0.24, x0=MARGIN_X, total=CONTENT_W, tag="", **kw):
    """items: [(name, colour, lead, [(label, text), ...]), ...] laid out side by side,
    all sharing the tallest measured height."""
    n = len(items)
    w = (total - gap * (n - 1)) / n
    need = max(pneed(w, it[0], it[2], it[3], **kw) for it in items)
    if need > bottom - top:
        print("  ! %s panels need %.2f\" but only %.2f\" available" % (tag or "?", need, bottom - top))
        for it in items:
            print("      %-22s %.2f\"" % (it[0][:22], pneed(w, it[0], it[2], it[3], **kw)))
    h = min(bottom - top, max(need, 1.6))
    s = ""
    for i, (name, col, lead, rows) in enumerate(items):
        s += prender(x0 + i * (w + gap), top, w, h, name, col, lead, rows, **kw)
    return s, top + h


def stats(y, items):
    gap = 0.20
    n = len(items)
    w = (CONTENT_W - gap * (n - 1)) / n
    caph = max(th(c, 8.5, w - 0.40, 122) for _, c in items)
    h = 0.10 + 0.34 + 0.04 + caph + 0.12
    s = ""
    for i, (num, cap) in enumerate(items):
        x = MARGIN_X + i * (w + gap)
        s += shape(x, y, w, h, TINT, "roundRect")
        s += tb(x + 0.20, y + 0.10, w - 0.40, 0.34, [[run(num, 2000, 1, NAVY)]], lnspc=110)
        s += tb(x + 0.20, y + 0.48, w - 0.40, caph, [[run(cap, 850, 0, BODY)]], lnspc=122)
    return s, y + h


def lead(y, text, sz=1250, col=NAVY):
    h = th(text, sz / 100.0, CONTENT_W, 124)
    return tb(MARGIN_X, y, CONTENT_W, h, [[run(text, sz, 1, col)]], lnspc=124), y + h


# ============================================================== 封面
def s_cover():
    _id[0] = 1
    s = shape(0, 0, 0.16, 7.5, NAVY)
    s += shape(MARGIN_X, 1.30, 1.60, 0.06, BLUE)
    s += tb(MARGIN_X, 1.62, CONTENT_W, 0.34, [[run("執行長與處長季度會議　單位回饋整合", 1350, 1, BLUE)]])
    s += tb(MARGIN_X, 2.06, CONTENT_W, 1.22,
            [[run("V6 處｜2027 高標業績成長策略", 3600, 1, INK)],
             [run("與 AI 導入後的組織職能", 3600, 1, INK)]], lnspc=116)
    s += tb(MARGIN_X, 3.34, CONTENT_W, 0.30,
            [[run("六份單位回饋，按 開發／大貨／主副採 三條功能線歸整", 1250, 0, BODY)]])

    items = [
        ("開發", NAVY, "三個單位", [
            ("日本線　Olga", "01–03 策略、執行機制、費用與獲利"),
            ("23／33／21 部門　Sofia", "01–04、06 AI 與組織職能、07 卡點、結語"),
            ("童裝部　Bruce", "01、02、03、05、06（多項待確認）"),
        ]),
        ("大貨", BLUE, "兩個單位", [
            ("V6-A　Rene", "05 五流程階段工作分野與組織、06 兩則案例"),
            ("V6-B　Peggy", "05 跟單作業與 AI 可代理、組織、06 兩則案例"),
        ]),
        ("主副採", NAVY, "一個單位（含駐廠業務）", [
            ("V6-C　Joyce", "04 OH 控管方向、05 量體與職能轉型、06 跨單位卡點"),
        ]),
    ]
    s += prow(items, 3.86, 6.70, tag="cover")[0]
    s += tb(MARGIN_X, 6.56, CONTENT_W, 0.26,
            [[run("提報對象：Arthur 處長　｜　彙整日期：2026-11-06", 1000, 1, MUTED)]])
    s += source(SRC_ALL)
    return wrap(s)


# ============================================================== ① 總覽
def s_overview():
    _id[0] = 1
    s = header(EYE + "①", "三條功能線 × 六份回饋 — 各自交付了什麼", 2400)
    x, y = lead(1.66, "共通結論：三條線的 AI 導入都停在「人核後才送出」，而六份回饋指向同一組卡點 —— 資訊不在系統裡、責任歸屬未定、專業判斷未提前介入。")
    s += x

    cols = [(2.45, "功能線／單位"), (3.30, "業績與開發主軸（01–03）"),
            (3.30, "AI 導入與組織（05）"), (3.04, "跨單位卡點（06）")]
    rows = [
        [("開發　日本線 Olga・23／21 部 Sofia・童裝 Bruce", 1, NAVY),
         "四項可控對策＋KPI 兩層（結果季檢、過程月追）；23 部 29.6 萬標打／NCMT $25.57，"
         "21 部 Q1・Q3 各 1.7 萬標打／NCMT $27；童裝三個 program（水陸兩用短褲一季約 20 萬件，"
         "印尼需降至 NCMT $16–17）",
         "目前不改變職能分工與架構，以提升繁瑣日常作業效率為主；AI 先做摘要、檢索與提醒；"
         "新系統、人力與授權費用須有成效證據後再申請",
         "GU 無專屬 TD、台北技師未提前介入；IE 加成致 CM 偏高、日美線未區隔邏輯；"
         "樣品往返簡化作工，商品失去賣點"],
        [("大貨　V6-A Rene・V6-B Peggy", 1, NAVY),
         "本次回饋聚焦 05／06；業績與商機由開發端提出",
         "A 部：五流程階段重新分野為「AI 備料、人把關」，5 人（主管 1＋組員 4）；"
         "B 部：五類跟單作業，約可節省操作系統與整合資料時間的一半，6 人（主管 1＋資深 2＋組員 3）",
         "用料不佳時的證據保留（補布證據需達補布量 70% kg）、做工資訊反覆更新；"
         "工廠端作業疏失、用料專案責任歸屬"],
        [("主副採　V6-C Joyce（含駐廠業務）", 1, NAVY),
         "04 OH：OH-台 2027 R0 為 11,553.9 萬（較 26 落點 +546.1 萬／+5.0%，來自用人費用、"
         "海外 OH 與樣品費用），業務可控的差旅交際費 −0.9%；編制 104→102 人",
         "25 類工作、17 項已導入；主採 AI 可協助比例 12.6%（占比加權）；三單位自估工時釋出 15% 上下，"
         "員額與架構暫不變動",
         "主採 5 個卡點均為處外單位；副採太多系統要重覆輸入；駐廠缺跨單位共有行事曆"],
    ]
    tx, _ = table(MARGIN_X, y + 0.18, cols, rows, fill_to=6.14)
    s += tx
    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("三條線各自已把能自動化的輸入端做完；下一輪的成效不在工具數量，而在", 1300, 1, INK),
        run("證據與狀態的共用、責任單位的釐清、技術判斷的前移", 1300, 1, NAVY),
        run("。", 1300, 1, INK),
    ])
    s += source(SRC_ALL)
    return wrap(s)


# ============================================================== ② 開發・業績
def s_dev_target():
    _id[0] = 1
    s = header(EYE + "②　開發", "業績目標與增量來源 — 三個開發單位的 01", 2400)
    x, y = lead(1.66, "共同前提：客戶主導企劃、設計、預算與訂單分配；我們能承諾的是報價與交樣速度、樣品品質、材料與工藝方案，以及問題預警效率。")
    s += x

    items = [
        ("日本線　Olga", NAVY, "四項可控對策，各帶追蹤指標", [
            ("① 守住既有訂單　→ 交樣準時率｜可歸責流失件數", "前移品質、材料及交期風險管理"),
            ("② 擴大開發參與　→ 新邀請款數｜新增得單金額", "提前掌握季度計畫，爭取新增詢價與品類"),
            ("③ 提高報價競爭力　→ 報價準時率｜貢獻毛利", "開發初期整合客戶目標價與材料／工藝方案"),
            ("④ 降低無效費用　→ 重工費用｜處理工時", "減少內部重工與重複確認，AI 輔助追蹤"),
            ("KPI 分兩層", [
                ("業績結果　季檢　", "同期營收、訂單保有、新增得單、貢獻毛利與 NCMT／OH"),
                ("可執行指標　月追　", "新邀請款數、報價／交樣準時率、可歸責重工次數與費用"),
                ("下單率　", "列為結果指標，不列為單一部門承諾"),
            ]),
        ]),
        ("23／21 部門　Sofia", BLUE, "既有維持量體、新增建立成長來源", [
            ("23 部｜既有業務", "2026 標打數較 2025 減少約 20%、NCMT 減少 2%；2027 目標 29.6 萬標打"
                              "（持平 2026）、NCMT $25.57（較 2026 減少 2%）"),
            ("21 部｜新增業務", "Q1 1.7 萬標打（已進單）、Q3 1.7 萬標打（待開發）、NCMT 目標 $27"),
            ("新 category 的基礎", "21 部為 2027 新增業務，起因於與 22 部門款式共布；鈴木 Leader 認可 23 部"
                                "過往在北越工廠的生產經驗與品質"),
            ("定位", "23 部以維持既有量體、爭取較高 IE／NCMT 商機為主；21 部以新增開發 27FW 聯名款"
                   "建立成長來源"),
        ]),
        ("童裝部　Bruce", NAVY, "三個 program；最大挑戰是價格", [
            ("SS｜水陸兩用 Nylon 短褲", "長年款，一季約 20 萬件；印尼產區需降至 NCMT $16–17"),
            ("SS｜Nylon work jogger", "長年款；價格差 $1~$1.5（比較基準、幣別待確認）"),
            ("FW｜13 部門格子襯衫／上衣", "需與孟加拉競爭（競爭條件待確認）"),
            ("待確認", "各 program 的達成路徑與具體作法、多項量與 NCMT 未填、是否屬客性相對弱勢 program "
                     "尚未標示 —— 提報前需補齊"),
        ]),
    ]
    s += prow(items, y + 0.20, 6.18, tag="②", lsz=900, bsz=850, pad=0.14)[0]
    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("目標面已清楚：", 1300, 1, INK),
        run("日本線以可控指標承諾、23／21 部以量體與聯名款承諾、童裝的達成路徑仍待補齊", 1300, 1, NAVY),
        run("。", 1300, 1, INK),
    ])
    s += source("資料來源：Olga 日本線開發策略（結論摘要、經營前提、01）、Sofia 23／21 部門 01 與 04、Bruce 童裝部 01。"
                "童裝部多項數據於原簡報標註「待確認」，此處照實保留。")
    return wrap(s)


# ============================================================== ③ 開發・價與商機
def s_dev_price():
    _id[0] = 1
    s = header(EYE + "③　開發", "「價」怎麼上去、商機從哪裡來 — 三個開發單位的 02／03", 2300)
    x, y = lead(1.66, "共同邏輯：價不是在議價桌上談出來的，而是在開發初期把目標價、材料與工藝方案一起定下來 —— 三個單位各從不同切口做同一件事。")
    s += x

    items = [
        ("排序與定價　日本線", NAVY, "先排序 program，再談價", [
            ("量 × NCMT 四象限", [
                ("高毛利　", "高量 守住擴大｜低量 放大"),
                ("低毛利　", "高量 降成本或重新議價｜低量 退出或併款"),
            ]),
            ("「價」的四個層次（由易到難）", [
                ("① 成本前置　", "開發初期確認目標價，採購與技術同步提替代材料／工藝，價不動而 NCMT 上升"),
                ("② 條件面　", "MOQ 分級、急單加價、開發費收費"),
                ("③ 結構面　", "提高機能布、排汗、抗 UV、回收材質比重"),
                ("④ 價值面　", "主動提案，以準時率與不良率數據支持議價"),
            ]),
            ("議價原則", "附成本拆解（面料、工繳、匯率、法規成本），並同時提出可維持原價的替代方案"),
            ("Pipeline 12–18 個月", "鎖定名單 → 初次接觸 → 提案打樣 → 驗廠審核 → 試單 → 定番化；"
                                 "以月度開發看板管理階段與卡關原因"),
        ]),
        ("商機清單　23／21 部門", BLUE, "五項超越同期的具體商機", [
            ("款式價值提升", [
                ("① ", "爭取 IE 較高、NCMT 較高的聯名款"),
                ("② ", "把握追加單商機，搭配中國產區生產"),
                ("③ ", "爭取水洗特工款"),
            ]),
            ("訂單機會擴大", [
                ("④ ", "爭取期中企畫商機"),
                ("⑤ ", "爭取針織布做平織工的款式"),
            ]),
            ("守住量", "23 部維持 29.6 萬標打，掌握期中企畫及追加單；追加單評估轉中國產區的交期與價格"),
            ("Program 管理建議", "逐客戶／款式列管目標商機、開發進度、預估標打數、IE／NCMT、實際得單；"
                             "新增開發邀請與實際訂單分開追蹤"),
        ]),
        ("新品類與副料提案　童裝", NAVY, "以自採副料與水洗款開新局", [
            ("Pricing 策略", "以自採副料提案取代客指副料（如 YKK／IRIS／MORITO）"),
            ("27SS 期中企劃　彈力繩調節扣", "Step 1：自採陸盾取代客指 MORITO，替客人省約 USD 3K；"
                                      "Step 2：再爭取與供應商議價，提高 NCMT"),
            ("新 category｜水洗款式（12／13 部門）", "27FW 燈心絨 127F057（2 萬件）／127F062（5 萬件）／"
                                              "127F068（2 萬件）；預計 2026-11-20 前取得 EC 會後反饋"),
            ("跨部門延伸", "外套款 27SS 由 13 部門轉 11 部門，繼續延伸拓展 11 部門（進度待確認）"),
        ]),
    ]
    s += prow(items, y + 0.20, 6.18, tag="③", lsz=900, bsz=850, pad=0.14)[0]
    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("三個單位的共同要求是同一件事：", 1300, 1, INK),
        run("採購與技術必須在開發初期就進場", 1300, 1, NAVY),
        run(" —— 等到報價階段再找替代方案，價已經談不動了。", 1300, 1, INK),
    ])
    s += source("資料來源：Olga 日本線開發策略 02／03、Sofia 23／21 部門 02 與 03、Bruce 童裝部 02／03。"
                "童裝部「開發路徑／目前進度」多項於原簡報標註「待確認」。")
    return wrap(s)


# ============================================================== ④ 開發・組織
def s_dev_org():
    _id[0] = 1
    s = header(EYE + "④　開發", "AI 導入與組織職能 — 三個開發單位的 05", 2400)
    x, y = lead(1.66, "三個開發單位結論一致：不改變職能分工與架構；AI 先進到繁瑣的日常作業，新系統與人力費用須有成效證據後再申請。")
    s += x

    items = [
        ("日本線　Olga", NAVY, "執行機制：五方分工＋開發初期三項標準作業", [
            ("五方分工", [
                ("業　", "日本線業務｜客戶需求、目標價、優先順序與對外回覆"),
                ("採　", "採購｜材料規格、供應商價格與供料交期"),
                ("技　", "開發／技術／品質｜打樣、工藝可行性、測試與品質"),
                ("工　", "工廠／生產｜產能、製程效率、製造成本與量產風險"),
                ("AI　", "資訊｜資訊整合、待辦提醒、案例查詢與權限管理"),
            ]),
            ("三項標準作業", [
                ("① 需求確認表　", "業務彙整版次、規格、目標交樣日與客戶要求，重大變更標註影響範圍"),
                ("② 開發風險檢核　", "採購確認材料與交期、技術確認工藝與測試、工廠檢查可量產條件"),
                ("③ 異常結案紀錄　", "列負責人、承諾回覆日與處理結論，客戶確認後再關閉"),
            ]),
            ("資源配置與導入範圍", [
                ("資源　", "第一階段先運用現有人力；AI 先做摘要、檢索及提醒；新系統、人力與授權費用須有成效證據後再申請"),
                ("範圍　", "先從重點客戶及高風險款式開始，沿用既有表單與系統"),
            ]),
        ]),
        ("23／33／21 部門　Sofia", BLUE, "兩人共同負責，分工與架構不變", [
            ("業務副理　Sofia", "與 Kimi 共同負責 23／33／21 部門開發業務；另負責業務六處 PJT 週報整理，以及每週／每月樣品預計"
                           "報表統計與整理；23／33／21 DIV. 的 PJT 窗口與預算窗口；實際跟單款式較 Kimi 略少"),
            ("業務 MD　Kimi", "與 Sofia 共同分擔 23／33 部門所有款式開發相關業務；承擔較多實際開發跟單款式；"
                          "持續配合現有分工執行"),
            ("組織結論", "目前不改變職能分工與架構，以提升繁瑣日常作業效率為主；AI 已協助繁瑣作業並逐步加快進度"),
            ("效率目標（結語）", "同 account OH 低於同期；AI 支援報表與資訊整理；台北技師提前介入製樣溝通。"
                          "按季檢視標打量、NCMT 與實際得單成果"),
        ]),
        ("童裝部　Bruce", NAVY, "兩人共同改善，各自已有個人工具", [
            ("兩人共同改善", "自動化進單與翻譯；DL 雙週差異（更快找到變動款式）；快速整理與整合大量資訊與文件、"
                        "搜尋舊資訊以利快速回覆"),
            ("業務副理　Bruce", "快速統整資訊並翻譯成日文草稿；自動化做輪值報告；報價統整，檢查款式間價格正確性"),
            ("業務 MD　Andie", "信件優化；製作重複性工作用工具；規範檢視；報告製作"),
            ("待確認", "成效（節省時間／人力，以實測為準）與下一步計畫 —— 原簡報標註「待確認」"),
        ]),
    ]
    s += prow(items, y + 0.18, 6.18, tag="④", lsz=900, bsz=850, pad=0.14)[0]
    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("開發端沒有要求調整編制或架構 —— ", 1300, 1, INK),
        run("AI 的落點是報表、翻譯與資訊整理，和大貨、主副採「AI 備料、人把關」的結論一致", 1300, 1, NAVY),
        run("。", 1300, 1, INK),
    ])
    s += source("資料來源：Olga 日本線開發策略（執行機制、費用與獲利）、Sofia 23／21 部門 06 AI 導入與組織職能及結語、"
                "Bruce 童裝部 05（AI 導入進展與組織圖）。童裝部成效與下一步於原簡報標註「待確認」。")
    return wrap(s)


# ============================================================== ④ 大貨
def s_bulk():
    _id[0] = 1
    s = header(EYE + "⑤　大貨", "AI 導入後的工作分野 —「AI 備料、人把關」", 2400)
    x, y = lead(1.66,
        "A 部 5 人（主管 1＋組員 4）、B 部 6 人（主管 1＋資深 2＋組員 3），員額不變；兩部把資料搬運、系統輸入與固定文件交給 CC 與 AI，"
        "人留在對人的協調、決策與簽核 —— B 部自估約可節省操作系統、整合資料所需時間的一半，但仍需保留人員核對確認的時間。")
    s += x

    cols = [(0.86, "部別"), (2.42, "作業類別"), (1.12, "AI 介入"),
            (7.69, "人保留　／　AI 備料（人核後送出）")]
    rows = [
        [("V6-A", 1, NAVY), ("① 進單／系統", 1, INK), ("中 → 高", 1, BLUE),
         "人：口徑判斷、布規與布價檢視、核單與轉檔　／　AI：開樣品單、維護系統、制式郵件草稿（布規表、Trim 表、OE 進單）"],
        [("V6-A", 1, NAVY), ("② 樣品 APP／REAPP／Bulk", 1, INK), ("低 → 中", 1, BLUE),
         "人：寄出前確認、與客人與工廠溝通做工與交期　／　AI：樣品需求表、工廠信、資料包一次備齊（目前做不到，未來展望）"],
        [("V6-A", 1, NAVY), ("③ 發單／RFPO／PO", 1, INK), ("低 → 中", 1, BLUE),
         "人：內容審核、判斷並發送　／　AI：測試報告蒐集、PO 取得、拆 PO、加工單／IAN、下 ACC 等制式郵件初稿"],
        [("V6-A", 1, NAVY), ("④ 生產追蹤／異常", 1, INK), ("低 → 低", 1, MUTED),
         "人：與客人／工廠協調與決策、客訴責任判定與談判、生產異常處理方式　／　AI：查詢比對、時間軸、異損預警草稿"],
        [("V6-A", 1, NAVY), ("⑤ 管理／報表／文件", 1, INK), ("中 → 高", 1, BLUE),
         "人：訂艙／空運決策、單據簽核、付款、資料確認與上傳　／　AI：BOM／MP／一覽表、週月報、定倉提醒、自動化巡檢、KAKEN 對帳"],
        [("V6-B", 1, BLUE), ("訂單系統相關操作", 1, INK), ("中-高", 1, BLUE),
         "AI：進單核對、拆 PO、開長度副料單、估碼單 —— 可執行，大貨人員檢視後才能發出　／　人：TA 邏輯掌控與控管"],
        [("V6-B", 1, BLUE), ("採購／加工單發行", 1, INK), ("中", 1, BLUE),
         "AI：BOM PO、副料採購 —— 可計算並完成草稿，大貨人員檢視後才能發出　／　人：生產加工單內容複雜，目前自行處理"],
        [("V6-B", 1, BLUE), ("生產管理／工廠事務溝通", 1, INK), ("低", 1, MUTED),
         "AI：僅協助數字與資訊整理　／　人：與工廠互動溝通、TA 執行管理"],
        [("V6-B", 1, BLUE), ("異常事務處理", 1, INK), ("低", 1, MUTED),
         "AI：無　／　人：互動細節、理解狀況，開會議討論處理方向與分工"],
        [("V6-B", 1, BLUE), ("報表／客戶系統／日常事務", 1, INK), ("中", 1, BLUE),
         "AI：可產出報表　／　人：實際操作須人員確認後執行"],
    ]
    tx, _ = table(MARGIN_X, y + 0.18, cols, rows, pad=0.10, fill_to=6.18)
    s += tx
    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("兩部在「異常處理」都是最低介入：", 1300, 1, INK),
        run("AI 省得下來的是輸入與報表，省不下來的是與工廠、客人的來回", 1300, 1, NAVY),
        run("；B 部並指出目前仍在磨合階段，人員需另外投入時間訓練 AI。", 1300, 1, INK),
    ])
    s += source("資料來源：Rene V6-A 大貨 05（工作分野重新規劃、組織職能與 AI 協作角色）、Peggy V6-B 大貨跟單 05"
                "（跟單作業與 AI 可代理狀況、組織架構、AI 協助後可能的轉變）。「節省一半」為 B 部自行評估之推估值，非工時實測。")
    return wrap(s)


# ============================================================== ⑤ 主副採
def s_proc():
    _id[0] = 1
    s = header(EYE + "⑥　主副採", "量體、已導入的工具，以及為什麼人力不等比例減少", 2400)
    sx, y = stats(1.64, [
        ("25 類", "已盤點工作類型（主 14＋副 11）"),
        ("10 位", "人力編制（主 7＋副 3）27R0"),
        ("17 項", "已導入 AI 功能（主 12＋副 5）"),
        ("12.6%", "主採 AI 可協助比例（加權）"),
    ])
    s += sx

    items = [
        ("已導入與實績", NAVY, "主採 12 項（8 已上線・2 試行中・2 測試中）＋副採個人 Claude 5 項", [
            ("主採 12 項", [
                ("已上線 8　", "出貨 3 項、郵件 2 項、報告 3 項"),
                ("試行中 2　", "ERP 開 PO（TOMS）、收料"),
                ("測試中 2　", "BCI 更新提醒、PH3 交期表統整"),
            ]),
            ("主採實績", "ERP 開 PO 44 張（YAKJIN 22、NAMTEX 22）、收料 69 筆"),
            ("副採 5 項", "船務 booking 掃信 KEY 入並回信（不確定就攔停）、ACC 訂單回填、GU PO 命名歸檔、"
                       "分廠副料計算、擬信與報價比較"),
            ("副採實績", "10/27–11/3 一週交辦 12 類 41 件（對內 30、對外 11），三人皆已使用"),
        ]),
        ("職能轉型與工時釋出", BLUE, "主採 10–15%｜副採 15–20%（約每月 26–35 小時）｜駐廠 15%", [
            ("減少的工作", "單據核對、交期催詢、報表彙整、手動 key in 等重複性工作"),
            ("加重的工作", "數據分析、供應商合作與議價、前端品質控管、面料開發、異常判斷"),
            ("職等重心上移", "主採五～二職等的重心都從「輸入」移到「核對、判斷與對外協商」；"
                        "4 職等同時是 AI 工具開發主力，3 職等從輸入資料的人變成核對 AI 結果的人"),
            ("組織", "員額與組織架構暫不變動；釋出的時間轉入議價、預防與異常判斷，屬職能升級"),
        ]),
        ("人力為何不等比例減少", NAVY, "六個共通原因，加上三類不納入取代規劃", [
            ("六個原因", [
                ("① ", "與人對談是雙方議定，不是我司認定"),
                ("② ", "等外部回應的時間不因 AI 變快"),
                ("③ ", "要看實物才能拍板（色差、移染、白汙、看色核色、手感）"),
                ("④ ", "溝通載體在微信與電話，不在系統裡"),
                ("⑤ ", "承諾與責任要人扛"),
                ("⑥ ", "資料更動會回頭改寫上一輪答案"),
            ]),
            ("量體分布", "主採最大一類占 15%、副採最大兩類各占 20%，沒有一塊大到值得整批自動化；"
                      "省下的 15–20% 散在十個人每天的切片裡"),
            ("不納入取代規劃與員額口徑", [
                ("不納入　", "重大品質異常、議價協商（含匯率）、看色核色、品質預防、跨廠催辦、跨國交涉"),
                ("員額口徑　", "以款數／PO 張數／異常件數等量體指標判斷"),
            ]),
        ]),
    ]
    s += prow(items, y + 0.14, 6.18, tag="⑤", lsz=900, bsz=850, pad=0.14)[0]
    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("04 OH 的對照：OH-台 +5.0% 來自用人費用、海外 OH 與樣品費用，", 1300, 1, INK),
        run("業務可控的差旅交際費實際是 −0.9%", 1300, 1, NAVY),
        run("；編制 104→102 人。", 1300, 1, INK),
    ])
    s += source("資料來源：Joyce V6-C 台北主副採 04–06（面料採購 04–06 回饋 002 版、副採 AI 人機協作程度 2026-11-04、"
                "駐廠業務 AI 提案 2026-11-04）。12.6% 為依表列占比加權之推算值，15–20% 為各單位自估，皆非工時實測。")
    return wrap(s)


# ============================================================== ⑥ 06 卡點
def s_blockers():
    _id[0] = 1
    s = header(EYE + "⑦", "跨單位作業卡點（06）— 按功能線歸整", 2400)
    x, y = lead(1.66, "十一個卡點裡只有兩個是溝通問題；其餘都是「證據與狀態沒有留在系統裡」或「責任單位沒有定下來」。")
    s += x

    items = [
        ("開發", NAVY, "專業判斷沒有提前進場", [
            ("23／21 部｜GU 在公司內沒有專屬 TD", "平湖樣品室版師／技師直接與台北開發業務溝通，遇專業問題需轉信或電話，"
                                        "影響效率。27FW 起各色樣開發樣下單時 CC 台北技師，並希望延伸至工廠 APP 製樣前討論"),
            ("童裝｜IE 加成導致 CM 偏高", "IE 各種加成造成 CM 高騰、失去價格競爭力，商品轉去他社。IE 往往站在工廠角度；"
                                 "日美線需區隔邏輯 —— 美國線或可接受加價，日本線須逐項細究（布種加成 % 是否過多？"
                                 "加工段 IE 時間是否過多？）"),
            ("童裝｜樣品往返簡化作工", "需多次往返簡化作工，簡化設計的過程讓商品失去賣點，進而轉去他社"),
            ("日本線｜無效費用", "內部重工與重複確認、反覆打樣與材料浪費、內部疏失衍生的急件；"
                        "列為可執行指標按月追蹤重工費用與處理工時"),
        ]),
        ("大貨", BLUE, "證據與責任留不住", [
            ("V6-A｜用料不佳時的證據保留", "補布需工廠提供證據並達補布量 70%（kg），重量不足部分變成我方另付布款；"
                                  "短碼證據不足即無法求償。現況到結算才發現、再回頭找 → 生產開始到下線完整保留證據，"
                                  "並即時通知業務／主採"),
            ("V6-A｜做工資訊不清、來回反覆", "STL 60227H062A 腰內貼補強襯外露許容 0.4→0.5cm 反覆更新；成品數量 9/18→9/22→"
                                   "9/24→11/4 四度更新 → 工廠內部統一提供一份最終版許容，確認大貨實際可執行後才對客更新，"
                                   "並正確管理不同版本成品的流向"),
            ("V6-B｜工廠端作業疏失", "資訊已完整傳達，工廠仍有疏失造成異常須緊急處理（277N090A 通知腰卡用車線固定，"
                            "工廠改用膠針）"),
            ("V6-B｜用料專案責任歸屬", "應由工廠控管的實際用料，目前由業務回應說明 —— 是否應回歸責任單位？"
                              "AI 協助後反而內耗於細究數字（126F054C 主料品質異常，申請排除卻被要求詳細說明數字）"),
        ]),
        ("主副採", NAVY, "系統之間沒有串聯", [
            ("主採｜5 個卡點均為處外單位", [
                ("品管　", "前端常無法提前發現大規模瑕疵"),
                ("成衣廠・品管　", "異常舉證不足或與實際不符"),
                ("整合部　", "PO 簽核後仍需人工，AI 開單省下的時間又花回去"),
                ("整合部　", "請款畫面 AI 無法操作"),
                ("整合部　", "收料單需印紙本給越南會計"),
            ]),
            ("副採｜太多系統要重覆輸入", "一般溝通沒有太大問題；惟公司內部系統或功能不少且有些不見得串聯、使用習慣不同，"
                               "仍會重工，不知如何收斂或共有資訊"),
            ("駐廠｜缺共有行事曆", "面料到布、缸差卡、乾濕摩擦測試、簽核狀態都記在個人 Excel；工廠另用 Teams／郵件回報、"
                          "滬辦缸差卡無回傳確認、台北採購缺件要另外寫信、台北大貨業務看不到即時狀態只能來問駐廠 → "
                          "資訊靠人轉傳、版本不一致、通知延遲，異常往往到開裁前才發現"),
        ]),
    ]
    s += prow(items, y + 0.20, 6.18, tag="⑥", lsz=900, bsz=850, pad=0.14)[0]
    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("同一個病灶出現在三條線上：", 1300, 1, INK),
        run("關鍵狀態與證據記在個人的 Excel、郵件與微信裡", 1300, 1, NAVY),
        run(" —— 所以異常總是到結算或開裁前才浮出來。", 1300, 1, INK),
    ])
    s += source("資料來源：Sofia 07 跨單位作業卡點與改善、Bruce 童裝部 06、Olga 執行機制與費用獲利、Rene V6-A 06（兩則案例）、"
                "Peggy V6-B 06（兩則案例）、Joyce V6-C 06（主採／副採／駐廠）。")
    return wrap(s)


# ============================================================== ⑦ 建議
def s_recs():
    _id[0] = 1
    s = header(EYE + "⑧", "綜合判讀與建議 — 可列入季度追蹤的五件事", 2400)
    x, y = lead(1.66, "六份回饋的結論一致：AI 已經把輸入端做完了，下一輪的成效要從跨單位的制度上取得，而不是再多加幾個工具。")
    s += x

    items = [
        ("共通根因", NAVY, "四條，三條線重覆出現", [
            ("① 資訊在人身上，不在系統裡", "大貨的索賠證據與成品數量靠人回頭找；駐廠的到布、缸差、簽核狀態記在個人 Excel；"
                                 "副採在多套未串聯的系統間重覆輸入"),
            ("② 責任歸屬未定，爭議就回到業務身上", "工廠應控管的實際用料由業務回應；索賠證據由業務回頭補件；"
                                     "AI 協助後反而內耗於細究數字"),
            ("③ 專業判斷未提前介入", "GU 無專屬 TD、台北技師未在製樣前參與；許容反覆改到 0.5cm；"
                            "IE 加成未區分日／美線邏輯"),
            ("④ AI 接手的是輸入端", "主採加權 12.6%、三單位自估釋出 15% 上下、B 部自估約省一半系統與整合時間，"
                           "但全部附帶「人核後才送出」"),
        ]),
        ("建議", BLUE, "五件事，均可指定責任單位與完成時點", [
            ("① 證據與狀態移進共用平台", "跨單位共用的狀態／行事曆（到布、缸差、測試、簽核）；生產開始到下線即時留存證據"
                               "並通知業務與主採，不等結算才補件"),
            ("② 釐清責任單位", "實際用料由工廠控管並回覆；索賠證據由工廠於生產當下提供，格式固定為款號、"
                        "布號／缸號、數量、照片、樣布"),
            ("③ 技術判斷前移", "27FW 起各色樣開發樣下單時 CC 台北技師，並延伸至工廠 APP 製樣前討論；"
                        "IE 加成逐項細究並區隔日／美線邏輯"),
            ("④ AI 以「備料、人把關」為限", "重大品質異常、議價協商、看色核色、品質預防、跨國交涉不納入取代規劃；"
                                   "員額評估改以量體指標"),
            ("⑤ 指標分兩層", "業績結果季度檢視、可執行指標月度追蹤；新增開發邀請與實際得單分開統計"),
        ]),
        ("提報前待補齊", MUTED, "照實列出，不代為填補", [
            ("童裝部", "三個 program 的達成路徑、多項量與 NCMT、是否屬客性相對弱勢、新 category 的開發路徑與進度"
                   " —— 原簡報均標註「待確認」"),
            ("V6-A", "補布證據需達 70%（kg）之規則，原簡報註明「需相關單位確認」"),
            ("V6-B", "用料專案的責任歸屬需裁示；工廠端狀況目前業務較無法著力"),
            ("主副採", "12.6%（占比加權）與 15–20%（自估）皆非工時實測；主採占比取自信箱關鍵字分類，"
                   "樣本僅一人信箱可見範圍、不含微信"),
            ("口徑差異", "04 OH 的「台北員額 44 人」與「編制 104→102 人」口徑不同：前者僅台北，後者含海外"),
        ]),
    ]
    s += prow(items, y + 0.20, 6.18, tag="⑦", lsz=900, bsz=850, pad=0.14)[0]
    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("建議的五件事都不需要新增編制或新系統授權 —— ", 1300, 1, INK),
        run("要的是把證據、狀態與責任各歸其位", 1300, 1, NAVY),
        run("，AI 才接得下去。", 1300, 1, INK),
    ])
    s += source(SRC_ALL + "　本頁之歸納與建議為彙整者判讀，各單位原始主張見前述各頁。")
    return wrap(s)


# ============================================================== package
PAGES = [s_cover, s_overview, s_dev_target, s_dev_price, s_dev_org, s_bulk, s_proc, s_blockers, s_recs]
BASE, UN, OUT = 'in_nov/base.pptx', 'arthur_un', 'arthur_out.pptx'

LAYOUT_REL = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
              'relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    os.chdir(here)
    if os.path.isdir(UN):
        shutil.rmtree(UN)
    zipfile.ZipFile(BASE).extractall(UN)

    sd = os.path.join(UN, 'ppt/slides')
    for f in os.listdir(sd):
        if f.endswith('.xml'):
            os.remove(os.path.join(sd, f))
    for f in os.listdir(os.path.join(sd, '_rels')):
        os.remove(os.path.join(sd, '_rels', f))

    names = []
    for i, fn in enumerate(PAGES):
        name = 'slide%d.xml' % (i + 1)
        names.append(name)
        open(os.path.join(sd, name), 'w', encoding='utf-8').write(fn())
        open(os.path.join(sd, '_rels', name + '.rels'), 'w', encoding='utf-8').write(LAYOUT_REL)

    ctp = os.path.join(UN, '[Content_Types].xml')
    ct = open(ctp, encoding='utf-8').read()
    ct = re.sub(r'<Override PartName="/ppt/slides/slide\d+\.xml"[^>]*/>', '', ct)
    ct = ct.replace('</Types>', ''.join(
        '<Override PartName="/ppt/slides/%s" ContentType="application/vnd.openxmlformats-'
        'officedocument.presentationml.slide+xml"/>' % n for n in names) + '</Types>')
    open(ctp, 'w', encoding='utf-8').write(ct)

    prp = os.path.join(UN, 'ppt/_rels/presentation.xml.rels')
    pr = open(prp, encoding='utf-8').read()
    pr = re.sub(r'<Relationship Id="rId\d+"[^>]*Target="slides/slide\d+\.xml"/>', '', pr)
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
    new = ''.join('<p:sldId id="%d" r:id="%s"/>' % (256 + i, r) for i, r in enumerate(rids))
    px = px.replace('<p:sldIdLst>%s</p:sldIdLst>' % lst, '<p:sldIdLst>%s</p:sldIdLst>' % new)
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


if __name__ == "__main__":
    main()
