# -*- coding: utf-8 -*-
"""Update the 副採 parts of the V6C deck from 副料採購 AI 運用規劃 Ver.2 (2026-11-11).

In-place corrections on pages ①②③ (副採 now has a same-basis weighted estimate,
and 省下工時 is separated from AI 運用 % because AI-handled work still needs
10–20% human review), plus two new pages ④⑤ carrying the half-year plan and the
12-process breakdown. Old ④⑤⑥ become ⑥⑦⑧.
"""
import os, re, shutil, sys, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gen2 import (emu, NAVY, BLUE, INK, BODY, MUTED, PANEL, TINT, RED, WHITE,
                  MARGIN_X, CONTENT_W, th, _id, run, tb, shape, header,
                  takeaway, source, wrap, table)
from arthur import prow, stats, lead, prender, pneed

SRC = 'v6c_un'
WORK = 'v6c_work'
OUT = 'v6c_out.pptx'

EYE = "V6 C　"
SRC_V2 = ("資料來源：副料採購 AI 運用現況→半年規劃 Ver.2（2026-11-11）。工時比重參考 Outlook 信件統計"
          "（2026-10-05～11-03：交期 32%、下單／調料 24%）；議價與品質多經電話與會議處理，已上調。"
          "各項比例皆為推估，待 2026-11～12 實測修正。")


# ============================================================ helpers
def band(y, rs, h=None, pad=0.10):
    """One-line tinted strip, for a note that does not deserve a whole panel."""
    txt = "".join(r["t"] for r in rs)
    sz = max(r["sz"] for r in rs) / 100.0
    hh = h or th(txt, sz, CONTENT_W - 0.48, 124) + pad * 2
    s = shape(MARGIN_X, y, CONTENT_W, hh, PANEL, "roundRect")
    s += shape(MARGIN_X, y, 0.06, hh, BLUE)
    s += tb(MARGIN_X + 0.24, y + pad, CONTENT_W - 0.48, hh - pad * 2, [rs], anchor="ctr", lnspc=124)
    return s, y + hh


def steps(y, items, gap=0.20):
    """Milestone cards in a row: bold heading over one or two lines of detail."""
    n = len(items)
    w = (CONTENT_W - gap * (n - 1)) / n
    iw = w - 0.40
    bh = max(th(b, 8.5, iw, 124) for _, b in items)
    h = 0.10 + 0.20 + 0.04 + bh + 0.12
    s = ""
    for i, (head, body) in enumerate(items):
        x = MARGIN_X + i * (w + gap)
        s += shape(x, y, w, h, TINT, "roundRect")
        s += tb(x + 0.20, y + 0.10, iw, 0.20, [[run(head, 1100, 1, NAVY)]], lnspc=112)
        s += tb(x + 0.20, y + 0.34, iw, bh, [[run(body, 850, 0, BODY)]], lnspc=124)
    return s, y + h


# ============================================================ ④ 半年規劃
def slide_plan():
    _id[0] = 1
    s = header(EYE + "④　副採", "AI 運用 15–20% → 35–40%：半年能再多省下多少工時", 2300)
    x, y = lead(1.66, "目標 2027 Q2（半年內）達成；其餘留給人力判斷 —— 議價、異常、供應商關係、跨部門協調。")
    s += x

    sx, y = stats(y + 0.14, [
        ("35–40%", "半年 AI 運用目標（現況 15–20%）"),
        ("28–36%", "半年省下工時（現況 12–18%）"),
        ("＋90 小時", "3 人每月再多省：78 → 168 小時"),
        ("＋1,080 小時", "比現況再多省的全年工時"),
    ])
    s += sx

    cols = [(2.05, "項目"), (2.35, "現況"), (3.35, "半年目標（2027 Q2）"), (4.34, "增加（中位數）")]
    rows = [
        [("AI 運用 %", 1, NAVY), "15–20%", ("35–40%", 1, BLUE), "＋約 20 個百分點"],
        [("省下工時 %", 1, NAVY), "12–18%", ("28–36%", 1, BLUE), "＋約 17 個百分點"],
        [("每人每月", 1, NAVY), "21–32 小時", ("49–63 小時", 1, BLUE), "＋約 30 小時"],
        [("3 人每月合計", 1, NAVY), "63–95 小時", ("148–190 小時", 1, BLUE), "＋約 90 小時"],
        [("時間用途", 1, NAVY), "日常作業＋微幅超時", "減少超時＋轉做成本・交期・品質", "回補超時，其餘轉做加值"],
    ]
    tx, y = table(MARGIN_X, y + 0.18, cols, rows, pad=0.10)

    s += tx
    bx, y = band(y + 0.12, [
        run("168 小時的用途　", 1000, 1, NAVY),
        run("① 先回補加班 33–66 小時，讓工時回歸常態　② 其餘約 100–135 小時轉做成本・交期・供應商管理，"
            "因應未來訂單款數擴大或支持新客戶", 1000, 0, BODY),
    ])
    s += bx

    lx, y = lead(y + 0.12, "何時達到 —— 依 12 項流程逐項推估（詳見第 ⑤ 頁）", 1050, NAVY)
    s += lx
    sx2, y = steps(y + 0.06, [
        ("2026-11～12　約 25%", "先把已有工具或資料基礎的流程做穩：出貨安排、採購計算與下單"),
        ("2027-01～03（Q1）　30–35%", "交期預警、催料信草稿、送核進度追蹤陸續上線"),
        ("2027-04～06（Q2）　35–40% 達成", "全流程覆蓋，並依 11～12 月實測結果修正"),
    ])
    s += sx2
    if y > 6.20:
        print('  ! ④ content reaches %.2f"' % y)

    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("AI 擴大的範圍是副料全流程；多省下的 90 小時／月", 1300, 1, INK),
        run("先回補超時，再轉做成本、交期與供應商管理", 1300, 1, NAVY),
        run(" —— 是承接更多款數的產能，不是可以騰出的編制。", 1300, 1, INK),
    ])
    s += source("推算基準：每人每月 176 小時（8 小時 × 22 天）× 3 人＝528 小時；AI 處理的工作仍需約 10–20% 時間人工複核，"
                "故省下工時 ≈ AI 運用 % × 0.8–0.9（中位數 × 0.85）；現行加班每人每日約 0.5–1 小時（3 人每月 33–66 小時）未計入基準。"
                "前提：ERP／Shipping Monitor 資料已可取得。資料來源：副料採購 AI 運用現況→半年規劃 Ver.2（2026-11-11）；"
                "各項比例皆為推估，待 2026-11～12 實測修正。")
    return wrap(s)


# ============================================================ ⑤ 12 項流程
def slide_processes():
    _id[0] = 1
    s = header(EYE + "⑤　副採", "全流程 12 項的 AI 可處理比例 — 加權後約 38%", 2300)
    x, y = lead(1.66, "優先開發的三項（採購計算與下單、交期管理、出貨安排）合計佔約 41% 工時，且已有工具或資料基礎 —— 以藍字標示。")
    s += x

    P = (1, BLUE)      # 優先開發
    N = (1, NAVY)
    cols = [(1.62, "流程"), (0.84, "工時比重"), (1.04, "半年可處理"), (0.84, "加權貢獻"),
            (3.96, "AI 做什麼"), (3.79, "人留什麼")]
    data = [
        ("款式品項", "8%", "40%", "3.2%", "整理 BOM 用料清單", "特殊工藝確認", 0),
        ("供應商", "5%", "20%", "1.0%", "評分表數據彙整", "選商、關係維護", 0),
        ("樣品", "8%", "40%", "3.2%", "調料信、寄樣格式（已有工具）", "急件判斷", 0),
        ("報價", "7%", "30%", "2.1%", "比價表、歷史價對照", "議價", 0),
        ("測試", "4%", "30%", "1.2%", "報告追蹤與摘要", "合格與否判斷", 0),
        ("送核", "6%", "40%", "2.4%", "進度追蹤、到期提醒", "核准判斷", 0),
        ("採購計算與下單", "15%", "50%", "7.5%", "分廠計算、PO 命名、ERP 資料檢核", "確認後下單", 1),
        ("交期管理", "18%", "40%", "7.2%", "交期預警、催料信草稿", "延誤處理、替代方案", 1),
        ("出貨安排", "8%", "70%", "5.6%", "船務 booking 自動補 key（已上線）", "例外處理", 1),
        ("品質異常判斷處理", "7%", "15%", "1.1%", "異常紀錄彙整與統計", "判斷與處理", 0),
        ("補料", "4%", "40%", "1.6%", "補料量計算、擬單", "責任歸屬確認", 0),
        ("各方溝通", "10%", "20%", "2.0%", "信件摘要、草稿", "協調與送出", 0),
    ]
    rows = []
    for name, w, cap, wt, ai, human, pri in data:
        c = BLUE if pri else NAVY
        rows.append([(name, 1, c), (w, 1 if pri else 0, c if pri else BODY),
                     (cap, 1 if pri else 0, c if pri else BODY),
                     (wt, 1, c), ai, human])
    rows.append([("合計", 1, NAVY), ("100%", 1, NAVY), "", ("約 38%", 1, NAVY),
                 ("→ 半年內合理上限 35–40%", 1, NAVY), ""])
    tx, ty = table(MARGIN_X, y + 0.18, cols, rows, pad=0.08)
    s += tx

    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("人的核心留在 ", 1300, 1, INK),
        run("議價、事前預防、品質異常判斷與處理、供應商關係、跨部門協調", 1300, 1, NAVY),
        run(" —— 副採由執行者轉型為成本・交期・品質的負責人。", 1300, 1, INK),
    ])
    s += source("「AI 可處理」＝AI 完成主要處理、人負責確認與送出；涉及客戶資料一律使用公司核准工具。" + SRC_V2)
    return wrap(s)


# ============================================================ in-place edits
TEXT_EDITS = {
    1: [
        # 副採 now has a same-basis weighted estimate, so the stat caption and the
        # 副採 panel footer stop saying it is pending.
        ('主採 AI 可協助比例（占比加權）；副採待評估',
         '主採 AI 可協助比例（占比加權）；副採現況 15–20%'),
        ('10/27–11/3 一週內交辦 12 類 41 件（對內 30、對外 11），Queenie、Nicky、Vivian 三人皆已使用。副採尚未做同口徑的加權估算。',
         '10/27–11/3 一週內交辦 12 類 41 件（對內 30、對外 11），Queenie、Nicky、Vivian 三人皆已使用。'
         '副採已完成同口徑推估：12 項流程加權後現況 15–20%，半年目標 35–40%（見第 ④⑤ 頁）。'),
        ('資料來源：面料採購 04–06 回饋（002 版，2026-11）；副採 AI 人機協作程度（2026-11-04）。'
         '12.6% 為依表列占比加權之推算值，非工時實測；主採占比取自信箱關鍵字分類（樣本僅一人信箱可見範圍、不含微信）。',
         '資料來源：面料採購 04–06 回饋（002 版，2026-11）；副採 AI 人機協作程度（2026-11-04）、'
         '副料採購 AI 運用現況→半年規劃 Ver.2（2026-11-11）。12.6% 與副採占比皆為依表列加權之推算值，非工時實測；'
         '主採占比取自信箱關鍵字分類（樣本僅一人信箱可見範圍、不含微信）。'),
    ],
    2: [
        # 副採's two largest categories are now measured: 交期管理 18%、採購計算與下單 15%
        ('主採最大的一類占 15%、副採最大的兩類各占 20% —— 沒有一塊大到值得整批自動化；而省下的 ',
         '主採最大的一類占 15%、副採最大的兩類占 18% 與 15% —— 沒有一塊大到值得整批自動化；而省下的 '),
        # joint figure: 主採 10–15%、副採 12–18%（AI 運用 % 扣掉人工複核後）
        ('15~20', '10~18'),
    ],
    3: [
        ('三個單位的工時釋出幅度都落在 15% 上下；釋出的時間轉入議價、預防與異常判斷',
         '三個單位的現況工時釋出都落在 15% 上下；副採另提出半年倍增規劃（見第 ④⑤ 頁）。釋出的時間轉入議價、預防與異常判斷'),
        ('15–20%', '12–18%'),
        ('約每月 26–35 小時（1 人 × 8 小時 × 22 天推估）',
         '現況可釋出工時（每人每月 21–32 小時）；半年目標 28–36%'),
        ('資料來源：主採／副採／駐廠業務 05 回饋彙整（2026-11）。工時釋出為各單位自行評估之推估值，非工時實測。',
         '資料來源：主採／副採／駐廠業務 05 回饋彙整（2026-11）、副料採購 AI 運用現況→半年規劃 Ver.2（2026-11-11）。'
         '副採數字已依 Ver.2 改為「省下工時」口徑（AI 運用 % 扣除 10–20% 人工複核）。工時釋出皆為推估值，非工時實測。'),
    ],
}

# slide 3: the 駐廠 column's 工作型態 body needs two lines; everything below drops 0.16"
GEOM_EDITS = {
    1: [
        # the 副採 footer gained a sentence and now runs to three lines
        ('<a:off x="6492240" y="4325874"/><a:ext cx="4910328" cy="288036"/>',
         '<a:off x="6492240" y="4325874"/><a:ext cx="4910328" cy="457200"/>'),
    ],
    3: [
        ('<a:off x="8302752" y="3203702"/><a:ext cx="3099816" cy="144018"/>',
         '<a:off x="8302752" y="3203702"/><a:ext cx="3099816" cy="288036"/>'),   # 工作型態 body
        ('<a:off x="8302752" y="3475736"/>', '<a:off x="8302752" y="3622040"/>'),  # 組織 label
        ('<a:off x="8302752" y="3661918"/>', '<a:off x="8302752" y="3808222"/>'),  # 組織 body
        ('<a:off x="8302752" y="3933952"/>', '<a:off x="8302752" y="4080256"/>'),  # AI 著力點 label
        ('<a:off x="8302752" y="4120134"/>', '<a:off x="8302752" y="4266438"/>'),  # AI 著力點 body
    ],
}

RENUMBER = {4: ('V6 C　④', 'V6 C　⑥'), 5: ('V6 C　⑤', 'V6 C　⑦'), 6: ('V6 C　⑥', 'V6 C　⑧')}

LAYOUT_REL = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
              'relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')


def xf(x, y, w, h):
    return '<a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/>' % (emu(x), emu(y), emu(w), emu(h))


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    os.chdir(here)
    if os.path.isdir(WORK):
        shutil.rmtree(WORK)
    shutil.copytree(SRC, WORK)

    for n, edits in TEXT_EDITS.items():
        p = os.path.join(WORK, 'ppt/slides/slide%d.xml' % n)
        t = open(p, encoding='utf-8').read()
        for a, b in edits:
            if a in t:
                t = t.replace(a, b)
            else:
                print('  ! slide%d text not found: %s' % (n, a[:40]))
        open(p, 'w', encoding='utf-8').write(t)

    for n, edits in GEOM_EDITS.items():
        p = os.path.join(WORK, 'ppt/slides/slide%d.xml' % n)
        t = open(p, encoding='utf-8').read()
        for a, b in edits:
            if a in t:
                t = t.replace(a, b, 1)
            else:
                print('  ! slide%d geom not found: %s' % (n, a[:46]))
        open(p, 'w', encoding='utf-8').write(t)

    for n, (a, b) in RENUMBER.items():
        p = os.path.join(WORK, 'ppt/slides/slide%d.xml' % n)
        t = open(p, encoding='utf-8').read()
        assert a in t, (n, a)
        open(p, 'w', encoding='utf-8').write(t.replace(a, b, 1))

    # the two new pages are written as slide7/slide8 and placed after slide3 by order
    new = [('slide7.xml', slide_plan), ('slide8.xml', slide_processes)]
    for name, fn in new:
        open(os.path.join(WORK, 'ppt/slides', name), 'w', encoding='utf-8').write(fn())
        open(os.path.join(WORK, 'ppt/slides/_rels', name + '.rels'), 'w',
             encoding='utf-8').write(LAYOUT_REL)

    ctp = os.path.join(WORK, '[Content_Types].xml')
    ct = open(ctp, encoding='utf-8').read().replace('</Types>', ''.join(
        '<Override PartName="/ppt/slides/%s" ContentType="application/vnd.openxmlformats-'
        'officedocument.presentationml.slide+xml"/>' % n for n, _ in new) + '</Types>')
    open(ctp, 'w', encoding='utf-8').write(ct)

    prp = os.path.join(WORK, 'ppt/_rels/presentation.xml.rels')
    pr = open(prp, encoding='utf-8').read()
    nxt = max(int(m) for m in re.findall(r'Id="rId(\d+)"', pr)) + 1
    rid_of, add = {}, ''
    for name, _ in new:
        rid = 'rId%d' % nxt; nxt += 1
        rid_of[name] = rid
        add += ('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/slide" Target="slides/%s"/>' % (rid, name))
    open(prp, 'w', encoding='utf-8').write(pr.replace('</Relationships>', add + '</Relationships>'))

    # existing slideN.xml -> rId, so the deck can be reordered 1,2,3,[new],4,5,6
    rid_for_file = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"', pr))
    rid_for_file = {v: k for k, v in rid_for_file.items()}

    pxp = os.path.join(WORK, 'ppt/presentation.xml')
    px = open(pxp, encoding='utf-8').read()
    lst = re.search(r'<p:sldIdLst>(.*?)</p:sldIdLst>', px, re.S).group(1)
    order = ['slide1.xml', 'slide2.xml', 'slide3.xml', 'slide7.xml', 'slide8.xml',
             'slide4.xml', 'slide5.xml', 'slide6.xml']
    rid_for_file.update(rid_of)
    body = ''.join('<p:sldId id="%d" r:id="%s"/>' % (256 + i, rid_for_file[f])
                   for i, f in enumerate(order))
    open(pxp, 'w', encoding='utf-8').write(
        px.replace('<p:sldIdLst>%s</p:sldIdLst>' % lst, '<p:sldIdLst>%s</p:sldIdLst>' % body))

    if os.path.exists(OUT):
        os.remove(OUT)
    zf = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
    for root, _, files in os.walk(WORK):
        for f in files:
            full = os.path.join(root, f)
            zf.write(full, os.path.relpath(full, WORK))
    zf.close()
    print('wrote %s (%d slides)' % (OUT, len(order)))


if __name__ == '__main__':
    main()
