# -*- coding: utf-8 -*-
"""Fold 主採 AI 推演（2026-11）into the V6C deck, mirroring the 副採 Ver.2 update.

Two new pages go in ahead of the 副採 pair (主採 is listed first throughout the
deck), so the 副採 pages move from ④⑤ to ⑥⑦ and everything after them shifts by
two. Cross-page references are rewritten to match.
"""
import os, re, shutil, sys, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gen2 import (NAVY, BLUE, INK, BODY, MUTED, PANEL, TINT, RED, WHITE,
                  MARGIN_X, CONTENT_W, th, _id, run, tb, shape, header,
                  takeaway, source, wrap, table)
from arthur import stats, lead
from fucai1111 import band, steps

SRC_DIR, WORK, OUT = 'v6c8_un', 'v6c10_work', 'v6c10_out.pptx'
EYE = "V6 C　"
SRC_Z = "資料來源：主採 AI 推演（2026-11）。"


# ============================================================ ④ 主採・比例推演
def slide_ratio():
    _id[0] = 1
    s = header(EYE + "④　主採", "AI 運用 10–15% → 2027 Q2 末 30%：結構上限 27–33%", 2300)
    x, y = lead(1.66, "依工時結構推算，主採 AI 可接手的上限約 27–33%；30% 可在 2027 Q2 末達成，33% 為上限，40% 超出結構上限。")
    s += x

    cols = [(5.55, "工作類別"), (1.70, "佔工時"), (2.32, "AI 可接手"), (2.52, "AI 運用")]
    rows = [
        [("系統輸入與單據（PO、收料、TA 等）", 1, NAVY), "20%", ("70–80%", 1, BLUE), ("14–16%", 1, BLUE)],
        [("資料彙整、報表、交期追蹤", 1, NAVY), "10%", ("60–70%", 1, BLUE), ("6–7%", 1, BLUE)],
        [("郵件與訊息溝通", 1, NAVY), "20%", "20–30%", "4–6%"],
        [("議價協商、對外交涉", 1, NAVY), "20%", ("5%", 1, RED), ("1%", 1, RED)],
        [("異常處理、看色核色、品質判斷", 1, NAVY), "20%", ("5%", 1, RED), ("1%", 1, RED)],
        [("等待外部回應、跨廠催辦", 1, NAVY), "10%", "10–20%", "1–2%"],
        [("合計（結構上限）", 1, NAVY), ("100%", 1, NAVY), "—", ("27–33%", 1, NAVY)],
    ]
    tx, y = table(MARGIN_X, y + 0.18, cols, rows, pad=0.09)
    s += tx

    bx, y = band(y + 0.14, [
        run("為什麼上不去　", 1000, 1, NAVY),
        run("議價協商、異常判斷、看色核色合計佔 40% 工時，AI 只能做資料準備（可接手約 5%）"
            " —— 這是 AI 運用比例難以超過 33% 的主因。", 1000, 0, BODY),
    ])
    s += bx

    sx, y = steps(y + 0.14, [
        ("2026 Q4　15–18%", "試行 4 項穩定上線；PO 擴大到新單與更多布廠"),
        ("2027 Q1　20–24%", "PO 簽核後自動記錄寄送、PO／Invoice 比對、交期擷取預警"),
        ("2027 Q2　27–30%", "價格資料庫、供應商評分、測試報告判讀、交接格式統一"),
        ("2027 Q3　30–33%", "請款畫面、收料電子檔、異常資料 AI 整理、餘布查詢"),
    ])
    s += sx
    if y > 6.20:
        print('  ! ④ 內容到 %.2f"' % y)

    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("30% 約在 2027 Q2 末（8 個月）達成、33% 為結構上限於 Q3 達到；", 1300, 1, INK),
        run("35–40% 需議價與異常類有新的 AI 突破，不列入本期承諾", 1300, 1, NAVY),
        run("。", 1300, 1, INK),
    ])
    s += source("推演基準：主採 7 人；工時結構為推估值（非工時實測），進度以 IT、財務、開發／大貨如期配合為前提。" + SRC_Z)
    return wrap(s)


# ============================================================ ⑤ 主採・節省工時
def slide_hours():
    _id[0] = 1
    s = header(EYE + "⑤　主採", "節省工時試算 — 達標後每月多省約 140 小時", 2300)
    x, y = lead(1.66, "AI 運用 12.5% → 31.5%（中間值）；扣除人工核對後（保留 30–40%），新增節省約 0.9 個人力、每年約 1,650 小時。")
    s += x

    sx, y = stats(y + 0.14, [
        ("70–120 小時", "現況每月節省（AI 運用 10–15%）"),
        ("200–260 小時", "達標後每月節省（AI 運用 30–33%）"),
        ("＋140 小時", "新增節省／月，約 0.9 人力（範圍 100–180）"),
        ("＋1,650 小時", "新增節省／年（範圍 1,200–2,150）"),
    ])
    s += sx

    cols = [(2.85, "時間"), (2.70, "AI 運用比例"), (3.30, "每月節省（扣人工核對後）"), (3.24, "比現況多省")]
    rows = [
        [("現況（2026/11）", 1, NAVY), "10–15%", "約 90 小時", ("—", 0, MUTED)],
        [("2026 Q4", 1, NAVY), "15–18%", "約 120 小時", ("＋30 小時", 1, BLUE)],
        [("2027 Q1", 1, NAVY), "20–24%", "約 160 小時", ("＋70 小時", 1, BLUE)],
        [("2027 Q2", 1, NAVY), "27–30%", "約 220 小時", ("＋130 小時", 1, BLUE)],
        [("2027 Q3 起", 1, NAVY), "30–33%", "約 240 小時", ("＋150 小時", 1, BLUE)],
    ]
    tx, y = table(MARGIN_X, y + 0.18, cols, rows, pad=0.09)
    s += tx

    for rs in [
        [run("計算方式　", 1000, 1, NAVY),
         run("主採 7 人 × 160 小時 ＝ 每月 1,120 小時；節省工時 ＝ AI 運用比例 × 1,120 × 65%"
             "（保留 35% 人工核對）。", 1000, 0, BODY)],
        [run("累計　", 1000, 1, NAVY),
         run("推進期間（2026/11–2027/9）約 930 小時；第一年（至 2027/10）約 1,080 小時。", 1000, 0, BODY)],
        [run("省下的時間用在　", 1000, 1, NAVY),
         run("① 吸收業務成長量，不另外加人　② 異常預防、議價、品質控管　"
             "③ 2、3 職等培訓，往判斷與溝通提升。", 1000, 0, BODY)],
    ]:
        bx, y = band(y + 0.12, rs)
        s += bx
    if y > 6.20:
        print('  ! ⑤ 內容到 %.2f"' % y)

    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("節省的時間散在每人每天的 10–30 分鐘，", 1300, 1, INK),
        run("撐住的是產能與判斷品質，而非可縮減的編制", 1300, 1, NAVY),
        run("。", 1300, 1, INK),
    ])
    s += source("數值為推估、非工時實測；建議選 1 個月做工時紀錄建立基準，之後每季比對。"
                "主採保留 30–40% 人工核對、副採保留 10–20%，兩單位口徑不同，跨單位加總前請先統一。" + SRC_Z)
    return wrap(s)


# ============================================================ 就地修改
TEXT_EDITS = {
    # ① 量體頁：主採補上推演目標，副採的頁次指引順移
    'slide1.xml': [
        ('主採 AI 可協助比例（占比加權）；副採現況 15–20%',
         '主採現況（占比加權）；結構上限 27–33%，2027 Q2 末目標 30%'),
        ('實績：ERP 開 PO 44 張（YAKJIN 22、NAMTEX 22）、收料 69 筆。十四類加權後的 AI 可協助比例為 12.6%（該簡報自述約 10–15%）。',
         '實績：ERP 開 PO 44 張（YAKJIN 22、NAMTEX 22）、收料 69 筆。十四類加權後的 AI 可協助比例為 12.6%（該簡報自述約 10–15%）。'
         '主採已完成工時結構推演：結構上限 27–33%，2027 Q2 末達 30%（見第 ④⑤ 頁）。'),
        ('半年目標 35–40%（見第 ④⑤ 頁）', '半年目標 35–40%（見第 ⑥⑦ 頁）'),
        ('副採 AI 人機協作程度（2026-11-04）、',
         '主採 AI 推演（2026-11）、副採 AI 人機協作程度（2026-11-04）、'),
        # 主採 footer 多了一句，需要第三行（副採那側上一輪已加高）
        ('<a:off x="786384" y="4325874"/><a:ext cx="4910328" cy="288036"/>',
         '<a:off x="786384" y="4325874"/><a:ext cx="4910328" cy="457200"/>'),
    ],
    # ③ 職能影響頁：主採欄改為 AI 運用比例口徑並補上目標
    'slide3.xml': [
        ('三個單位的現況工時釋出都落在 15% 上下；副採另提出半年倍增規劃（見第 ④⑤ 頁）。釋出的時間轉入議價、預防與異常判斷',
         '三個單位的現況釋出都落在 15% 上下；主採與副採均已提出推進規劃（主採見第 ④⑤ 頁、副採見第 ⑥⑦ 頁）。釋出的時間轉入議價、預防與異常判斷'),
        # 第一個「預估可釋出工時」為主採欄（文件順序：主採→副採→駐廠）
        ('<a:t>預估可釋出工時</a:t>',
         '<a:t>現況 AI 運用比例｜每月節省 70–120 小時；2027 Q2 末目標 30%</a:t>'),
        ('資料來源：主採／副採／駐廠業務 05 回饋彙整（2026-11）、副料採購 AI 運用現況→半年規劃 Ver.2（2026-11-11）。',
         '資料來源：主採／副採／駐廠業務 05 回饋彙整（2026-11）、主採 AI 推演（2026-11）、副料採購 AI 運用現況→半年規劃 Ver.2（2026-11-11）。'),
    ],
    # 副採半年規劃頁內的附件頁指引
    'slide7.xml': [('（詳見第 ⑤ 頁）', '（詳見第 ⑦ 頁）')],
}

RENUMBER = {'slide7.xml': ('V6 C　④　副採', 'V6 C　⑥　副採'),
            'slide8.xml': ('V6 C　⑤　副採', 'V6 C　⑦　副採'),
            'slide4.xml': ('V6 C　⑥', 'V6 C　⑧'),
            'slide5.xml': ('V6 C　⑦', 'V6 C　⑨'),
            'slide6.xml': ('V6 C　⑧', 'V6 C　⑩')}

NEW = [('slide9.xml', slide_ratio), ('slide10.xml', slide_hours)]
ORDER = ['slide1.xml', 'slide2.xml', 'slide3.xml', 'slide9.xml', 'slide10.xml',
         'slide7.xml', 'slide8.xml', 'slide4.xml', 'slide5.xml', 'slide6.xml']

LAYOUT_REL = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
              'relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    os.chdir(here)
    if os.path.isdir(WORK):
        shutil.rmtree(WORK)
    shutil.copytree(SRC_DIR, WORK)

    for name, edits in TEXT_EDITS.items():
        p = os.path.join(WORK, 'ppt/slides', name)
        t = open(p, encoding='utf-8').read()
        for a, b in edits:
            if a in t:
                t = t.replace(a, b, 1)
            else:
                print('  ! %s 找不到：%s' % (name, a[:44]))
        open(p, 'w', encoding='utf-8').write(t)

    for name, (a, b) in RENUMBER.items():
        p = os.path.join(WORK, 'ppt/slides', name)
        t = open(p, encoding='utf-8').read()
        assert a in t, (name, a)
        open(p, 'w', encoding='utf-8').write(t.replace(a, b, 1))

    for name, fn in NEW:
        open(os.path.join(WORK, 'ppt/slides', name), 'w', encoding='utf-8').write(fn())
        open(os.path.join(WORK, 'ppt/slides/_rels', name + '.rels'), 'w',
             encoding='utf-8').write(LAYOUT_REL)

    ctp = os.path.join(WORK, '[Content_Types].xml')
    ct = open(ctp, encoding='utf-8').read().replace('</Types>', ''.join(
        '<Override PartName="/ppt/slides/%s" ContentType="application/vnd.openxmlformats-'
        'officedocument.presentationml.slide+xml"/>' % n for n, _ in NEW) + '</Types>')
    open(ctp, 'w', encoding='utf-8').write(ct)

    prp = os.path.join(WORK, 'ppt/_rels/presentation.xml.rels')
    pr = open(prp, encoding='utf-8').read()
    rid = {f: r for r, f in re.findall(r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"', pr)}
    nxt = max(int(m) for m in re.findall(r'Id="rId(\d+)"', pr)) + 1
    add = ''
    for name, _ in NEW:
        r = 'rId%d' % nxt; nxt += 1
        rid[name] = r
        add += ('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/slide" Target="slides/%s"/>' % (r, name))
    open(prp, 'w', encoding='utf-8').write(pr.replace('</Relationships>', add + '</Relationships>'))

    pxp = os.path.join(WORK, 'ppt/presentation.xml')
    px = open(pxp, encoding='utf-8').read()
    lst = re.search(r'<p:sldIdLst>(.*?)</p:sldIdLst>', px, re.S).group(1)
    body = ''.join('<p:sldId id="%d" r:id="%s"/>' % (256 + i, rid[f]) for i, f in enumerate(ORDER))
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
    print('wrote %s (%d slides)' % (OUT, len(ORDER)))


if __name__ == '__main__':
    main()
