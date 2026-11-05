# -*- coding: utf-8 -*-
"""04 OH 控管方向 — 人力與費用（來源：Joyce「人力及費用相關」郵件）.

Arthur 於季度會議指定此項由 Joyce 提出。兩頁，沿用 V6 C 版面。
"""
import os, re, shutil, zipfile, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gen2 import (tb, shape, run, header, takeaway, wrap, table, th, _id, emu,
                  MARGIN_X, CONTENT_W, NAVY, BLUE, INK, BODY, MUTED,
                  PANEL, TINT, RED, REDBG, WHITE, SRC_Y, SRC_H)

EYE = "V6 C ‧ 04 OH 控管方向　"


def src(text):
    return tb(MARGIN_X, SRC_Y, CONTENT_W, SRC_H, [[run(text, 850, 0, MUTED)]], lnspc=122)


# ============================================================ ① 2027 R0 預算總覽
def slide_budget():
    _id[0] = 1
    s = header(EYE + "①", "2027 R0 預算總覽 — OH-台 增加 546.1 萬，但業務可控的差旅交際費是減少的", 2200)

    stats = [("11,553.9 萬", "OH-台 2027 R0 預算（較 26 落點 +546.1 萬、5.0%）", RED),
             ("302.79 萬", "標打（26 落點 334.95 萬，−9.6%）", NAVY),
             ("44 人", "台北員額（26 落點 44.6 人）", NAVY),
             ("30.50", "匯率（26 落點 31.30）", NAVY)]
    gap = 0.24
    sw = (CONTENT_W - gap * 3) / 4
    for i, (num, lab, col) in enumerate(stats):
        x = MARGIN_X + i * (sw + gap)
        s += shape(x, 1.68, sw, 0.92, REDBG if col is RED else PANEL, "roundRect")
        s += shape(x, 1.68, 0.06, 0.92, col)
        s += tb(x + 0.24, 1.77, sw - 0.44, 0.38, [[run(num, 1900, 1, col)]])
        s += tb(x + 0.24, 2.17, sw - 0.44, 0.36, [[run(lab, 900, 0, BODY)]], lnspc=122)

    lw, rw = 7.50, 4.20
    rx = MARGIN_X + lw + 0.39

    s += tb(MARGIN_X, 2.70, lw, 0.26,
            [[run("費用別　2026 落點-09W3　vs　2027 預算-R0（元）", 1250, 1, NAVY)]])
    cols = [(1.86, "項目"), (1.74, "2026 落點"), (1.74, "2027 預算"), (1.42, "(增)減"), (0.74, "%")]
    # the source's convention: parentheses = 費用增加（紅）；無括號 = 費用減少（藍）
    data = [
        ("用人費用", "51,663,595", "54,049,204", "(2,385,610)", "-4.6%", 1),
        ("海外 OH", "41,665,019", "44,191,821", "(2,526,801)", "-6.1%", 1),
        ("差旅交際費", "2,449,200", "2,427,918", "21,283", "0.9%", 0),
        ("其他費用", "6,302,228", "6,726,715", "(424,487)", "-6.7%", 1),
        ("樣品費用", "4,186,454", "4,660,057", "(473,603)", "-11.3%", 1),
        ("折舊攤提費用", "3,811,424", "3,482,976", "328,448", "8.6%", 0),
        ("OH-台", "110,077,920", "115,538,690", "(5,460,770)", "-5.0%", 1),
    ]
    rows = []
    for name, a, b, d, pc, up in data:
        tot = (name == "OH-台")
        col = RED if up else BLUE
        rows.append([(name, 1, INK), (a, 0, BODY), (b, 1 if tot else 0, INK if tot else BODY),
                     (d, 1, col), (pc, 1, col)])
    tbl, bottom = table(MARGIN_X, 3.00, cols, rows, head_h=0.34, fill_to=5.92)
    s += tbl

    s += tb(MARGIN_X, bottom + 0.04, lw, 0.20,
            [[run("括號＝費用增加（紅）；無括號＝費用減少（藍），與系統畫面一致。", 850, 0, MUTED)]])

    s += tb(rx, 2.70, rw, 0.26, [[run("業務可控費用　拆解", 1250, 1, NAVY)]])
    s += shape(rx, 3.00, rw, 3.10, PANEL, "roundRect")
    s += shape(rx, 3.00, 0.06, 3.10, BLUE)
    y, iw = 3.14, rw - 0.48
    blocks = [
        ("差旅交際費　預算 242.8 萬", "較 26 落點減少 2.1 萬（0.9%）", BLUE,
         [("交際費 減少 17.3 萬", BLUE), ("差旅費 增加 13.0 萬", RED), ("交通費 增加 2.2 萬", RED)]),
        ("樣品費用　預算 466.0 萬", "較 26 落點增加 47.4 萬（11.3%）", RED,
         [("AD 銷樣／製樣 增加 30.6 萬", RED), ("其他費用 增加 24.0 萬", RED),
          ("樣品費收回 增加 11.7 萬", RED), ("快遞費 減少 20.7 萬", BLUE)]),
        ("專案費用", "新客戶開發專案-V6 預算 0", MUTED,
         [("去年未按 VT 拆分，無可比同期數", MUTED)]),
    ]
    for head, sub, hcol, items in blocks:
        hh = th(head, 10.0, iw, 122)
        s += tb(rx + 0.24, y, iw, hh, [[run(head, 1000, 1, hcol)]], lnspc=122)
        y += hh + 0.02
        sh_ = th(sub, 9.0, iw, 122)
        s += tb(rx + 0.24, y, iw, sh_, [[run(sub, 900, 0, BODY)]], lnspc=122)
        y += sh_ + 0.06
        for t, c in items:
            ih = th("· " + t, 9.0, iw - 0.18, 122)
            s += tb(rx + 0.42, y, iw - 0.18, ih, [[run("· " + t, 900, 0, c)]], lnspc=122)
            y += ih + 0.04
        y += 0.12

    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("OH-台 +5.0% 來自用人費用、海外 OH 與樣品費用；", 1300, 1, INK),
        run("業務可控的差旅交際費實際是 −0.9%", 1300, 1, RED),
        run("，而樣品費的增加屬開源投入。", 1300, 1, INK),
    ])
    s += src("資料來源：V6（業務六處）2027 預算-R0 對 2026 落點-09W3（含海外 OH 印尼／日本／柬埔寨／滬辦／北越）。"
             "六項費用加總等於 OH-台，已核對無誤。此處「台北員額 44 人」與次頁「編制 104→102 人」口徑不同：前者僅台北，後者為含海外之總編制。")
    return wrap(s)


# ============================================================ ② 人力與費用結構
def slide_overview():
    _id[0] = 1
    s = header(EYE + "②", "人力與費用控管方向 — 編制再減 2 人，費用用在開源與止損")

    stats = [("104 人", "2027 R0 原編制（含 6D 2 人）", NAVY),
             ("102 人", "R0 再減版調整後編制", NAVY),
             ("−2 人", "6A −1（取消台籍新增員額）、6C 印尼端 −1", RED),
             ("+13 萬", "2027 R0 差旅費較前期增加（NT$）", BLUE)]
    gap = 0.24
    sw = (CONTENT_W - gap * 3) / 4
    for i, (num, lab, col) in enumerate(stats):
        x = MARGIN_X + i * (sw + gap)
        s += shape(x, 1.68, sw, 0.98, REDBG if col is RED else PANEL, "roundRect")
        s += shape(x, 1.68, 0.06, 0.98, col)
        s += tb(x + 0.24, 1.77, sw - 0.44, 0.40, [[run(num, 2100, 1, col)]])
        s += tb(x + 0.24, 2.19, sw - 0.44, 0.38, [[run(lab, 900, 0, BODY)]], lnspc=122)

    lx, lw = MARGIN_X, 5.85
    rx, rw = MARGIN_X + 6.24, 5.85
    py, ph = 3.06, 2.74

    s += tb(lx, 2.74, lw, 0.26, [[run("人力｜2027 R0 編制調整", 1250, 1, NAVY)]])
    s += shape(lx, py, lw, ph, TINT, "roundRect")
    s += shape(lx, py, 0.06, ph, NAVY)
    y, iw = py + 0.16, lw - 0.48
    for lab, txt in [
        ("原編制", "2027 年 R0 編制為 104 人（含 6D 2 人）。"),
        ("再減版", "調整為 102 人。"),
        ("6A", "減少 1 人 —— 取消原規劃之台籍新增員額。"),
        ("6C", "印尼端減少 1 人。"),
        ("員額移轉", "印尼籍 Irfan 員額轉移至品管部。"),
    ]:
        h = max(th(lab, 9.5, 1.30, 124), th(txt, 9.0, iw - 1.44, 124))
        s += tb(lx + 0.24, y, 1.30, h, [[run("· " + lab, 950, 1, NAVY)]], anchor="ctr", lnspc=124)
        s += tb(lx + 0.24 + 1.44, y, iw - 1.44, h, [[run(txt, 900, 0, BODY)]], anchor="ctr", lnspc=124)
        y += h + 0.12

    s += tb(rx, 2.74, rw, 0.26, [[run("費用｜兩個方向", 1250, 1, BLUE)]])
    s += shape(rx, py, rw, ph, PANEL, "roundRect")
    s += shape(rx, py, 0.06, ph, BLUE)
    y, iw = py + 0.16, rw - 0.48
    for lab, txt, col in [
        ("1. 差旅交際費", "2027 R0 較前期增加約 NT$13 萬，來自兩項支援工作：台北特工赴印尼"
                       "控管水洗，以及駐廠／Local Assistant 攜樣赴 JKT 與日本取得客戶核可。"
                       "屬止損性支出 —— 見次頁成本效益。", BLUE),
        ("2. 樣品費", "樣品是開源之道。客人整體換設計師、銷售不佳，更應配合投入開發資源以增加商機；"
                    "建議不列為壓縮對象。", RED),
    ]:
        lh = th(lab, 10.0, iw, 124)
        bh = th(txt, 9.0, iw, 128)
        s += tb(rx + 0.24, y, iw, lh, [[run(lab, 1000, 1, col)]], lnspc=124)
        s += tb(rx + 0.24, y + lh + 0.06, iw, bh, [[run(txt, 900, 0, BODY)]], lnspc=128)
        y += lh + bh + 0.26

    s += takeaway(6.10, [
        run("→ ", 1300, 1, NAVY),
        run("人力已按 R0 再減版下修 2 人；費用面的增加集中在", 1300, 1, INK),
        run("「避免更大損失」與「開源」", 1300, 1, RED),
        run("兩類，非一般性成長。", 1300, 1, INK),
    ])
    s += src("資料來源：人力及費用相關（2026-11）。金額為預算與實際支出之彙整值，正式引用前請管理部複核。")
    return wrap(s)


# ============================================================ ② 差旅交際費
def slide_travel():
    _id[0] = 1
    s = header(EYE + "③", "差旅交際費 — 增加的兩個原因，以及為什麼仍屬划算", 2400)

    lx, lw = MARGIN_X, 5.85
    rx, rw = MARGIN_X + 6.24, 5.85
    py, ph = 1.70, 2.20

    s += shape(lx, py, lw, ph, PANEL, "roundRect")
    s += shape(lx, py, 0.06, ph, BLUE)
    s += tb(lx + 0.24, py + 0.14, lw - 0.48, 0.24,
            [[run("① 台北特工人員支援印尼專案", 1150, 1, BLUE)]])
    y, iw = py + 0.46, lw - 0.48
    for txt in [
        "小歐多次赴印尼出差，每次停留約 2～3 週，協助控管水洗相關作業。",
        "工作內容：水洗進度追蹤、品質確認，以及異常問題回報與改善對策執行。",
        "案件由去年的雪花洗專案延續至今年 #22 N020 酵素水洗訂單。",
    ]:
        h = th("· " + txt, 9.5, iw, 126)
        s += tb(lx + 0.24, y, iw, h, [[run("· " + txt, 950, 0, BODY)]], lnspc=126)
        y += h + 0.10
    s += shape(lx + 0.24, y + 0.06, iw, 0.46, WHITE, "roundRect")
    s += tb(lx + 0.40, y + 0.06, iw - 0.32, 0.46,
            [[run("相關費用　2025 年約 NT$14.9 萬　｜　2026 年 1–9 月約 NT$14.4 萬",
                  1000, 1, NAVY)]], anchor="ctr")

    s += shape(rx, py, rw, ph, PANEL, "roundRect")
    s += shape(rx, py, 0.06, ph, BLUE)
    s += tb(rx + 0.24, py + 0.14, rw - 0.48, 0.24,
            [[run("② 駐廠／Local Assistant 攜樣送核（H/C）", 1150, 1, BLUE)]])
    y, iw = py + 0.46, rw - 0.48
    for txt in [
        "協助將 H/C 樣品送至 JKT 或日本進行核可。",
        "赴 JKT：布料或品質異常樣需立即取得客戶確認，無法等待客戶下次訪廠。",
        "赴日本：APP 尚未完成核可，但為取得 PO 或滿足提前接線需求。",
    ]:
        h = th("· " + txt, 9.5, iw, 126)
        s += tb(rx + 0.24, y, iw, h, [[run("· " + txt, 950, 0, BODY)]], lnspc=126)
        y += h + 0.10
    s += shape(rx + 0.24, y + 0.06, iw, 0.46, WHITE, "roundRect")
    s += tb(rx + 0.40, y + 0.06, iw - 0.32, 0.46,
            [[run("人員快遞費用（2026）　JKT 1–6 月 NT$12.8 萬 ＋ JP NT$6.3 萬 ＝ NT$19.1 萬",
                  1000, 1, NAVY)]], anchor="ctr")

    s += tb(MARGIN_X, 4.02, CONTENT_W, 0.26,
            [[run("2026 年實際案例 — 全年 4 次，合計約 USD 7,564，平均每次約 USD 1,900",
                  1250, 1, NAVY)]])
    cols = [(1.70, "日期"), (2.10, "人員"), (2.10, "金額（USD）"), (6.19, "說明")]
    rows = [
        [("01/18", 1, INK), "Louis", ("1,693", 1, BLUE), "攜 H/C 樣品送核"],
        [("01/31", 1, INK), "Louis", ("1,850", 1, BLUE), "攜 H/C 樣品送核"],
        [("03/04", 1, INK), "Louis", ("2,044", 1, BLUE), "攜 H/C 樣品送核"],
        [("08/04", 1, INK), "Sylvia", ("1,977", 1, BLUE), "攜 H/C 樣品送核"],
    ]
    tbl, bottom = table(MARGIN_X, 4.30, cols, rows, head_h=0.34, fill_to=6.00)
    s += tbl

    band_y, band_h = bottom + 0.14, 0.72
    s += shape(MARGIN_X, band_y, CONTENT_W, band_h, REDBG, "roundRect")
    s += shape(MARGIN_X, band_y, 0.06, band_h, RED)
    s += tb(MARGIN_X + 0.24, band_y + 0.09, 2.05, 0.24, [[run("★ 成本效益", 1050, 1, RED)]])
    s += tb(MARGIN_X + 2.40, band_y + 0.07, CONTENT_W - 2.64, 0.58, [
        [run("單次 H/C 約 USD 1,800～1,900，但相較於延誤核可造成的產能損失、"
             "甚至後續成衣空運成本，仍屬必要且相對較低的支出。", 1000, 1, BODY)],
        [run("此外，人員頻繁出差造成的體力負荷、工時耗損與本職工作影響，"
             "未計入上述金額，亦應納入整體管理成本考量。", 950, 0, BODY)],
    ], lnspc=128)

    s += src("資料來源：人力及費用相關（2026-11）。2025 年人員快遞費用未獨立統計。"
             "USD 金額為單次實際支出，合計 7,564／4 次。")
    return wrap(s)


# ============================================================ package
SLIDES = [slide_budget, slide_overview, slide_travel]
BASE, UN, OUT = 'in_nov/base.pptx', 'oh_un', 'oh04_out.pptx'


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    os.chdir(here)
    if os.path.isdir(UN):
        shutil.rmtree(UN)
    zipfile.ZipFile(BASE).extractall(UN)
    sd = os.path.join(UN, 'ppt/slides')
    shutil.rmtree(sd)
    os.makedirs(os.path.join(sd, '_rels'))

    names = []
    for i, fn in enumerate(SLIDES, 1):
        n = 'slide%d.xml' % i
        names.append(n)
        open(os.path.join(sd, n), 'w', encoding='utf-8').write(fn())
        open(os.path.join(sd, '_rels', n + '.rels'), 'w', encoding='utf-8').write(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
            'relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')

    ctp = os.path.join(UN, '[Content_Types].xml')
    ct = re.sub(r'<Override PartName="/ppt/slides/[^"]*"[^>]*/>', '',
                open(ctp, encoding='utf-8').read())
    ct = ct.replace('</Types>', ''.join(
        '<Override PartName="/ppt/slides/%s" ContentType="application/vnd.openxmlformats-'
        'officedocument.presentationml.slide+xml"/>' % n for n in names) + '</Types>')
    open(ctp, 'w', encoding='utf-8').write(ct)

    prp = os.path.join(UN, 'ppt/_rels/presentation.xml.rels')
    pr = re.sub(r'<Relationship [^>]*Target="slides/[^"]*"[^>]*/>', '',
                open(prp, encoding='utf-8').read())
    used = {int(m) for m in re.findall(r'Id="rId(\d+)"', pr)}
    nxt = (max(used) if used else 0) + 1
    rids, add = [], ''
    for n in names:
        rid = 'rId%d' % nxt; nxt += 1
        rids.append(rid)
        add += ('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                'relationships/slide" Target="slides/%s"/>' % (rid, n))
    open(prp, 'w', encoding='utf-8').write(pr.replace('</Relationships>', add + '</Relationships>'))

    pxp = os.path.join(UN, 'ppt/presentation.xml')
    px = open(pxp, encoding='utf-8').read()
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
