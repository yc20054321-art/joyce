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


# ============================================================ ① 人力與費用結構
def slide_overview():
    _id[0] = 1
    s = header(EYE + "①", "人力與費用控管方向 — 編制再減 2 人，費用用在開源與止損")

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
    s = header(EYE + "②", "差旅交際費 — 增加的兩個原因，以及為什麼仍屬划算", 2400)

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
SLIDES = [slide_overview, slide_travel]
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
