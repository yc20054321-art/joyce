# -*- coding: utf-8 -*-
"""Re-lay and tidy the user's reorganised 副採 deck (8b3597f5 / 2026-09-07 Q1).

What the user's revision changed, and what this script does about it:

* The 11-card 人力細項 page was moved in as slide 4, but still carries the eyebrow
  「四、副採」 from when it lived in the business deck. Re-labelled into section 二 and
  retitled so it no longer repeats slide 3's title.
* Slide 3's takeaway ended up carrying BOTH its own conclusion and slide 4's, in one
  paragraph. The text box overruns its band by 0.18". The borrowed sentence is removed
  (it still closes slide 4, where the eleven items actually are).
* A 12.13" x 4.50" SCREENSHOT of an earlier 工時特性 page was pasted onto slide 2,
  covering the whole system table and running 0.06" past the bottom of the slide.
  The bitmap is removed and its content rebuilt as a native page 7 — using slide 2's
  own authoritative figures, not the estimates the screenshot was made from.
* Slide 2's four stat cards are misaligned in the source (cards 1-2 overlap 0.17";
  gaps -0.17 / 0.37 / 0.10; card 2's label indented 0.37" vs 0.10"). Re-laid evenly.
* Cover subtitle "(人力編制-3位)" has no explicit size and uses theme colour bg2@75%,
  which is not guaranteed to read on the navy cover. Restyled to 27pt / C9DFF7.
  The line itself is kept.
* Slide 6's "見上一頁對照表" points three pages back, not one. Reworded.
"""
import os, re, shutil, zipfile
from gen import (tb, shape, run, header, takeaway, source, wrap, emu,
                 MARGIN_X, CONTENT_W, NAVY, BLUE, INK, BODY, PANEL, TINT, RED, REDBG, _id)

E = 914400.0
SRC, UN, OUT = 'q1.pptx', 'q1_unpacked', 'fucai_v2_out.pptx'


# ---------------------------------------------------------------- page fixes
def fix_stat_row(xml):
    pat = re.compile(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>')
    hits = [m for m in pat.finditer(xml)
            if 1.60 <= int(m.group(2)) / E <= 1.90 and 0.60 <= int(m.group(4)) / E <= 1.00]
    cards = [m for m in hits if int(m.group(4)) / E > 0.85]
    labels = [m for m in hits if int(m.group(4)) / E <= 0.85]
    if len(cards) != 4 or len(labels) != 4:
        print('  ! stat row unexpected (%d/%d) — left as is' % (len(cards), len(labels)))
        return xml
    cw = int(cards[0].group(3)) / E
    gap = (CONTENT_W - cw * 4) / 3
    edits = {}
    for i, m in enumerate(cards):
        edits[m.span()] = '<a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/>' % (
            emu(MARGIN_X + i * (cw + gap)), emu(1.70), m.group(3), m.group(4))
    for i, m in enumerate(labels):
        edits[m.span()] = '<a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/>' % (
            emu(MARGIN_X + i * (cw + gap) + 0.10), emu(1.80), m.group(3), m.group(4))
    for sp in sorted(edits, reverse=True):
        xml = xml[:sp[0]] + edits[sp] + xml[sp[1]:]
    print('  slide2: stat row re-laid (4 cards, even %.2f" gaps)' % gap)
    return xml


def drop_screenshot(xml):
    i = xml.find('<p:pic>')
    assert i != -1, 'expected the pasted screenshot on slide 2'
    j = xml.find('</p:pic>') + len('</p:pic>')
    print('  slide2: pasted screenshot removed (was 12.13" x 4.50", over the table)')
    return xml[:i] + xml[j:]


def fix_cover(xml):
    i = xml.find('人力編制')
    assert i != -1, '(人力編制-3位) missing from the cover'
    s = xml.rfind('<a:p>', 0, i); e = xml.find('</a:p>', i) + len('</a:p>')
    para = re.sub(
        r'<a:rPr([^>]*?)>\s*<a:solidFill>.*?</a:solidFill>',
        lambda m: '<a:rPr%s sz="2700" b="1"><a:solidFill><a:srgbClr val="C9DFF7"/></a:solidFill>'
                  % m.group(1).replace(' sz="2700"', '').replace(' b="1"', ''),
        xml[s:e], flags=re.S)
    xml = xml[:s] + para + xml[e:]
    print('  slide1: subtitle restyled 27pt/C9DFF7, 人力編制-3位 kept')

    # align the agenda wording with the section eyebrows used inside the deck
    xml = xml.replace('<a:t>二、人力為主業務範疇（11大類）</a:t>',
                      '<a:t>二、人力處理／判斷範疇（11大類）</a:t>')
    return xml


def fix_eyebrow_slash(xml):
    """Slide 3's eyebrow uses a half-width slash; the rest of the deck uses ／."""
    i = xml.find('<a:t>處理</a:t>')
    if i == -1:
        return xml
    j = xml.find('<a:t>/</a:t>', i)
    if j == -1 or j - i > 900:
        return xml
    print('  slide3: eyebrow slash normalised to ／')
    return xml[:j] + '<a:t>／</a:t>' + xml[j + len('<a:t>/</a:t>'):]


def fix_dup_takeaway(xml):
    """Slide 3's takeaway absorbed slide 4's closing line; drop the borrowed runs."""
    i = xml.find('系統只負責「登錄結果」')
    assert i != -1, "slide 3 no longer carries slide 4's line"
    end = xml.find('</a:r>', i) + len('</a:r>')
    start = xml.rfind('<a:r>', 0, i)
    # the stray "→ " run immediately before it goes too
    prev_start = xml.rfind('<a:r>', 0, start)
    prev = xml[prev_start:start]
    if '→' in prev:
        start = prev_start
    xml = xml[:start] + xml[end:]

    # the box was auto-fit to the doubled text: it now overruns its band by 0.18"
    old_xf = '<a:off x="786383" y="5495544"/><a:ext cx="10914549" cy="612648"/>'
    if old_xf in xml:
        xml = xml.replace(old_xf, '<a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/>'
                          % (emu(0.86), emu(5.98), emu(11.61), emu(0.49)))
        print('  slide3: takeaway box refitted to its band')
    print('  slide3: removed the duplicated closing line (belongs to slide 4)')
    return xml


def fix_moved_page(xml):
    xml = xml.replace('<a:t>四、副採</a:t>', '<a:t>二、人力處理／判斷範疇</a:t>')
    xml = xml.replace('<a:t>人力為主＋微 AI 協助為輔 — 十一項無法交給系統的協調作業</a:t>',
                      '<a:t>十一大類的實際內容 — 每一項都要「對人」談，不是把資料丟進系統</a:t>')
    xml = xml.replace('<a:t>資料來源：副採業務工作範疇彙整（2026-09-02 內部盤點）。</a:t>',
                      '<a:t>資料來源：副採業務工作範疇彙整（2026-09-02 內部盤點），對應本簡報'
                      '「人力為主的業務內容」頁十一大類之細項展開；文字為 AI 依工作描述歸納，'
                      '正式引用前請本人複核用詞。</a:t>')
    print('  slide4: eyebrow moved into section 二, title de-duplicated, source linked')
    return xml


def fix_page_ref(xml):
    assert '見上一頁對照表' in xml
    print('  slide6: page reference reworded')
    return xml.replace('見上一頁對照表', '見「系統操作」頁對照表')


# ------------------------------------------------- new native page (ex-screenshot)
def slide_hours_detail():
    _id[0] = 1
    s = header("四、工時特性", "可量化的部分 vs 不可量化的部分 — 為什麼「加總」會失真")

    s += shape(MARGIN_X, 1.70, CONTENT_W, 0.52, PANEL, "roundRect")
    s += shape(MARGIN_X, 1.70, 0.06, 0.52, BLUE)
    s += tb(MARGIN_X + 0.24, 1.78, CONTENT_W - 0.48, 0.38,
            [[run("每項工作耗時 10–30 分鐘 —— 但真正讓工時壓不下來的，是右欄那些不會出現在任何紀錄上的部分。",
                  1100, 1, BODY)]], anchor="ctr")

    lx, lw = MARGIN_X, 5.85
    rx, rw = MARGIN_X + 6.24, 5.85

    s += tb(lx, 2.40, lw, 0.26, [[run("① 可量化的部分 —— 系統操作", 1250, 1, NAVY)]])
    s += shape(lx, 2.72, lw, 3.34, PANEL, "roundRect")
    s += shape(lx, 2.72, 0.06, 3.34, BLUE)
    rows = [
        ("開立 PO（BY 款 BY 品項 per CBD）", "10–15 分／款"),
        ("列印 PO（BY PO 逐張）", "3–5 分／款"),
        ("提供檔案給供應商（每次 CBD 全流程）", "10–30 分／款"),
        ("SHIPPING MONITOR 每筆 shipment key-in", "1–3 分／筆"),
        ("實驗室品管／BV 送測 的申請書與報告管理", "視測試排程"),
        ("電子表單 異損申請", "視情況"),
    ]
    y = 2.88
    for a, b in rows:
        s += tb(lx + 0.26, y, lw - 2.05, 0.30, [[run("· " + a, 1000, 0, BODY)]], anchor="ctr")
        s += tb(lx + lw - 1.75, y, 1.50, 0.30, [[run(b, 1000, 1, BLUE)]], algn="r", anchor="ctr")
        y += 0.38
    s += tb(lx + 0.26, y + 0.06, lw - 0.52, 0.76,
            [[run("合計：一款進單，單就 ERP 三段作業即約 23–50 分（依上述區間相加）；"
                  "一款約開 3–7 張 PO，每張 PO 出貨 1–2 次，key-in 再各自累加。", 950, 1, RED)]],
            lnspc=128)

    s += tb(rx, 2.40, rw, 0.26, [[run("② 不可量化的部分 —— 為什麼「加總」會失真", 1250, 1, NAVY)]])
    s += shape(rx, 2.72, rw, 3.34, REDBG, "roundRect")
    s += shape(rx, 2.72, 0.06, 3.34, RED)
    notes = [
        "其間會有各項事務同時發生中 —— emails／Teams／電話／會議／供應商來訪，作業經常被中斷後再接回",
        "這些細節不像做報告那種「跑資料＋分析」可以一次做完，而是零碎時間的持續累積",
        "同一款的工時會隨 PO 張數、品項數、顏色數倍增 —— 款數不能代表 loading",
        "訂單切割、分廠重算、TA 變動時，前面已完成的計算與 PO 要整批重跑一次",
        "被中斷後重新接回原本的判斷脈絡，本身就是成本，且不會出現在任何工時紀錄上",
    ]
    y = 2.86
    for n in notes:
        s += tb(rx + 0.26, y, rw - 0.52, 0.56, [[run("· " + n, 1000, 0, BODY)]], lnspc=125)
        y += 0.62

    s += takeaway(6.30, [
        run("→ ", 1300, 1, NAVY),
        run("副採的工時不會集中出現在任何一張報表上 —— 它分散在每一款、每一張 PO、每一個顏色的重複操作與即時協調裡。",
            1300, 1, INK),
    ], h=0.58)
    s += source("資料來源：ERP／Shipping Monitor 系統實績（見「系統操作」頁）＋副採本人經驗訪談整理"
                "（2026-09-07）。合計為區間相加之推算值，非系統實際紀錄。")
    return wrap(s)


# ---------------------------------------------------------------- package work
def main():
    if os.path.isdir(UN):
        shutil.rmtree(UN)
    zipfile.ZipFile(SRC).extractall(UN)
    sl = lambda n: os.path.join(UN, 'ppt/slides/slide%d.xml' % n)

    for n, fns in [(1, [fix_cover]), (2, [drop_screenshot, fix_stat_row]),
                   (3, [fix_dup_takeaway, fix_eyebrow_slash]), (4, [fix_moved_page]), (6, [fix_page_ref])]:
        x = open(sl(n), encoding='utf-8').read()
        for fn in fns:
            x = fn(x)
        open(sl(n), 'w', encoding='utf-8').write(x)

    # the screenshot's relationship and media part are now unreferenced
    rp = os.path.join(UN, 'ppt/slides/_rels/slide2.xml.rels')
    r = open(rp, encoding='utf-8').read()
    r = re.sub(r'<Relationship [^>]*Target="\.\./media/image1\.png"[^>]*/>', '', r)
    open(rp, 'w', encoding='utf-8').write(r)
    img = os.path.join(UN, 'ppt/media/image1.png')
    if os.path.exists(img):
        os.remove(img)
        os.rmdir(os.path.dirname(img))
    ctp = os.path.join(UN, '[Content_Types].xml')
    ct = open(ctp, encoding='utf-8').read().replace(
        '<Default Extension="png" ContentType="image/png"/>', '')

    # append the rebuilt page as slide 7
    open(sl(7), 'w', encoding='utf-8').write(slide_hours_detail())
    open(os.path.join(UN, 'ppt/slides/_rels/slide7.xml.rels'), 'w', encoding='utf-8').write(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
        'relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
    ct = ct.replace('</Types>',
                    '<Override PartName="/ppt/slides/slide7.xml" ContentType="application/vnd.'
                    'openxmlformats-officedocument.presentationml.slide+xml"/></Types>')
    open(ctp, 'w', encoding='utf-8').write(ct)

    prp = os.path.join(UN, 'ppt/_rels/presentation.xml.rels')
    pr = open(prp, encoding='utf-8').read()
    rid = 'rId%d' % (max(int(m) for m in re.findall(r'Id="rId(\d+)"', pr)) + 1)
    pr = pr.replace('</Relationships>',
                    '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/'
                    '2006/relationships/slide" Target="slides/slide7.xml"/></Relationships>' % rid)
    open(prp, 'w', encoding='utf-8').write(pr)

    pxp = os.path.join(UN, 'ppt/presentation.xml')
    px = open(pxp, encoding='utf-8').read()
    lst = re.search(r'<p:sldIdLst>(.*?)</p:sldIdLst>', px, re.S).group(1)
    nid = max(int(m) for m in re.findall(r'id="(\d+)"', lst)) + 1
    px = px.replace(lst, lst + '<p:sldId id="%d" r:id="%s"/>' % (nid, rid))
    open(pxp, 'w', encoding='utf-8').write(px)
    print('  slide7: 可量化 vs 不可量化 rebuilt natively (was the pasted bitmap)')

    if os.path.exists(OUT):
        os.remove(OUT)
    zf = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
    for root, _, files in os.walk(UN):
        for f in files:
            full = os.path.join(root, f)
            zf.write(full, os.path.relpath(full, UN))
    zf.close()
    print('wrote', OUT)


main()
