# -*- coding: utf-8 -*-
"""Re-lay and tidy the user's reorganised 主採 deck (143d6bde / 2026-09-07 G).

The revision mixes two page families that never matched:

  * pages 2 / 4 / 6 came from the business deck — 微軟正黑體, navy 1C4F8F, a 0.16"
    left rail, content column x=0.62 w=12.09, divider at 1.46, source at 7.18 — but
    still carry the eyebrow 「三、主採」 from that deck's numbering.
  * pages 3 / 5 / 7-11 are new — Calibri (no CJK glyphs, so PowerPoint substitutes),
    navy 1E2761, no rail, no divider, column x=0.50 w=12.33, source at 7.05.

This script normalises the second family onto the first, re-orders the deck into its
own four declared sections, and sharpens the framing the user asked for: AI assists,
it does not replace people.
"""
import os, re, shutil, zipfile
import defusedxml.ElementTree as ET
from gen import (tb, shape, run, takeaway, emu, MARGIN_X, CONTENT_W,
                 NAVY, BLUE, INK, BODY, MUTED, RED, _id)

E = 914400.0
SRC, UN, OUT = 'g1.pptx', 'g1_unpacked', 'zhucai_v2_out.pptx'
S = CONTENT_W / 12.33          # their column -> the family column
DY = 0.14                      # body drop, to clear the divider at 1.46

THEIRS = [3, 5, 7, 8, 9, 10, 11]

COLOURS = {                    # their palette -> the deck family's
    '1E2761': NAVY, '5A6B8C': MUTED, 'C0392B': RED,
    'CADCFC': 'E8F1FC', 'F7F9FC': 'FFFFFF',
}


# ------------------------------------------------------------------ primitives
PAIR = re.compile(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>')
BODY_TOP, BODY_BOT = 1.68, 7.02        # the family's content band
TAKEAWAY_BOT = 6.20                    # when a closing band has to fit underneath


def body_extent(xml):
    """Vertical span of the page body, ignoring header (<1.20) and source (>=7.00)."""
    lo, hi = 99.0, 0.0
    for m in PAIR.finditer(xml):
        y, h = int(m.group(2)) / E, int(m.group(4)) / E
        if y < 1.20 or y >= 7.00:
            continue
        lo, hi = min(lo, y), max(hi, y + h)
    return (lo, hi) if hi else (None, None)


def remap_geometry(xml, bottom):
    """Map x onto the family column, and the body band onto [BODY_TOP, bottom].

    Pages that already fit are shifted; the two that do not (their content is taller
    than the band) are compressed, heights and table rows included."""
    lo, hi = body_extent(xml)
    if lo is None:
        return xml
    src_h = hi - lo
    dst_h = bottom - BODY_TOP
    vs = min(1.0, dst_h / src_h)          # only ever compress, never stretch
    if vs < 0.999:
        print('     body %.2f" -> %.2f" (%.1f%% compression)' % (src_h, dst_h, (1 - vs) * 100))

    def fix(m):
        x0, y0 = int(m.group(1)) / E, int(m.group(2)) / E
        w, h = int(m.group(3)) / E, int(m.group(4)) / E
        nx, nw = MARGIN_X + (x0 - 0.50) * S, w * S
        if 1.20 <= y0 < 7.00:
            ny, nh = BODY_TOP + (y0 - lo) * vs, h * vs
        else:
            ny, nh = y0, h
        return '<a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/>' % (
            emu(nx), emu(ny), emu(nw), emu(nh))

    xml = PAIR.sub(fix, xml)
    # <p:spTree>'s own degenerate xfrm must stay at the origin
    xml = re.sub(r'(<p:grpSpPr><a:xfrm>)<a:off x="-?\d+" y="-?\d+"/><a:ext cx="0" cy="0"/>',
                 r'\1<a:off x="0" y="0"/><a:ext cx="0" cy="0"/>', xml)
    xml = re.sub(r'<a:gridCol w="(\d+)"',
                 lambda m: '<a:gridCol w="%s"' % emu(int(m.group(1)) / E * S), xml)
    if vs < 0.999:
        xml = re.sub(r'<a:tr h="(\d+)"',
                     lambda m: '<a:tr h="%s"' % emu(int(m.group(1)) / E * vs), xml)
    return xml


def restyle(xml):
    for a, b in COLOURS.items():
        xml = xml.replace('val="%s"' % a, 'val="%s"' % b)
    xml = re.sub(r'typeface="Calibri"', 'typeface="微軟正黑體"', xml)
    return xml


def _shape_span(xml, y_lo, y_hi):
    """Span of the first <p:sp> whose offset y falls in [y_lo, y_hi)."""
    for m in re.finditer(r'<p:sp>', xml):
        s = m.start(); e = xml.find('</p:sp>', s) + len('</p:sp>')
        om = re.search(r'<a:off x="(-?\d+)" y="(-?\d+)"/>', xml[s:e])
        if not om:
            continue
        y = int(om.group(2)) / E
        if y_lo <= y < y_hi:
            return s, e
    return None


def place(xml, span, x, y, w, h):
    s, e = span
    seg = re.sub(r'<a:off x="-?\d+" y="-?\d+"/><a:ext cx="\d+" cy="\d+"/>',
                 '<a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/>'
                 % (emu(x), emu(y), emu(w), emu(h)), xml[s:e], count=1)
    return xml[:s] + seg + xml[e:], (s, s + len(seg))


def set_runs(xml, span, sz=None, colour=None):
    s, e = span
    seg = xml[s:e]
    if sz is not None:
        seg = re.sub(r'(<a:rPr[^>]*?)\ssz="\d+"', r'\1', seg)
        seg = re.sub(r'<a:rPr\b', '<a:rPr sz="%d"' % sz, seg)
    if colour is not None:
        seg = re.sub(r'<a:srgbClr val="[0-9A-F]{6}"/>',
                     '<a:srgbClr val="%s"/>' % colour, seg)
    return xml[:s] + seg + xml[e:]


def add_chrome(xml):
    """Left rail + header divider, matching every other page in the family."""
    _id[0] = 900
    chrome = shape(0, 0, 0.16, 7.5, NAVY) + shape(MARGIN_X, 1.46, CONTENT_W, 0.03, NAVY)
    i = xml.find('</p:grpSpPr>') + len('</p:grpSpPr>')
    return xml[:i] + chrome + xml[i:]


def add_takeaway(xml, runs, y=6.36, h=0.58):
    _id[0] = 950
    i = xml.find('</p:spTree>')
    return xml[:i] + takeaway(y, runs, h=h) + xml[i:]


def normalise(xml, bottom=BODY_BOT):
    xml = restyle(xml)
    xml = remap_geometry(xml, bottom)

    eb = _shape_span(xml, 0.30, 0.60)
    if eb:
        xml, eb = place(xml, eb, MARGIN_X, 0.42, CONTENT_W, 0.32)
        xml = set_runs(xml, eb, sz=1250, colour=BLUE)
    ti = _shape_span(xml, 0.70, 1.00)
    if ti:
        xml, ti = place(xml, ti, MARGIN_X, 0.76, CONTENT_W, 0.46)
        xml = set_runs(xml, ti, colour=INK)
    sr = _shape_span(xml, 7.00, 7.40)
    if sr:
        xml, sr = place(xml, sr, MARGIN_X, 7.18, CONTENT_W, 0.26)
        xml = set_runs(xml, sr, sz=850, colour=MUTED)
    return add_chrome(xml)


def sub(xml, pairs, where=''):
    for a, b in pairs:
        if a not in xml:
            print('  ! not found on %s: %s' % (where, a[:44]))
            continue
        xml = xml.replace(a, b)
    return xml


# ------------------------------------------------------------------- messaging
SEC1, SEC2, SEC3, SEC4 = '一、主採範疇', '二、AI 能協助到哪裡', '三、已導入的 AI 工具', '四、結論與建議'

# final order -> (source slide number, eyebrow)
ORDER = [
    (1,  None),
    (2,  SEC1), (3, SEC1), (5, SEC1), (7, SEC1), (8, SEC1),
    (4,  SEC2), (10, SEC2),
    (6,  SEC3), (9, SEC3),
    (11, SEC4),
]

EDITS = {
    # ---- cover: name section 二 for what it actually argues
    1: [('二、AI 可代勞程度分析', '二、AI 能協助到哪裡、不能代替什麼'),
        # the 40pt title sits in a 1.0" box and overlaps the subtitle at 2.55"
        ('<a:off x="731520" y="1554480"/><a:ext cx="10698480" cy="914400"/>',
         '<a:off x="731520" y="1554480"/><a:ext cx="10698480" cy="731520"/>')],

    # ---- mine: volume page. "100% 全評為低" reads as "AI is useless"; the honest
    #      and stronger frame is the team's own 10-15% assist estimate.
    2: [('主採業務的工作範疇 — 14 類工作，AI 可代勞程度全部評為「低」',
         '十四大類的量體 — 工時不跟信件數走，跟事情的難度走'),
        ('100%', '10–15%'),
        ('全數評為「AI 可代勞程度：低」', 'AI 可協助的整體估算，其餘靠人'),
        ('樣本與口徑限制見本章第三頁。', '樣本與口徑限制見「AI 接手的是填表與核對」頁。')],

    3: [('主採業務的工作範疇 — 十四大類（依信箱樣本占比估計）',
         '十四大類與 AI 可協助比例 — 最高的一類也只有 25%，沒有一類能由 AI 獨立完成'),
        ('AI可代勞程度', 'AI可獨立完成')],

    5: [('文書只是表象 — 每一類工作真正吃時間的是人的判斷',
         '文書只是表象 — AI 幫得上的是表象那一層，吃時間的判斷仍在人身上')],

    7: [('重大品質異常個案 — 移染／白汙案跨月延燒',
         '重大品質異常個案 — 移染／白汙跨月延燒，AI 只能協助草稿')],

    # ---- mine: the "why" page, retitled away from "全評為低"
    4: [('為什麼十四類全評為「低」 — 六個共通的結構性原因',
         'AI 為什麼只能協助、不能代替 — 六個共通的結構性原因'),
        ('這三件事沒有一件會因為 AI 而變快。',
         '這三件事 AI 幫得上忙，卻沒有一件能代替人完成。')],

    # ---- mine: the tools page
    6: [('AI 已接手的與接不了的 — 以及這份數字的口徑',
         'AI 接手的是「填表與核對」— 判斷與承諾仍留在人身上')],

    # ---- conclusion: the headline goes in the title (the page is too full for a band)
    11: [('已有 4 支工具覆蓋結構化數字比對，但多數決策仍需人',
          'AI 是協作者，不是替代人力的方案 — 行政步驟壓縮 10–15%，決策仍須人拍板'),
         ('本簡報部分內容由 AI 依既有資料整理，僅供參考；正式引用或對外前請相關單位及本人複核。',
          '剩下的 85–90% 仍然要人來談、來判斷、來承擔。｜本簡報部分內容由 AI 依既有資料整理，'
          '僅供參考；正式引用或對外前請相關單位及本人複核。')],
}

TAKEAWAYS = {
    8:  ([run('→ ', 1300, 1, NAVY),
          run('核色是一連串「看了才知道下一步」的判斷 —— AI 讀不到微信裡的零散對話與圖片，', 1300, 1, INK),
          run('這條鏈上每一環都要人在線上', 1300, 1, RED), run('。', 1300, 1, INK)], 6.36),
    10: ([run('→ ', 1300, 1, NAVY),
          run('AI 可以整理與提示，但「答案合不合理」與「誰來扛」這兩件事，', 1300, 1, INK),
          run('沒有任何一項能交出去', 1300, 1, RED), run('。', 1300, 1, INK)], 6.60),
    9:  ([run('→ ', 1300, 1, NAVY),
          run('四支工具接手的都是「比對與填入」；要不要覆蓋、算不算核可、報告能不能寄出，', 1300, 1, INK),
          run('仍然是人按下確認鍵', 1300, 1, RED), run('。', 1300, 1, INK)], 6.36),
}


# ---------------------------------------------------------------- package work
def main():
    if os.path.isdir(UN):
        shutil.rmtree(UN)
    zipfile.ZipFile(SRC).extractall(UN)
    sl = lambda n: os.path.join(UN, 'ppt/slides/slide%d.xml' % n)

    for n, eyebrow in ORDER:
        x = open(sl(n), encoding='utf-8').read()
        if n in THEIRS:
            x = normalise(x, TAKEAWAY_BOT if n in TAKEAWAYS else BODY_BOT)
            print('  slide%-2d normalised onto the family grid/palette/font' % n)
        if n in EDITS:
            x = sub(x, EDITS[n], 'slide%d' % n)
        if eyebrow:
            x = re.sub(r'(<a:t>)(三、主採|一、主採範疇|二、AI 導入現況|三、結論與建議)(</a:t>)',
                       r'\1' + eyebrow + r'\3', x, count=1)
        if n in TAKEAWAYS:
            runs, y = TAKEAWAYS[n]
            x = add_takeaway(x, runs, y=y)
            print('  slide%-2d closing line added' % n)
        open(sl(n), 'w', encoding='utf-8').write(x)

    # reorder <p:sldIdLst> to ORDER
    prp = os.path.join(UN, 'ppt/_rels/presentation.xml.rels')
    rels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="slides/slide(\d+)\.xml"',
                           open(prp, encoding='utf-8').read()))
    by_slide = {int(v): k for k, v in rels.items()}
    pxp = os.path.join(UN, 'ppt/presentation.xml')
    px = open(pxp, encoding='utf-8').read()
    lst = re.search(r'<p:sldIdLst>(.*?)</p:sldIdLst>', px, re.S).group(1)
    ids = [int(m) for m in re.findall(r'id="(\d+)"', lst)]
    new = ''.join('<p:sldId id="%d" r:id="%s"/>' % (ids[i], by_slide[n])
                  for i, (n, _) in enumerate(ORDER))
    open(pxp, 'w', encoding='utf-8').write(px.replace(lst, new))
    print('  deck reordered: %s' % [n for n, _ in ORDER])

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
