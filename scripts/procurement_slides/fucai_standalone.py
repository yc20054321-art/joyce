# -*- coding: utf-8 -*-
"""Polish the 副採 deck (V6副採工作範疇簡略說明, 2026-09-07) as a STANDALONE file.

It is deliberately not merged into the business deck. Three fixes are applied:

1. 系統操作 page — the four stat cards are misaligned in the source: cards 1 and 2
   overlap by 0.17", the three gaps differ (-0.17 / 0.37 / 0.10), and card 2's label
   is indented 0.37" instead of 0.10". Re-laid on an even grid across the content column.
2. Cover subtitle "(人力編制-3位)" carries no explicit size and a theme colour
   (bg2 @ 75% lum) that is not guaranteed to read on the navy cover. Restyled to the
   same 27pt / C9DFF7 treatment the sister deck's cover subtitle uses. The line itself
   is kept — the headcount belongs on the cover.
3. 工時特性 page — "見上一頁對照表" points three pages back, not one. Reworded.
"""
import os, re, zipfile

E = 914400.0
MARGIN_X, CONTENT_W = 0.62, 12.09
SRC, UN = 'fucai_src.pptx', 'fucai_unpacked'
OUT = 'fucai_out.pptx'

def emu(v): return str(int(round(v * E)))

def fix_stat_row(xml):
    pat = re.compile(r'<a:off x="(-?\d+)" y="(-?\d+)"/><a:ext cx="(\d+)" cy="(\d+)"/>')
    hits = [m for m in pat.finditer(xml)
            if 1.60 <= int(m.group(2)) / E <= 1.90 and 0.60 <= int(m.group(4)) / E <= 1.00]
    cards = [m for m in hits if int(m.group(4)) / E > 0.85]
    labels = [m for m in hits if int(m.group(4)) / E <= 0.85]
    if len(cards) != 4 or len(labels) != 4:
        print('  ! stat row not in the expected shape (%d cards / %d labels) — left as is'
              % (len(cards), len(labels)))
        return xml
    card_w = int(cards[0].group(3)) / E
    gap = (CONTENT_W - card_w * 4) / 3
    edits = {}
    for i, m in enumerate(cards):
        edits[m.span()] = '<a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/>' % (
            emu(MARGIN_X + i * (card_w + gap)), emu(1.70), m.group(3), m.group(4))
    for i, m in enumerate(labels):
        edits[m.span()] = '<a:off x="%s" y="%s"/><a:ext cx="%s" cy="%s"/>' % (
            emu(MARGIN_X + i * (card_w + gap) + 0.10), emu(1.80), m.group(3), m.group(4))
    for span in sorted(edits, reverse=True):
        xml = xml[:span[0]] + edits[span] + xml[span[1]:]
    print('  stat row re-laid: 4 cards, even %.2f" gaps' % gap)
    return xml

def fix_cover_subtitle(xml):
    i = xml.find('人力編制')
    assert i != -1, '(人力編制-3位) missing from the cover'
    s = xml.rfind('<a:p>', 0, i); e = xml.find('</a:p>', i) + len('</a:p>')
    para = xml[s:e]
    styled = re.sub(r'<a:rPr([^>]*?)>\s*<a:solidFill>.*?</a:solidFill>',
                    lambda m: '<a:rPr%s sz="2700" b="1"><a:solidFill><a:srgbClr val="C9DFF7"/></a:solidFill>'
                              % m.group(1).replace(' sz="2700"', '').replace(' b="1"', ''),
                    para, flags=re.S)
    print('  cover subtitle restyled (27pt / C9DFF7), 人力編制-3位 kept')
    return xml[:s] + styled + xml[e:]

def main():
    if os.path.isdir(UN):
        import shutil; shutil.rmtree(UN)
    zipfile.ZipFile(SRC).extractall(UN)

    # read fully before opening for write — open(...,'w') truncates before the arg is evaluated
    p2 = os.path.join(UN, 'ppt/slides/slide2.xml')
    x2 = fix_stat_row(open(p2, encoding='utf-8').read())
    open(p2, 'w', encoding='utf-8').write(x2)

    p1 = os.path.join(UN, 'ppt/slides/slide1.xml')
    x1 = fix_cover_subtitle(open(p1, encoding='utf-8').read())
    open(p1, 'w', encoding='utf-8').write(x1)

    p5 = os.path.join(UN, 'ppt/slides/slide5.xml')
    x5 = open(p5, encoding='utf-8').read()
    assert '見上一頁對照表' in x5
    open(p5, 'w', encoding='utf-8').write(x5.replace('見上一頁對照表', '見「系統操作」頁對照表'))
    print('  page-reference reworded')

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
