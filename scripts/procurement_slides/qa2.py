# -*- coding: utf-8 -*-
"""Geometry QA: out-of-bounds, real (non-nested) overlaps, estimated text overflow."""
import sys, unicodedata, glob, os, re
import defusedxml.ElementTree as ET
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
P = '{http://schemas.openxmlformats.org/presentationml/2006/main}'
E = 914400.0
def wide(c): return unicodedata.east_asian_width(c) in ('W', 'F')
def wpt(s, sz): return sum(sz if wide(c) else sz * 0.52 for c in s)
def inside(a, b):
    return (a[0] >= b[0]-.01 and a[1] >= b[1]-.01
            and a[0]+a[2] <= b[0]+b[2]+.01 and a[1]+a[3] <= b[1]+b[3]+.01)

root, only = sys.argv[1], set(int(x) for x in sys.argv[2:])
tot_b = tot_o = tot_t = 0
files = sorted(glob.glob(root + '/ppt/slides/slide*.xml'),
               key=lambda f: int(re.search(r'(\d+)', os.path.basename(f)).group(1)))
for f in files:
    n = int(re.search(r'(\d+)', os.path.basename(f)).group(1))
    if only and n not in only: continue
    tree = ET.parse(f).getroot().find(P+'cSld').find(P+'spTree')
    panels, worst = [], []
    for sp in tree:
        o = sp.find('.//'+A+'off'); e = sp.find('.//'+A+'ext')
        if o is None or e is None or None in (o.get('x'), e.get('cx')): continue
        X, Y = int(o.get('x'))/E, int(o.get('y'))/E
        W, H = int(e.get('cx'))/E, int(e.get('cy'))/E
        if X < -.01 or Y < -.01 or X+W > 13.34 or Y+H > 7.51:
            print('  s%d !! OUT OF BOUNDS %.2f %.2f %.2fx%.2f' % (n, X, Y, W, H)); tot_b += 1
        if W >= 0.2 and H <= 7.0: panels.append((X, Y, W, H))
        tx = sp.find('.//'+P+'txBody')
        if tx is None: continue
        need, first = 0, ''
        for p in tx.findall(A+'p'):
            runs = p.findall(A+'r')
            if not runs: continue
            s = ''.join((r.find(A+'t').text or '') for r in runs)
            rpr = runs[0].find(A+'rPr')
            sz = int(rpr.get('sz'))/100.0 if (rpr is not None and rpr.get('sz')) else 18.0
            ls = p.find(A+'pPr/'+A+'lnSpc/'+A+'spcPct')
            pct = int(ls.get('val'))/100000.0 if ls is not None else 1.2
            lines = max(1, -(-int(wpt(s, sz)*100) // max(1, int(W*72*100))))
            need += lines*sz*pct/72.0
            first = first or s[:34]
        if first:
            worst.append((need/H, H, need, first))
            if need > H + 0.02:
                print('  s%d ~~ OVERFLOW h=%.2f need=%.2f  %s' % (n, H, need, first)); tot_t += 1
    for i in range(len(panels)):
        for j in range(i+1, len(panels)):
            a, b = panels[i], panels[j]
            if (min(a[0]+a[2], b[0]+b[2])-max(a[0], b[0]) > .02
                    and min(a[1]+a[3], b[1]+b[3])-max(a[1], b[1]) > .02
                    and not inside(a, b) and not inside(b, a)):
                print('  s%d !! OVERLAP %s %s' % (n, a, b)); tot_o += 1
    worst.sort(reverse=True)
    print('s%d  tightest: ' % n + ' | '.join('%.0f%% %s' % (r*100, t[:22]) for r, _, _, t in worst[:3]))
print('TOTAL out-of-bounds=%d overlaps=%d overflow=%d' % (tot_b, tot_o, tot_t))
