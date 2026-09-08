# -*- coding: utf-8 -*-
"""Geometry QA: bounds, real (non-nested) overlaps, estimated text fit."""
import sys, unicodedata, glob, os, re
import defusedxml.ElementTree as ET
A='{http://schemas.openxmlformats.org/drawingml/2006/main}'
P='{http://schemas.openxmlformats.org/presentationml/2006/main}'
E=914400.0
def wide(c): return unicodedata.east_asian_width(c) in ('W','F')
def wpt(s,sz): return sum(sz if wide(c) else sz*0.52 for c in s)
def inside(a,b):
    return (a[0]>=b[0]-.01 and a[1]>=b[1]-.01
            and a[0]+a[2]<=b[0]+b[2]+.01 and a[1]+a[3]<=b[1]+b[3]+.01)

def check(path, only=None):
    files=sorted(glob.glob(path+'/ppt/slides/slide*.xml'),
                 key=lambda f:int(re.search(r'(\d+)',os.path.basename(f)).group(1)))
    bad_b=bad_o=bad_t=0
    for f in files:
        n=int(re.search(r'(\d+)',os.path.basename(f)).group(1))
        if only and n not in only: continue
        tree=ET.parse(f).getroot().find(P+'cSld').find(P+'spTree')
        panels=[]
        for sp in tree:
            o=sp.find('.//'+A+'off'); e=sp.find('.//'+A+'ext')
            if o is None or e is None: continue
            if None in (o.get('x'), o.get('y'), e.get('cx'), e.get('cy')): continue
            X,Y=int(o.get('x'))/E,int(o.get('y'))/E
            W,H=int(e.get('cx'))/E,int(e.get('cy'))/E
            if X<-.01 or Y<-.01 or X+W>13.34 or Y+H>7.51:
                print('  %s !! OUT OF BOUNDS %.2f %.2f %.2f %.2f'%(os.path.basename(f),X,Y,W,H)); bad_b+=1
            if W>=0.2 and H<=7.0: panels.append((X,Y,W,H))
            tx=sp.find('.//'+P+'txBody')
            if tx is None: continue
            tot=0; first=''
            for p in tx.findall(A+'p'):
                runs=p.findall(A+'r')
                if not runs: continue
                s=''.join((r.find(A+'t').text or '') for r in runs)
                rpr=runs[0].find(A+'rPr'); sz=int(rpr.get('sz'))/100.0 if (rpr is not None and rpr.get('sz')) else 18.0
                ls=p.find(A+'pPr/'+A+'lnSpc/'+A+'spcPct')
                pct=int(ls.get('val'))/100000.0 if ls is not None else 1.2
                lines=max(1,-(-int(wpt(s,sz)*100)//max(1,int(W*72*100))))
                tot+=lines*sz*pct/72.0
                first=first or s[:30]
            if first and tot>H+0.02:
                print('  %s ~~ MAY OVERFLOW h=%.2f need=%.2f  %s'%(os.path.basename(f),H,tot,first)); bad_t+=1
        for i in range(len(panels)):
            for j in range(i+1,len(panels)):
                a,b=panels[i],panels[j]
                if (min(a[0]+a[2],b[0]+b[2])-max(a[0],b[0])>0.02
                        and min(a[1]+a[3],b[1]+b[3])-max(a[1],b[1])>0.02
                        and not inside(a,b) and not inside(b,a)):
                    print('  %s !! OVERLAP %s %s'%(os.path.basename(f),a,b)); bad_o+=1
    print('%s -> out-of-bounds=%d, real overlaps=%d, possible text overflow=%d'%(path,bad_b,bad_o,bad_t))

print('--- 業務deck: 我新增/修改的頁 (1 封面, 16-18 主採) ---')
check('unpacked', only={1,16,17,18})
print('--- 副採deck: 全部 5 頁 ---')
check('fucai_unpacked')
