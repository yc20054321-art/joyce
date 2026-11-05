# -*- coding: utf-8 -*-
"""Dump a .docx as text, keeping table structure."""
import sys, zipfile, re, html
import xml.etree.ElementTree as ET
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
z = zipfile.ZipFile(sys.argv[1])
root = ET.fromstring(z.read('word/document.xml'))
body = root.find(W + 'body')

def para_text(p):
    return ''.join(t.text or '' for t in p.iter(W + 't'))

def walk(el, depth=0):
    for ch in el:
        if ch.tag == W + 'p':
            t = para_text(ch).strip()
            if t:
                print(t)
        elif ch.tag == W + 'tbl':
            print('  ┌─ TABLE')
            for tr in ch.findall(W + 'tr'):
                cells = []
                for tc in tr.findall(W + 'tc'):
                    cells.append(' '.join(para_text(p).strip() for p in tc.iter(W + 'p')).strip())
                print('  │ ' + ' | '.join(cells))
            print('  └─')
walk(body)
