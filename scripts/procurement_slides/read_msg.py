# -*- coding: utf-8 -*-
"""Minimal .msg (OLE compound) reader: subject, sender, body, attachment names."""
import sys, olefile, re, html

f = olefile.OleFileIO(sys.argv[1])
def get(sid, uni=True):
    try:
        d = f.openstream(sid).read()
        return d.decode('utf-16-le' if uni else 'cp950', errors='replace').rstrip('\x00')
    except Exception:
        return None

PROPS = {'subject': '0037', 'sender': '0C1A', 'sendermail': '0C1F',
         'to': '0E04', 'cc': '0E03', 'body': '1000', 'date': '0E06'}
out = {}
for name, tag in PROPS.items():
    for typ in ('001F', '001E'):
        v = get('__substg1.0_%s%s' % (tag, typ), uni=(typ == '001F'))
        if v:
            out[name] = v
            break

html_body = None
for typ in ('001F', '001E', '0102'):
    try:
        d = f.openstream('__substg1.0_1013%s' % typ).read()
        html_body = d.decode('utf-8', errors='replace') if typ == '0102' else \
                    d.decode('utf-16-le' if typ == '001F' else 'cp950', errors='replace')
        break
    except Exception:
        pass

atts = []
for entry in f.listdir():
    if entry[0].startswith('__attach') and entry[-1].startswith('__substg1.0_3707'):
        v = get('/'.join(entry), uni=entry[-1].endswith('001F'))
        if v:
            atts.append(v)
    elif entry[0].startswith('__attach') and entry[-1].startswith('__substg1.0_3704'):
        v = get('/'.join(entry), uni=entry[-1].endswith('001F'))
        if v and v not in atts:
            atts.append(v)

print('SUBJECT :', out.get('subject'))
print('FROM    :', out.get('sender'), out.get('sendermail'))
print('TO      :', (out.get('to') or '')[:300])
print('CC      :', (out.get('cc') or '')[:300])
print('ATTACH  :', sorted(set(atts)))
print('=' * 70)

body = out.get('body')
if not body and html_body:
    t = re.sub(r'(?is)<(script|style)[^>]*>.*?</\1>', ' ', html_body)
    t = re.sub(r'(?i)</(p|div|tr|h[1-6])>', '\n', t)
    t = re.sub(r'(?i)<br[^>]*>', '\n', t)
    t = re.sub(r'(?i)</t[dh]>', ' | ', t)
    t = re.sub(r'<[^>]+>', '', t)
    body = html.unescape(t)
body = re.sub(r'[ \t ]+', ' ', body or '')
body = re.sub(r'\n\s*\n+', '\n', body)
print(body.strip())
