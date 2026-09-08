# -*- coding: utf-8 -*-
"""Cover agenda + section renumbering after inserting 三、主採 into the business deck.

副採 is NOT part of this deck — it ships as its own file (see fucai_standalone.py).
"""
E = 914400
def emu(v): return str(int(round(v * E)))

p = 'unpacked/ppt/slides/slide1.xml'
x = open(p, encoding='utf-8').read()

x = x.replace('<a:t>開發／大貨在做什麼</a:t>', '<a:t>開發／大貨／主採在做什麼</a:t>')

# 三 -> 四, 四 -> 五 (shift the later one first so the two rewrites don't collide)
assert x.count('<a:t>四、</a:t>') == 1 and x.count('<a:t>三、</a:t>') == 1
x = x.replace('<a:t>四、</a:t>', '<a:t>五、</a:t>').replace('<a:t>三、</a:t>', '<a:t>四、</a:t>')

# clone the 二、大貨工作範疇 paragraph for the new 主採 section
i = x.find('<a:t>二、大貨工作範疇</a:t>')
start = x.rfind('<a:p>', 0, i); end = x.find('</a:p>', i) + len('</a:p>')
x = x[:end] + x[start:end].replace('二、大貨工作範疇', '三、主採工作範疇') + x[end:]

# the agenda box now holds 5 lines: tighten leading and give it more room
ag_off = '<a:off x="960120" y="4176412"/><a:ext cx="10241280" cy="1348959"/>'
assert ag_off in x
s = x.find(ag_off); e = x.find('</p:sp>', s)
box = (x[s:e]
       .replace(ag_off, '<a:off x="960120" y="%s"/><a:ext cx="10241280" cy="%s"/>' % (emu(4.42), emu(1.70)))
       .replace('<a:spcPct val="130000"/>', '<a:spcPct val="125000"/>')
       .replace('<a:spcPts val="700"/>', '<a:spcPts val="500"/>'))
x = x[:s] + box + x[e:]

# lift the accent rule to keep its gap above the taller agenda
x = x.replace('<a:off x="960120" y="3913632"/><a:ext cx="2743200" cy="32004"/>',
              '<a:off x="960120" y="%s"/><a:ext cx="2743200" cy="32004"/>' % emu(4.15))
open(p, 'w', encoding='utf-8').write(x)

for f, a, b in [('slide14.xml', '<a:t>三、A</a:t>', '<a:t>四、A</a:t>'),
                ('slide15.xml', '<a:t>四、結論</a:t>', '<a:t>五、結論</a:t>')]:
    q = 'unpacked/ppt/slides/' + f
    y = open(q, encoding='utf-8').read()
    assert a in y, f
    open(q, 'w', encoding='utf-8').write(y.replace(a, b))
print('renumber ok')
