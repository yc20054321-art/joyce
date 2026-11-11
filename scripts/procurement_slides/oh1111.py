# -*- coding: utf-8 -*-
"""04 OH 控管方向：人力與費用頁加入 TOMS 轉入 V6 的綜效節省，並將前兩頁對調.

底稿為使用者手改過的 V6 GU 04 OH 控管方向（2026-11-05）。人力與費用那一頁整頁
重建（四張數字卡 → 五張，人力面板增一列），另兩頁只改眉標頁次與因對調而失效的
「次頁」指向。
"""
import os, re, shutil, sys, zipfile
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gen2 import (tb, shape, run, header, takeaway, wrap, th, _id,
                  MARGIN_X, CONTENT_W, NAVY, BLUE, INK, BODY, MUTED,
                  PANEL, TINT, RED, REDBG, SRC_Y, SRC_H)

SRC_DIR, WORK, OUT = 'oh_un2', 'oh_work', 'oh1111_out.pptx'
EYE = "V6 ‧ GU 04 OH 控管方向　"

# 數字卡改用「紅＝不利／綠＝有利」的號誌色，綠色為本版新增
GREEN, GREENBG = "1E7A4D", "EDF6F1"


def src(text):
    return tb(MARGIN_X, SRC_Y, CONTENT_W, SRC_H, [[run(text, 850, 0, MUTED)]], lnspc=122)


# ============================================================ 人力與費用（新 ①）
def slide_people():
    _id[0] = 1
    s = header(EYE + "①", "人力與費用控管方向 — 編制再減 2 人，費用用在開源與避免更大損")

    # (數字, 卡片色, 底色, 前綴(與數字同色), 說明)
    stats = [("104 人", NAVY, PANEL, "", "2027 R0 原編制（含 6D 2 人）"),
             ("102 人", NAVY, PANEL, "", "R0 再減版調整後編制"),
             ("−2 人", GREEN, GREENBG, "", "6A −1（取消台籍新增員額）、6C 印尼端 −1"),
             ("+13 萬", RED, REDBG, "", "2027 R0 差旅費較前期增加（NT$）"),
             ("US$22 萬", GREEN, GREENBG, "≈ NT$660 萬　", "TOMS 轉入 V6 的團隊綜效，可節省 OH")]
    gap = 0.24
    sw = (CONTENT_W - gap * (len(stats) - 1)) / len(stats)
    for i, (num, col, bg, pre, lab) in enumerate(stats):
        x = MARGIN_X + i * (sw + gap)
        s += shape(x, 1.68, sw, 0.98, bg, "roundRect")
        s += shape(x, 1.68, 0.06, 0.98, col)
        s += tb(x + 0.24, 1.77, sw - 0.44, 0.40, [[run(num, 2100, 1, col)]])
        rs = ([run(pre, 900, 1, col)] if pre else []) + [run(lab, 900, 0, BODY)]
        if th(pre + lab, 9.0, sw - 0.44, 122) > 0.38:
            print('  ! 卡片說明超出 0.38"：%s' % (pre + lab)[:30])
        s += tb(x + 0.24, 2.19, sw - 0.44, 0.38, [rs], lnspc=122)

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
        ("TOMS 轉入", "TOMS 轉入 V6，團隊綜效成本預計可節省 OH 0.22／標打；"
                    "2027 年以 100 萬標打估算，預計約 US$22 萬元（≈NT$660 萬）。"),
    ]:
        h = max(th(lab, 9.5, 1.30, 124), th(txt, 9.0, iw - 1.44, 124))
        s += tb(lx + 0.24, y, 1.30, h, [[run("· " + lab, 950, 1, NAVY)]], anchor="ctr", lnspc=124)
        s += tb(lx + 0.24 + 1.44, y, iw - 1.44, h, [[run(txt, 900, 0, BODY)]], anchor="ctr", lnspc=124)
        y += h + 0.12
    if y > py + ph:
        print('  ! 人力面板內容到 %.2f"，面板底部 %.2f"' % (y, py + ph))

    s += tb(rx, 2.74, rw, 0.26, [[run("費用｜兩個方向", 1250, 1, BLUE)]])
    s += shape(rx, py, rw, ph, PANEL, "roundRect")
    s += shape(rx, py, 0.06, ph, BLUE)
    y, iw = py + 0.16, rw - 0.48
    for lab, txt, col in [
        # 對調後差旅費那一頁變成第 ③ 頁，原本的「見次頁」已不成立
        ("1. 差旅費", "2027 R0 較前期增加約 NT$13 萬，來自兩項支援工作：台北特工赴印尼"
                   "控管水洗，以及駐廠／Local Assistant 攜樣赴 JKT 與日本取得客戶核可。"
                   "屬止損性支出 —— 見第 ③ 頁成本效益。", BLUE),
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
        run("人力已按 R0 再減版下修 2 人，TOMS 轉入另有約 US$22 萬（≈NT$660 萬）綜效節省；費用面的增加集中在", 1300, 1, INK),
        run("「避免更大損失」與「開源」", 1300, 1, RED),
        run("兩類，非一般性成長。", 1300, 1, INK),
    ])
    s += src("資料來源：人力及費用相關（2026-11）。金額為預算與實際支出之彙整值，正式引用前請管理部複核。"
             "TOMS 綜效節省為 0.22 USD／標打 × 2027 年 100 萬標打之推估（≈US$22 萬），NT$ 以 1:30 換算；"
             "次頁 2027 R0 預算採匯率 30.50，按該匯率為 NT$671 萬。待確認是否已反映於上述 R0 編制與預算數字。")
    return wrap(s)


# ============================================================ 另兩頁的就地修改
PATCH = {
    # 預算總覽：對調後成為第 ② 頁；它指向編制頁的「次頁」要改成「前頁」
    1: [('<a:t>04 OH 控管方向　①</a:t>', '<a:t>04 OH 控管方向　②</a:t>'),
        ('與次頁「編制 104→102 人」', '與前頁「編制 104→102 人」')],
    # 差旅費頁次不變，只修眉標的「V6  GU‧」排法
    3: [('<a:t>V6  </a:t>', '<a:t>V6 ‧ </a:t>'),
        ('<a:t>‧ 04 OH 控管方向　③</a:t>', '<a:t> 04 OH 控管方向　③</a:t>')],
}

ORDER = ['slide2.xml', 'slide1.xml', 'slide3.xml']


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    os.chdir(here)
    if os.path.isdir(WORK):
        shutil.rmtree(WORK)
    shutil.copytree(SRC_DIR, WORK)

    open(os.path.join(WORK, 'ppt/slides/slide2.xml'), 'w',
         encoding='utf-8').write(slide_people())

    for n, edits in PATCH.items():
        p = os.path.join(WORK, 'ppt/slides/slide%d.xml' % n)
        t = open(p, encoding='utf-8').read()
        for a, b in edits:
            if a in t:
                t = t.replace(a, b, 1)
            else:
                print('  ! slide%d not found: %s' % (n, a[:46]))
        open(p, 'w', encoding='utf-8').write(t)

    pr = open(os.path.join(WORK, 'ppt/_rels/presentation.xml.rels'), encoding='utf-8').read()
    rid = {f: r for r, f in re.findall(r'Id="(rId\d+)"[^>]*Target="slides/(slide\d+\.xml)"', pr)}

    pxp = os.path.join(WORK, 'ppt/presentation.xml')
    px = open(pxp, encoding='utf-8').read()
    lst = re.search(r'<p:sldIdLst>(.*?)</p:sldIdLst>', px, re.S).group(1)
    body = ''.join('<p:sldId id="%d" r:id="%s"/>' % (256 + i, rid[f]) for i, f in enumerate(ORDER))
    open(pxp, 'w', encoding='utf-8').write(
        px.replace('<p:sldIdLst>%s</p:sldIdLst>' % lst, '<p:sldIdLst>%s</p:sldIdLst>' % body))

    if os.path.exists(OUT):
        os.remove(OUT)
    zf = zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED)
    for root, _, files in os.walk(WORK):
        for f in files:
            full = os.path.join(root, f)
            zf.write(full, os.path.relpath(full, WORK))
    zf.close()
    print('wrote %s (%d slides, order %s)' % (OUT, len(ORDER), ' '.join(ORDER)))


if __name__ == '__main__':
    main()
