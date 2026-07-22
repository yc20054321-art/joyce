#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PJT 二週差異分析 — 固化腳本
比對同一個 PJT xlsx 內「本週 sheet」與「上週 sheet」的款號差異，
輸出 RECAP Excel（Q×產區二週差異 / 異動說明清單 / 部門產區季度 RECAP 三個 sheet）。

用法：
  python pjt_diff.py <PJT.xlsx> [--current 05-26] [--prev 05-19] [--out <輸出.xlsx>]

若不給 --current / --prev，會自動抓「最前面兩個以 MM-DD 命名的 sheet」當本週/上週。

設計重點（這些是經過多次校正後定下來的規則，請勿隨意更動）：
- 「已進單」= 灰底白字儲存格，不算入開單量。有兩種格式變體都要抓：
    (a) fill theme 0 + 負 tint(<-0.15)；(b) fill theme 1 + 正 tint(>0.3)；或 RGB 灰。
    字色為白(theme 0 或 #FFFFFF)。
- noise 列（TOTAL / DIFF / revised / 上週 PJT / BUFFER / EXP MONTH / 已進單交期 / 安進產能）
    在「款號層級」要過濾掉；但 BUFFER 與「已進單交期/IE調整」在做款號細項說明時可選擇性帶入。
- section 標題以 # 開頭；dept 用 regex 解析；產區依關鍵字 IND/CAB/CHN/N.Vin/SOLO 判斷。
- 官方「部門×產區×季度」表在本週 sheet 的 cols 44-60、rows 4-16（含 total）。對外口徑以此為準。
- 數量單位換算：÷10000 = 萬打；加量(>0.05萬)=藍/減量(<-0.05萬)=紅/持平=黑。
"""
import re
import sys
import argparse
from collections import defaultdict
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as gl
from openpyxl.comments import Comment

# --- 容忍壞掉的圖片參照 (PJT 大表常見 'xl/drawings/NULL' KeyError) ---
# 有些 PJT 檔在編輯後會留下壞掉的 drawing 關聯，讓 openpyxl 載入時整個炸掉。
# 我們不需要圖片，所以攔截 find_images 的錯誤，回空 list，讓儲存格樣式照常讀取。
try:
    import openpyxl.reader.excel as _xlreader
    _orig_find_images = _xlreader.find_images

    def _safe_find_images(archive, path):
        try:
            return _orig_find_images(archive, path)
        except (KeyError, Exception):
            return [], []
    _xlreader.find_images = _safe_find_images
except Exception:
    pass

try:
    from openpyxl.cell.rich_text import TextBlock, CellRichText
    from openpyxl.cell.text import InlineFont
    HAS_RICHTEXT = True
except Exception:
    HAS_RICHTEXT = False

REGIONS = ['印尼', '柬埔寨', '中國', '北越']
QUARTERS = ['Q2', 'Q3', 'Q4', 'Q1']  # 含 2027-Q1
Q_MONTHS = {'Q1': [0, 1, 2], 'Q2': [3, 4, 5], 'Q3': [6, 7, 8], 'Q4': [9, 10, 11]}
DEPT_ORDER = ['#22', '#22  solo', '#23', '#24', '#27', '#kids ', '#32-non solo',
              '#32-solo', '#34 ', '#33', '#37-non solo', '#37 Solo']

TC = '微軟正黑體'
EN = 'Calibri'


# ---------------------------------------------------------------- 解析邏輯
def is_noise(s):
    s = s.strip().upper()
    return any(s.startswith(p) for p in
               ['上週 PJT', 'DIFF', 'REVISED', '已進單交期', '安進產能', 'BUFFER', 'EXP MONTH', 'TOTAL'])


def is_confirmed(cell):
    """灰底白字 = 已進單。兩種 fill 變體 + 白字。"""
    try:
        fill = cell.fill
        font = cell.font
        if fill.patternType != 'solid':
            return False
        fg = fill.fgColor
        grey = False
        if fg.type == 'theme':
            if fg.value == 0 and fg.tint < -0.15:
                grey = True
            elif fg.value == 1 and fg.tint > 0.3:
                grey = True
        elif fg.type == 'rgb' and fg.value:
            rgb = fg.value.upper()
            if len(rgb) == 8:
                r_, g_, b_ = rgb[2:4], rgb[4:6], rgb[6:8]
                if r_ == g_ == b_ and 0x50 <= int(r_, 16) <= 0xC8:
                    grey = True
        if not grey:
            return False
        if font.color is None:
            return False
        fc = font.color
        if fc.type == 'theme' and fc.value == 0:
            return True
        if fc.type == 'rgb' and fc.value and fc.value.upper() in ('FFFFFFFF', '00FFFFFF'):
            return True
        return False
    except Exception:
        return False


def map_region(s):
    u = s.upper()
    if 'NON-SOLO' in u or 'NON SOLO' in u:
        return 'IND-NonSOLO'
    if 'SOLO' in u:
        return 'IND-SOLO'
    if 'CAB' in u:
        return 'CAB'
    if 'CHN' in u:
        return 'CHN'
    if 'VIN' in u:
        return 'N.Vin'
    if 'IND' in u:
        return 'IND'
    return '?'


def map_dept(s):
    m = re.match(r'^(#\s*\d+(?:\s*[&]\s*#?\s*\d+)*)', s)
    return m.group(1).replace(' ', '') if m else None


def parse_sheet(ws, keep_buffer=False):
    """回傳款號層級 rows。keep_buffer=True 時把 BUFFER/IE調整當成 pseudo-style 留下。"""
    rows = []
    sec = dept = reg = None
    for r in range(1, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if not (a and isinstance(a, str)):
            continue
        s = a.strip()
        if s.startswith('#'):
            sec, dept, reg = s, map_dept(s), map_region(s)
            continue
        is_buf = False
        buf_label = ''
        su = s.upper()
        if keep_buffer and su.startswith('BUFFER'):
            is_buf, buf_label = True, 'BUFFER'
        elif keep_buffer and (s.startswith('已進單交期/IE調整') or s.startswith('已進單交期/')):
            is_buf, buf_label = True, 'IE調整'
        elif is_noise(s):
            continue
        monthly = [0.0] * 12
        any_val = False
        for c in range(2, 14):
            cell = ws.cell(r, c)
            v = cell.value
            if isinstance(v, (int, float)) and v != 0:
                any_val = True
                if not is_confirmed(cell):
                    monthly[c - 2] = v
        if not any_val:
            continue
        if is_buf:
            code = buf_label
        else:
            code = ''
            # 不用 \b 邊界：款號後面常接 _pleats / *1.63 / 等非空白字元，
            # 用 \b 會抓不到。直接抓第一個「3數字+字母+3數字(+字母)」樣式即可。
            m = re.search(r'(\d{3}[A-Z]\d{3}[A-Z]?)', s)
            if m:
                code = m.group(1)
        rows.append({'section': sec, 'dept': dept, 'region': reg, 'style': s,
                     'code': code, 'monthly': monthly, 'is_buffer': is_buf})
    return rows


def keyfn(r):
    if r.get('is_buffer'):
        return (r['section'], r['code'])
    return (r['section'], r['code'] if r['code'] else r['style'][:25])


def build_items(cur_rows, prev_rows):
    cur_map, prev_map = defaultdict(list), defaultdict(list)
    for r in cur_rows:
        cur_map[keyfn(r)].append(r)
    for r in prev_rows:
        prev_map[keyfn(r)].append(r)
    items = []
    for k in set(cur_map) | set(prev_map):
        p_m = [sum(x['monthly'][i] for x in prev_map.get(k, [])) for i in range(12)]
        c_m = [sum(x['monthly'][i] for x in cur_map.get(k, [])) for i in range(12)]
        d_m = [c_m[i] - p_m[i] for i in range(12)]
        if all(abs(d) < 1 for d in d_m):
            continue
        ref = (cur_map.get(k) or prev_map.get(k))[0]
        items.append({**ref, 'p_m': p_m, 'c_m': c_m, 'd_m': d_m, 'total': sum(d_m)})
    return items


# ---------------------------------------------------------------- 官方表 & 標註
def read_official_table(ws):
    """本週 sheet cols 44-60, rows 4-15 (部門 × 產區 × 季度)。
    動態讀 row 3 欄位標題，不假設固定 Q 排序（各週 sheet Q 順序可能不同）。
    """
    col_map = {}  # col_idx (1-based) -> (q_key, reg_key)
    for c in range(45, 61):
        h = ws.cell(3, c).value
        if not h:
            continue
        h_str = str(h).strip()
        m = re.match(r'Q[-]?(\d)\s*[-\s]*(.*)', h_str)
        if not m:
            continue
        q_key = 'Q' + m.group(1)
        reg_raw = m.group(2).strip().lstrip('-').strip()
        reg = next((r for r in REGIONS if r in reg_raw), None)
        if q_key in QUARTERS and reg:
            col_map[c] = (q_key, reg)

    out = []
    for r in range(4, 16):
        dept = ws.cell(r, 44).value
        if not dept:
            continue
        dl = str(dept).strip()
        row = {'dept': dl}
        for col, (q, reg) in col_map.items():
            row[(q, reg)] = ws.cell(r, col).value
        out.append(row)
    return out


def read_textbox_annotations(path, sheet_name):
    """從 drawing XML 抓 textbox 文字，對應到 (dept_row_idx, q, region)。
    回傳 dict[(dept_label, q, region)] = annotation_text。
    這是 best-effort：若抓不到就回空 dict，款號細項仍可運作。"""
    import zipfile
    anns = {}
    try:
        with zipfile.ZipFile(path, 'r') as z:
            # find which drawing belongs to the sheet
            wb_xml = z.read('xl/workbook.xml').decode('utf-8', 'ignore')
            sheet_ids = re.findall(r'<sheet[^>]*name="([^"]+)"[^>]*r:id="(rId\d+)"', wb_xml)
            # map sheet name -> sheetN.xml via rels
            rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8', 'ignore')
            rid_target = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
            target = None
            for nm, rid in sheet_ids:
                if nm == sheet_name:
                    target = rid_target.get(rid)
                    break
            if not target:
                return anns
            sheet_file = 'xl/' + target.replace('../', '').lstrip('/')
            sheet_base = sheet_file.split('/')[-1]
            try:
                srels = z.read(f'xl/worksheets/_rels/{sheet_base}.rels').decode('utf-8', 'ignore')
            except KeyError:
                return anns
            m = re.search(r'Target="\.\./drawings/(drawing\d+\.xml)"', srels)
            if not m:
                return anns
            draw = z.read(f'xl/drawings/{m.group(1)}').decode('utf-8', 'ignore')
            anchors = re.findall(r'<xdr:(?:two|one)CellAnchor.*?</xdr:(?:two|one)CellAnchor>', draw, re.DOTALL)
            # Map anchor (col,row 0-indexed) -> nearest (dept,q,region). Detail table cols 44-60 -> 0-idx 43-59
            # Q1 印尼 starts at col 45 (1-idx)=44(0-idx); each Q is 4 cols.
            for anc in anchors:
                fm = re.search(r'<xdr:from>(.*?)</xdr:from>', anc, re.DOTALL)
                texts = re.findall(r'<a:t[^>]*>([^<]+)</a:t>', anc)
                text = ''.join(texts).strip()
                if not (fm and text):
                    continue
                cr = re.search(r'<xdr:col>(\d+)</xdr:col>.*?<xdr:row>(\d+)</xdr:row>', fm.group(1), re.DOTALL)
                if not cr:
                    continue
                col0, row0 = int(cr.group(1)), int(cr.group(2))
                # detail table region columns: 0-idx 44..59 => Q index & region index
                if not (44 <= col0 <= 59):
                    continue
                rel = col0 - 44
                qi, ri = rel // 4, rel % 4
                if qi > 3 or ri > 3:
                    continue
                q = QUARTERS[qi]
                reg = REGIONS[ri]
                # dept row: detail table header at row 3 (0-idx 2). dept rows 4-15 => 0-idx 3..14
                dept_idx = row0 - 3  # 0-based into DEPT_ORDER-ish
                if 0 <= dept_idx < len(DEPT_ORDER):
                    dept = DEPT_ORDER[dept_idx].strip()
                    anns[(dept, q, reg)] = text
    except Exception:
        pass
    return anns


# ---------------------------------------------------------------- 對應款號到官方格
def map_official_dept_codes(o):
    o = o.strip().lower()
    if 'kids' in o:
        return ['#12&#14&#13&#17', '#12&#14&#17', '#12&#14', '#13', '#17']
    for n in ('22', '23', '24', '27', '32', '34', '33', '37'):
        if n in o:
            return ['#' + n]
    return []


def is_solo(o):
    return 'solo' in o.strip().lower() and 'non' not in o.strip().lower()


def is_nonsolo(o):
    return 'non solo' in o.strip().lower() or 'non-solo' in o.strip().lower()


def map_oreg(o):
    return {'印尼': ['IND', 'IND-SOLO', 'IND-NonSOLO'], '柬埔寨': ['CAB'],
            '中國': ['CHN'], '北越': ['N.Vin']}.get(
        next((r for r in REGIONS if r in o), ''), [])


def style_matches(it, odept, oreg):
    tc = map_official_dept_codes(odept)
    if not tc or it['dept'] not in tc:
        return False
    if it['region'] not in map_oreg(oreg):
        return False
    item_solo = 'SOLO' in (it['region'] or '').upper()
    item_nonsolo = 'NON' in (it['region'] or '').upper()
    if is_nonsolo(odept):
        if item_solo and not item_nonsolo:
            return False
    elif is_solo(odept):
        if not item_solo:
            return False
    else:
        if item_solo:
            return False
    return True


def truncate_code(code, style_text=''):
    if not code:
        return style_text.split()[0][:8] if style_text else '?'
    if code in ('BUFFER', 'IE調整'):
        return code
    if len(code) > 5 and code[:3].isdigit():
        return code[3:]
    return code


def categorize(ann):
    if not ann:
        return ''
    au = ann.upper()
    if 'MOVE' in au:
        return '產區移轉'
    if 'NEW ADD' in au:
        return '新增'
    if 'DROP' in au:
        return '取消／剔除'
    if 'ADD' in au:
        return '加量'
    if 'REDUCE' in au:
        return '減量'
    if '退' in ann:
        return '退量'
    return ''


# ---------------------------------------------------------------- 顏色/數值
BLUE, RED, BLACK = '0070C0', 'C00000', '000000'


def label_for(v):
    k = (v or 0) / 10000
    if k > 0.05:
        return ('加量', BLUE)
    if k < -0.05:
        return ('減量', RED)
    return ('持平', BLACK)


def wan(v):
    if not isinstance(v, (int, float)):
        return 0.0
    return 0.0 if abs(v) < 0.5 else round(v / 10000, 1)


# ---------------------------------------------------------------- Excel 樣式
def fz(b=False, c='000000', s=11):
    return Font(name=TC, bold=b, color=c, size=s)


def fe(b=False, c='000000', s=11):
    return Font(name=EN, bold=b, color=c, size=s)


THIN = Side(border_style='thin', color='8FAADC')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_HDR = PatternFill('solid', start_color='305496')
F_ORANGE = PatternFill('solid', start_color='BF8F00')
F_QTOTAL = PatternFill('solid', start_color='FFC000')
F_QBAND = {'Q1': PatternFill('solid', start_color='FFF2CC'),
           'Q2': PatternFill('solid', start_color='D9D2E9'),
           'Q3': PatternFill('solid', start_color='CFE2F3'),
           'Q4': PatternFill('solid', start_color='D9EAD3')}
F_REGION = PatternFill('solid', start_color='F8CBAD')
F_GREY = PatternFill('solid', start_color='F2F2F2')
F_CAT = {'產區移轉': PatternFill('solid', start_color='FFE699'),
         '新增': PatternFill('solid', start_color='C6E0B4'),
         '加量': PatternFill('solid', start_color='C6E0B4'),
         '取消／剔除': PatternFill('solid', start_color='F8CBAD'),
         '減量': PatternFill('solid', start_color='F4B084'),
         '退量': PatternFill('solid', start_color='F4B084'),
         '': PatternFill('solid', start_color='FFFFFF')}


def rich(parts):
    if not parts or not HAS_RICHTEXT:
        return ''.join(p[0] for p in parts) if parts else ''
    return CellRichText(*[TextBlock(InlineFont(rFont=TC, sz=9, b=b, color=c), t) for t, c, b in parts])


# ---------------------------------------------------------------- 主流程
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pjt_file')
    ap.add_argument('--current', default=None, help='本週 sheet 名 (如 05-26)')
    ap.add_argument('--prev', default=None, help='上週 sheet 名 (如 05-19)')
    ap.add_argument('--out', default=None, help='輸出 xlsx 路徑')
    args = ap.parse_args()

    wb_src = load_workbook(args.pjt_file, data_only=True)
    sheets = wb_src.sheetnames
    date_sheets = [s for s in sheets if re.match(r'^\d{2}-\d{2}$', s.strip())]
    cur_name = args.current or (date_sheets[0] if date_sheets else sheets[0])
    prev_name = args.prev or (date_sheets[1] if len(date_sheets) > 1 else sheets[1])
    print(f'本週 sheet: {cur_name} | 上週 sheet: {prev_name}')

    ws_cur = wb_src[cur_name]
    ws_prev = wb_src[prev_name]

    # style-level items (含 buffer 供細項說明)
    cur_rows = parse_sheet(ws_cur, keep_buffer=True)
    prev_rows = parse_sheet(ws_prev, keep_buffer=True)
    items = build_items(cur_rows, prev_rows)

    official = read_official_table(ws_cur)
    anns = read_textbox_annotations(args.pjt_file, cur_name)
    print(f'款號差異 {len(items)} 筆 | 官方部門 {len(official)} 列 | textbox 標註 {len(anns)} 筆')

    # subtotals
    region_q, q_sub, all_total = {}, {}, 0.0
    for q in QUARTERS:
        qs = 0.0
        for reg in REGIONS:
            s = sum(r.get((q, reg), 0) for r in official if isinstance(r.get((q, reg)), (int, float)))
            region_q[(q, reg)] = s
            qs += s
        q_sub[q] = qs
        all_total += qs

    out = args.out or (args.pjt_file.rsplit('.', 1)[0] + ' RECAP 二週差異.xlsx')
    build_recap(out, items, official, anns, region_q, q_sub, all_total)
    print(f'已輸出: {out}')
    print(f'★ 全年淨差: {wan(all_total):+.1f} 萬打 ({all_total:+,.0f} 標打)')
    for q in QUARTERS:
        lbl, _ = label_for(q_sub[q])
        print(f'  {q}: {wan(q_sub[q]):+.1f} 萬打 → {lbl}')


def contributing(items, dept, q, reg):
    qm = Q_MONTHS[q]
    out = []
    for it in items:
        if style_matches(it, dept, reg):
            qd = sum(it['d_m'][m] for m in qm)
            if abs(qd) >= 50:
                out.append({'code': it['code'], 'style': it['style'], 'qd': qd, 'is_buffer': it.get('is_buffer')})
    reg_ = [x for x in out if not x['is_buffer']]
    buf = [x for x in out if x['is_buffer']]
    reg_.sort(key=lambda x: -abs(x['qd']))
    return reg_ + buf


def dept_parts(items, anns, dept, q, reg, official_v):
    """組一個部門在 (q,reg) 的異動說明 rich-text parts。
    規則：
    - 款號/buffer 貢獻四捨五入後為 0.0 萬打的不列出。
    - buffer (BUFFER / IE調整) 用語意化文字：增加→ADD ? 萬打、減少→Reduce ? 萬打。
    - 不顯示「+N 筆」尾巴。
    - 合計只有負數才用紅字，其餘黑字。"""
    tb = anns.get((dept, q, reg))
    contribs = contributing(items, dept, q, reg)
    parts = []
    if tb:
        parts.append((f'【{tb}】 ', BLACK, True))
    sps = []
    for c in contribs:
        if wan(c['qd']) == 0.0:        # 貢獻太小、四捨五入為 0.0 → 不列
            continue
        if c.get('is_buffer'):          # buffer 數量變化用 ADD / Reduce 語意
            mag = abs(c['qd']) / 10000
            sps.append(f'ADD {mag:.1f} 萬打' if c['qd'] > 0 else f'Reduce {mag:.1f} 萬打')
        else:
            code = truncate_code(c['code'], c['style']).replace('待MAPPING', 'buffer').replace('待mapping', 'buffer')
            sps.append(f'{code} {"+" if c["qd"]>=0 else ""}{c["qd"]/10000:.1f}')
    if sps:
        parts.append((' / '.join(sps) + ' ', '595959', False))
    tw = wan(official_v)
    if tw > 0.05:       # 加量 → 藍字
        parts.append((f'合計 +{tw} 萬打', BLUE, True))
    elif tw < -0.05:    # 減量 → 紅字
        parts.append((f'合計 {tw} 萬打', RED, True))
    else:
        parts.append((f'合計 {tw} 萬打', BLACK, True))
    return parts


def build_recap(out_path, items, official, anns, region_q, q_sub, all_total):
    wb = Workbook()
    DEPT_INDEX = {d.strip(): i for i, d in enumerate(DEPT_ORDER)}

    # ===== Sheet 1: Q×產區二週差異 =====
    ws = wb.active
    ws.title = 'Q×產區二週差異'
    ws['A1'] = '5-26 vs 5-19 二週差異 BY 季度 × 產區'
    ws['A1'].font = Font(name=TC, bold=True, size=15, color='FFFFFF')
    ws['A1'].fill = F_ORANGE
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws.merge_cells('A1:G1')
    ws.row_dimensions[1].height = 28
    ws['A2'] = '資料來源：本週 PJT 大表官方表｜單位：萬打 (÷10000)'
    ws['A2'].font = fz(c='595959', s=10)
    ws.merge_cells('A2:G2')
    for j, h in enumerate(['季度', '印尼', '柬埔寨', '中國', '北越', '季度合計', '結論'], 1):
        c = ws.cell(4, j, h)
        c.font = fz(b=True, c='FFFFFF', s=12)
        c.fill = F_HDR
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = BORDER
    row = 5
    for q in QUARTERS:
        c = ws.cell(row, 1, q)
        c.font = fe(b=True, c='305496', s=12)
        c.fill = F_QBAND[q]
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = BORDER
        qt = 0
        for ri, reg in enumerate(REGIONS):
            v = region_q.get((q, reg), 0)
            qt += v
            lbl, col = label_for(v)
            cell = ws.cell(row, 2 + ri, wan(v))
            cell.font = fe(b=abs(v) >= 5000, c=col)
            cell.number_format = '+0.0;-0.0;0.0'
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.fill = F_QBAND[q] if abs(v) > 50 else F_GREY
            cell.border = BORDER
        lbl, col = label_for(qt)
        c = ws.cell(row, 6, wan(qt))
        c.font = fe(b=True, c=col, s=12)
        c.number_format = '+0.0;-0.0;0.0'
        c.alignment = Alignment(horizontal='right', vertical='center')
        c.fill = F_QTOTAL
        c.border = BORDER
        c = ws.cell(row, 7, lbl)
        c.font = fz(b=True, c=col, s=12)
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.fill = F_QTOTAL
        c.border = BORDER
        ws.row_dimensions[row].height = 24
        row += 1
    # 產區合計
    c = ws.cell(row, 1, '產區合計')
    c.font = fz(b=True, c='FFFFFF', s=12)
    c.fill = F_HDR
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.border = BORDER
    for ri, reg in enumerate(REGIONS):
        s = sum(region_q.get((q, reg), 0) for q in QUARTERS)
        lbl, col = label_for(s)
        cell = ws.cell(row, 2 + ri, wan(s))
        cell.font = fe(b=True, c=col)
        cell.number_format = '+0.0;-0.0;0.0'
        cell.alignment = Alignment(horizontal='right', vertical='center')
        cell.fill = F_REGION
        cell.border = BORDER
    lbl, col = label_for(all_total)
    c = ws.cell(row, 6, wan(all_total))
    c.font = fe(b=True, c=col, s=14)
    c.number_format = '+0.0;-0.0;0.0'
    c.alignment = Alignment(horizontal='right', vertical='center')
    c.fill = F_QTOTAL
    c.border = BORDER
    c = ws.cell(row, 7, lbl)
    c.font = fz(b=True, c=col, s=13)
    c.alignment = Alignment(horizontal='center', vertical='center')
    c.fill = F_QTOTAL
    c.border = BORDER
    ws.column_dimensions['A'].width = 14
    for cc in range(2, 7):
        ws.column_dimensions[gl(cc)].width = 13
    ws.column_dimensions['G'].width = 12
    ws.freeze_panes = 'A5'

    # ===== Sheet 2: 異動說明清單 =====
    ws2 = wb.create_sheet('異動說明清單')
    ws2.sheet_properties.outlinePr.summaryBelow = False
    ws2['A1'] = 'PJT 二週差異 — 異動說明清單'
    ws2['A1'].font = Font(name=TC, bold=True, size=15, color='305496')
    ws2.merge_cells('A1:F1')
    ws2['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws2.row_dimensions[1].height = 28
    ws2['A2'] = '排序 Q2→Q3→Q4｜產區 印尼→柬埔寨→中國→北越｜萬打｜部門細項預設收起(點左側 + 展開)'
    ws2['A2'].font = fz(c='595959', s=10)
    ws2.merge_cells('A2:F2')
    for j, h in enumerate(['季度', '產區', '部門', '差異(萬打)', '二週差異', '異動說明'], 1):
        c = ws2.cell(4, j, h)
        c.font = fz(b=True, c='FFFFFF', s=12)
        c.fill = F_HDR
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = BORDER
    r = 5
    for q in QUARTERS:
        lbl, col = label_for(q_sub[q])
        for j, v in enumerate([q + ' 小計', '全產區', 'ALL', wan(q_sub[q]), lbl, ''], 1):
            c = ws2.cell(r, j, v)
            c.border = BORDER
            c.fill = F_QBAND[q]
            c.alignment = Alignment(horizontal='center', vertical='center')
            if j == 4:
                c.font = fe(b=True, c=col, s=12)
                c.number_format = '+0.0;-0.0;0.0'
                c.alignment = Alignment(horizontal='right', vertical='center')
            elif j == 5:
                c.font = fz(b=True, c=col)
            else:
                c.font = fz(b=True, c='305496')
        ws2.row_dimensions[r].height = 22
        r += 1
        for reg in REGIONS:
            rv = region_q.get((q, reg), 0)
            lbl, col = label_for(rv)
            # ALL row rich text — 先「1.加量」(藍) 再「2.減量」(紅)，各部門自己一行。
            # 合計 0.0 的部門不列出。
            inc, dec = [], []
            for d_row in official:
                dl = d_row['dept'].strip()
                v = d_row.get((q, reg))
                w = wan(v)
                if w == 0.0:
                    continue
                (inc if w > 0 else dec).append((dl, v))
            inc.sort(key=lambda x: -x[1])    # 加量大的在前
            dec.sort(key=lambda x: x[1])     # 減量大的在前
            # 換行併進有內容的 run（避免純空白 run 被 Excel 吃掉）
            line_parts = []
            if inc:
                line_parts.append(('1. 加量:', BLUE, True))
                for dl, v in inc:
                    line_parts.append((f'\n　{dl}: ', '305496', True))
                    line_parts.extend(dept_parts(items, anns, dl, q, reg, v))
            if dec:
                line_parts.append(('\n2. 減量:' if line_parts else '2. 減量:', RED, True))
                for dl, v in dec:
                    line_parts.append((f'\n　{dl}: ', '305496', True))
                    line_parts.extend(dept_parts(items, anns, dl, q, reg, v))
            for j, v in enumerate([q, reg, 'ALL', wan(rv), lbl, rich(line_parts) if line_parts else ''], 1):
                c = ws2.cell(r, j, v)
                c.border = BORDER
                c.fill = F_REGION
                c.alignment = Alignment(horizontal='center' if j <= 5 else 'left', vertical='top', wrap_text=(j == 6))
                if j == 4:
                    c.font = fe(b=True, c=col)
                    c.number_format = '+0.0;-0.0;0.0'
                    c.alignment = Alignment(horizontal='right', vertical='center')
                elif j == 5:
                    c.font = fz(b=True, c=col)
                else:
                    c.font = fz(b=True, c='305496' if j in (1, 3) else '000000')
            nlines = 1 + sum(1 for p in line_parts if p[0] == '\n')
            ws2.row_dimensions[r].height = max(22, nlines * 18 + 4)
            r += 1
            # dept detail (collapsed) — 合計為 0.0 萬打的部門不列出
            drows = []
            for d_row in official:
                dl = d_row['dept'].strip()
                v = d_row.get((q, reg))
                if wan(v) != 0.0:
                    drows.append({'dept': dl, 'v': v or 0})
            drows.sort(key=lambda x: DEPT_INDEX.get(x['dept'], 99))
            for d in drows:
                tb = anns.get((d['dept'], q, reg))
                cat = categorize(tb)
                lbl, col = label_for(d['v'])
                parts = dept_parts(items, anns, d['dept'], q, reg, d['v'])
                for j, v in enumerate(['', '', d['dept'], wan(d['v']), lbl, rich(parts)], 1):
                    c = ws2.cell(r, j, v)
                    c.border = BORDER
                    c.fill = F_CAT.get(cat, PatternFill('solid', start_color='FFFFFF'))
                    c.alignment = Alignment(horizontal='center' if j <= 5 else 'left', vertical='center', wrap_text=(j == 6))
                    if j == 3:
                        c.font = fz(b=True)
                    elif j == 4:
                        c.font = fe(b=True, c=col)
                        c.number_format = '+0.0;-0.0;0.0'
                        c.alignment = Alignment(horizontal='right', vertical='center')
                    elif j == 5:
                        c.font = fz(b=True, c=col)
                clen = sum(len(str(p[0])) for p in parts)
                ws2.row_dimensions[r].height = max(22, min(60, 22 + 14 * (clen // 70)))
                ws2.row_dimensions[r].outline_level = 1
                ws2.row_dimensions[r].hidden = True
                r += 1
        r += 1
    for k, w in {1: 11, 2: 11, 3: 18, 4: 14, 5: 11, 6: 100}.items():
        ws2.column_dimensions[gl(k)].width = w
    ws2.freeze_panes = 'A5'

    # ===== Sheet 3: 部門產區季度 RECAP =====
    ws3 = wb.create_sheet('部門產區季度 RECAP')
    ws3['A1'] = '部門 × 產區 × 季度（萬打）'
    ws3['A1'].font = Font(name=TC, bold=True, size=14, color='305496')
    ws3.merge_cells('A1:R1')
    ws3.row_dimensions[1].height = 26
    ws3['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws3.cell(3, 1, '部門').font = fz(b=True, s=12)
    ws3.cell(3, 1).fill = F_REGION
    ws3.cell(3, 1).alignment = Alignment(horizontal='center', vertical='center')
    ws3.cell(3, 1).border = BORDER
    ws3.merge_cells(start_row=3, start_column=1, end_row=4, end_column=1)
    for qi, q in enumerate(QUARTERS):
        sc = 2 + qi * 4
        ws3.merge_cells(start_row=3, start_column=sc, end_row=3, end_column=sc + 3)
        qc = ws3.cell(3, sc, f'2026-{q}')
        qc.font = fe(b=True, c='305496', s=12)
        qc.fill = F_QBAND[q]
        qc.alignment = Alignment(horizontal='center', vertical='center')
        qc.border = BORDER
        for ri, reg in enumerate(REGIONS):
            c = ws3.cell(4, sc + ri, f'{q}-{reg}')
            c.font = fz(b=True, c='305496')
            c.fill = F_QBAND[q]
            c.alignment = Alignment(horizontal='center', vertical='center')
            c.border = BORDER
    ws3.row_dimensions[3].height = 24
    ws3.row_dimensions[4].height = 22
    for di, d_row in enumerate(official):
        er = 5 + di
        dl = d_row['dept'].strip()
        dc = ws3.cell(er, 1, dl)
        dc.font = fz(b=True)
        dc.fill = F_REGION
        dc.alignment = Alignment(horizontal='center', vertical='center')
        dc.border = BORDER
        for qi, q in enumerate(QUARTERS):
            for ri, reg in enumerate(REGIONS):
                v = d_row.get((q, reg))
                col = 2 + qi * 4 + ri
                cell = ws3.cell(er, col)
                tb = anns.get((dl, q, reg))
                if isinstance(v, (int, float)) and abs(v) > 0.5:
                    lbl, color = label_for(v)
                    cell.value = round(v / 10000, 1)
                    cell.font = fe(b=abs(v) >= 5000, c=color)
                    cell.number_format = '+0.0;-0.0;0.0'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    cell.fill = F_CAT.get(categorize(tb), F_QBAND[q])
                    parts = dept_parts(items, anns, dl, q, reg, v)
                    if parts:
                        ct = ''.join(p[0] for p in parts).replace('待MAPPING', 'buffer') + f'\n原值：{int(round(v)):+,} 標打'
                        cell.comment = Comment(ct, 'PJT')
                elif v == 0:
                    cell.value = 0
                    cell.font = fe(c='999999')
                    cell.number_format = '0.0'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                    cell.fill = F_GREY
                else:
                    cell.fill = F_GREY
                cell.border = BORDER
    sub = 5 + len(official)
    ws3.cell(sub, 1, '產區小計').font = fz(b=True, c='FFFFFF')
    ws3.cell(sub, 1).fill = F_HDR
    ws3.cell(sub, 1).alignment = Alignment(horizontal='center', vertical='center')
    ws3.cell(sub, 1).border = BORDER
    for qi, q in enumerate(QUARTERS):
        for ri, reg in enumerate(REGIONS):
            rv = region_q.get((q, reg), 0)
            lbl, col = label_for(rv)
            c = ws3.cell(sub, 2 + qi * 4 + ri, wan(rv))
            c.font = fe(b=True, c=col)
            c.number_format = '+0.0;-0.0;0.0'
            c.alignment = Alignment(horizontal='right', vertical='center')
            c.fill = F_REGION
            c.border = BORDER
    qr = sub + 1
    ws3.cell(qr, 1, '季度小計').font = fz(b=True, c='FFFFFF')
    ws3.cell(qr, 1).fill = F_HDR
    ws3.cell(qr, 1).alignment = Alignment(horizontal='center', vertical='center')
    ws3.cell(qr, 1).border = BORDER
    for qi, q in enumerate(QUARTERS):
        sc = 2 + qi * 4
        lbl, col = label_for(q_sub[q])
        ws3.merge_cells(start_row=qr, start_column=sc, end_row=qr, end_column=sc + 3)
        c = ws3.cell(qr, sc, wan(q_sub[q]))
        c.font = fe(b=True, c=col, s=12)
        c.number_format = '+0.0;-0.0;0.0'
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.fill = F_QTOTAL
        c.border = BORDER
    ws3.column_dimensions['A'].width = 16
    for cc in range(2, 18):
        ws3.column_dimensions[gl(cc)].width = 11
    ws3.freeze_panes = 'B5'

    wb.save(out_path)
    _preserve_whitespace(out_path)


def _preserve_whitespace(path):
    """openpyxl 寫 rich text 時，含換行/空白的 <t> 沒有 xml:space="preserve"，
    Excel 讀取時會把換行(\\n)與前導空白吃掉，導致『減量』黏在『加量』後面不換行。
    這裡在存檔後替所有 <t> 補上 preserve，確保儲存格內換行正常顯示。"""
    import zipfile
    import os
    tmp = path + '.tmp'
    with zipfile.ZipFile(path, 'r') as zin, \
            zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith('xl/worksheets/') or item.filename == 'xl/sharedStrings.xml':
                try:
                    txt = data.decode('utf-8')
                    txt = txt.replace('<t>', '<t xml:space="preserve">')
                    data = txt.encode('utf-8')
                except Exception:
                    pass
            zout.writestr(item, data)
    os.replace(tmp, path)


if __name__ == '__main__':
    main()
