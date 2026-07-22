#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PJT 二週差異 — 5 張 JPG 渲染器

呼叫範例:
  python pjt_jpgs.py <PJT.xlsx> --current 06-02 --prev 05-26 --out-dir <dir>

產出 (檔名前綴用 --current 那週日期):
  <dir>/PJT MM-DD Q產區矩陣.jpg     — Q×產區 二週差異主表 (Q2-Q4，省略 Q1)
  <dir>/PJT MM-DD Q3 Buffer.jpg    — Q3 Buffer 本週 | 二週差異 並排
  <dir>/PJT MM-DD Q4 Buffer.jpg    — Q4 Buffer 本週 | 二週差異 並排
  <dir>/PJT MM-DD BY部門.jpg        — 部門 × 季度 (Q2/Q3/Q4/ALL，省略 Q1)
  <dir>/PJT MM-DD BY產區.jpg        — 二週差異 + 季度差異 + 上週 + 本週 + 調整後季 完整全塊

設計重點:
- 共用 pjt_diff.py 的 is_confirmed() / wan() / openpyxl 壞圖容錯。
- 配色與 RECAP Excel 一致：加量藍 / 減量紅 / 持平黑；季度分色。
- 數字單位：標打 (BY 部門/BY 產區) 或 萬打 (Q×產區矩陣/Buffer)。
- BY 產區的 Apr/May 註解 + 本週CHN 註解 + 上週IND Dec 註解 自動帶入。
"""
import re
import sys
import argparse
import os
from openpyxl import load_workbook
import openpyxl.reader.excel as _xlreader

# --- 容忍 PJT 大表常見的壞 drawing 參照 ---
_orig_find_images = _xlreader.find_images
def _safe_find_images(archive, path):
    try:
        return _orig_find_images(archive, path)
    except Exception:
        return [], []
_xlreader.find_images = _safe_find_images

import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'SimHei', 'Arial']
matplotlib.rcParams['axes.unicode_minus'] = False
import matplotlib.pyplot as plt
import matplotlib.patches as patches

# ========= 共用配色與常數 =========
BLUE, RED, BLACK = '#0070C0', '#C00000', '#000000'
GREEN = '#00B050'  # Buffer 減少（好事）用綠字；Buffer 增加用紅字
F_HDR = '#305496'
F_QTOTAL = '#FFC000'
F_QBAND = {'Q1': '#FFF2CC', 'Q2': '#D9D2E9', 'Q3': '#CFE2F3', 'Q4': '#D9EAD3'}
F_REGION = '#F8CBAD'
F_GREY = '#F2F2F2'
MONTH_BAND = {'Apr': '#E4DFEC', 'May': '#E4DFEC', 'JUN': '#E4DFEC',
              'Jul': '#DAEEF3', 'Aug': '#DAEEF3', 'Sep': '#DAEEF3',
              'Oct': '#E2EFDA', 'Nov': '#E2EFDA', 'Dec': '#E2EFDA',
              '27-Jan': '#FFF2CC', '27-Feb': '#FFF2CC', '27-Mar': '#FFF2CC'}
REGIONS = ['印尼', '柬埔寨', '中國', '北越']
QUARTERS = ['Q1', 'Q2', 'Q3', 'Q4']

REGION_LABEL = {'IND': 'IND (印尼)', 'CAB': 'CAB (柬埔寨)',
                'CHN': 'CHN (中國)', 'NVN': 'NVN (北越)'}


def wan(v):
    if not isinstance(v, (int, float)):
        return 0.0
    return 0.0 if abs(v) < 0.5 else round(v / 10000, 1)


def label_color(v):
    k = (v or 0) / 10000 if isinstance(v, (int, float)) else 0
    if k > 0.05:
        return ('加量', BLUE)
    if k < -0.05:
        return ('減量', RED)
    return ('持平', BLACK)


# ========= 1) Q×產區 矩陣 =========
def render_q_region_matrix(ws, out_path):
    """Q2-Q4 + 2027-Q1 × 4 產區 + 季度合計 + 產區合計（動態讀 row 3 標題）"""
    # 動態讀 row 3 欄位標題，避免各週 sheet Q 排序不同造成欄位錯位
    col_map = {}
    for c in range(45, 61):
        h = ws.cell(3, c).value
        if not h:
            continue
        m = re.match(r'Q[-]?(\d)\s*[-\s]*(.*)', str(h).strip())
        if not m:
            continue
        q_key = 'Q' + m.group(1)
        reg_raw = m.group(2).strip().lstrip('-').strip()
        reg = next((r for r in REGIONS if r in reg_raw), None)
        if reg:
            col_map[(q_key, reg)] = c

    totals = {}
    for (q, reg), c in col_map.items():
        v = ws.cell(16, c).value
        totals[(q, reg)] = v if isinstance(v, (int, float)) else 0

    # Q2-Q4 + 2027-Q1
    display_quarters = ['Q2', 'Q3', 'Q4', 'Q1']
    display_rows = [[q] + [totals.get((q, r), 0) / 10000 for r in REGIONS] +
                    [sum(totals.get((q, r), 0) for r in REGIONS) / 10000] for q in display_quarters]
    # 產區合計
    sum_row = ['產區合計'] + [sum(totals.get((q, r), 0) for q in display_quarters) / 10000 for r in REGIONS] + \
              [sum(totals.get((q, r), 0) for q in display_quarters for r in REGIONS) / 10000]
    display_rows.append(sum_row)

    header = ['季度', '印尼', '柬埔寨', '中國', '北越', '季度合計', '結論']
    fig, ax = plt.subplots(figsize=(8.5, 3.0))
    ax.axis('off')
    col_widths = [0.10, 0.13, 0.13, 0.13, 0.13, 0.18, 0.20]
    xs = [sum(col_widths[:i]) for i in range(len(header) + 1)]
    row_h = 0.18
    y0 = 0.95
    # Header
    for ci, h in enumerate(header):
        ax.add_patch(patches.Rectangle((xs[ci], y0 - row_h), col_widths[ci], row_h,
                                       facecolor=F_HDR, edgecolor='#8FAADC', linewidth=0.8))
        ax.text(xs[ci] + col_widths[ci] / 2, y0 - row_h / 2, h,
                ha='center', va='center', fontsize=11, color='white', weight='bold')

    for ri, row in enumerate(display_rows):
        y = y0 - row_h * (ri + 2)
        q_label = row[0]
        band = F_QBAND.get(q_label, F_REGION if q_label == '產區合計' else '#FFFFFF')
        for ci, val in enumerate(row):
            ax.add_patch(patches.Rectangle((xs[ci], y), col_widths[ci], row_h,
                                           facecolor=band, edgecolor='#8FAADC', linewidth=0.5))
            if ci == 0:
                ax.text(xs[ci] + col_widths[ci] / 2, y + row_h / 2, q_label,
                        ha='center', va='center', fontsize=11, color=F_HDR, weight='bold')
            elif ci == 6:
                lbl, col = label_color((val or 0) * 10000)
                ax.text(xs[ci] + col_widths[ci] / 2, y + row_h / 2, lbl,
                        ha='center', va='center', fontsize=11, color=col, weight='bold')
            else:
                v = val or 0
                if abs(v) < 0.05:
                    text = '0.0'
                    text_color = '#666666'
                else:
                    text = f'{v:+.1f}'
                    text_color = RED if v < 0 else BLUE
                weight = 'bold' if abs(v) >= 0.1 else 'normal'
                if ci == 5:
                    weight = 'bold'
                ax.text(xs[ci] + col_widths[ci] / 2, y + row_h / 2, text,
                        ha='right', va='center', fontsize=11, color=text_color, weight=weight)

    ax.set_xlim(0, sum(col_widths))
    ax.set_ylim(0, 1)
    plt.title('2026 Q2-Q4 + 2027-Q1 二週差異 BY 季度 × 產區（單位：萬打）', fontsize=12, weight='bold',
              color=F_HDR, pad=10)
    plt.savefig(out_path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()


# ========= 2) Q3 / Q4 Buffer (本週 | 二週差異 並排) =========
def _read_buffer_pivot(ws, col_start, col_end):
    rows = []
    for r in range(10, 60):
        vals = [ws.cell(r, c).value for c in range(col_start, col_end + 1)]
        if all(v is None for v in vals):
            continue
        if vals[0] == '列標籤':
            continue
        rows.append(vals)
        if vals[0] == '總計':
            break
    return rows


def render_buffer(q_name, months, this_week_pivot, two_diff_pivot, total_wan, diff_wan, out_path):
    n_max = max(len(this_week_pivot), len(two_diff_pivot))
    fig, axes = plt.subplots(1, 2, figsize=(14, max(4.5, 0.30 * n_max + 1.5)))

    def draw_table(ax, title, rows, months_local, is_diff):
        ax.axis('off')
        cols = ['列標籤'] + months_local + ['總計']
        col_widths = [0.24] + [0.16] * len(months_local) + [0.20]
        xs = [sum(col_widths[:i]) for i in range(len(cols) + 1)]
        nrows = len(rows) + 1
        row_h = 0.88 / nrows
        y0 = 0.95
        for ci, h in enumerate(cols):
            ax.add_patch(patches.Rectangle((xs[ci], y0 - row_h), col_widths[ci], row_h,
                                           facecolor=F_HDR, edgecolor='#8FAADC', linewidth=0.8))
            ax.text(xs[ci] + col_widths[ci] / 2, y0 - row_h / 2, h,
                    ha='center', va='center', fontsize=9, color='white', weight='bold')
        for ri, row in enumerate(rows):
            y = y0 - row_h * (ri + 2)
            label = row[0] or ''
            is_region = label in ('IND', 'CAB', 'CHN', 'NVN', '總計')
            display_label = REGION_LABEL.get(label, ('　　' + label) if not is_region else label)
            band = '#FFE699' if label == '總計' else ('#D9E1F2' if is_region else '#FFFFFF')
            for ci, val in enumerate(row):
                ax.add_patch(patches.Rectangle((xs[ci], y), col_widths[ci], row_h,
                                               facecolor=band, edgecolor='#BFBFBF', linewidth=0.5))
                if ci == 0:
                    text_color = F_HDR if is_region else BLACK
                    weight = 'bold' if is_region else 'normal'
                    ax.text(xs[ci] + 0.01, y + row_h / 2, display_label, ha='left', va='center',
                            fontsize=9, color=text_color, weight=weight)
                else:
                    if not isinstance(val, (int, float)):
                        text, color, weight = '-', '#999999', 'normal'
                    elif val == 0:
                        text, color, weight = '0', '#999999', 'normal'
                    else:
                        v_wan = val / 10000
                        text = f'{v_wan:+.1f}' if is_diff else f'{v_wan:.1f}'
                        if is_diff:
                            # Buffer 差異：減少=好(綠) / 增加=不好(紅) / 持平=黑
                            color = GREEN if v_wan < -0.05 else (RED if v_wan > 0.05 else BLACK)
                        else:
                            # Buffer 本週數量：純絕對值，正藍
                            color = BLUE if v_wan > 0.05 else BLACK
                        weight = 'bold' if abs(v_wan) >= 0.1 else 'normal'
                    ax.text(xs[ci] + col_widths[ci] / 2, y + row_h / 2, text,
                            ha='center', va='center', fontsize=9, color=color, weight=weight)
        ax.set_title(title, fontsize=11, weight='bold', color=F_HDR, pad=4)
        ax.set_xlim(0, sum(col_widths))
        ax.set_ylim(0, 1)

    draw_table(axes[0], f'本週 {q_name} Buffer 數量（萬打）', this_week_pivot, months, False)
    draw_table(axes[1], f'{q_name} Buffer 二週差異（萬打）', two_diff_pivot, months, True)

    # Buffer 增加=不好(紅) / 減少=好(綠) / 持平=黑
    if diff_wan > 0.05:
        diff_word, diff_color = '增加', RED
    elif diff_wan < -0.05:
        diff_word, diff_color = '減少', GREEN
    else:
        diff_word, diff_color = '持平', BLACK
    summary = (f'★ {q_name} Buffer 共計 {total_wan:.1f} 萬打，本週{diff_word} {abs(diff_wan):.1f} 萬打'
               if abs(diff_wan) > 0.05
               else f'★ {q_name} Buffer 共計 {total_wan:.1f} 萬打，本週持平')
    fig.suptitle(summary, fontsize=14, weight='bold', color=diff_color, y=0.99)
    plt.tight_layout(rect=[0, 0, 1, 0.94])
    plt.savefig(out_path, dpi=140, bbox_inches='tight', facecolor='white')
    plt.close()


def render_q3_buffer(buffer_ws, out_path):
    this_w = _read_buffer_pivot(buffer_ws, 26, 30)
    diff = _read_buffer_pivot(buffer_ws, 32, 36)
    total_wan = (this_w[-1][-1] or 0) / 10000 if this_w else 0
    diff_wan = (diff[-1][-1] or 0) / 10000 if diff else 0
    render_buffer('Q3', ['Jul', 'Aug', 'Sep'], this_w, diff, total_wan, diff_wan, out_path)


def render_q4_buffer(buffer_ws, out_path):
    this_w = _read_buffer_pivot(buffer_ws, 8, 12)
    diff = _read_buffer_pivot(buffer_ws, 14, 18)
    total_wan = (this_w[-1][-1] or 0) / 10000 if this_w else 0
    diff_wan = (diff[-1][-1] or 0) / 10000 if diff else 0
    render_buffer('Q4', ['OCT', 'NOV', 'DEC'], this_w, diff, total_wan, diff_wan, out_path)


# ========= 3) BY 部門 (Q3/Q4/ALL/2027-Q1，省略 Q2) =========
def render_by_dept(ws, out_path):
    """官方部門季度表，動態找 2026 Q3/Q4/ALL 及 2027-Q1 欄（Q2 不顯示）"""
    # col 35 = 部門；掃 36-46 找 2026 Q3/Q4/ALL 或 2027-Q1（跳過 Q2）
    target_cols = [35]
    for c in range(35, 46):
        h = str(ws.cell(3, c).value or '').strip()
        if ('2026' in h and any(k in h for k in ('Q3', 'Q4', 'ALL'))) or \
           ('2027' in h and 'Q1' in h):
            target_cols.append(c)
    cols = target_cols[:5]  # 部門 + Q3/Q4/ALL/2027-Q1
    header = [ws.cell(3, c).value for c in cols]
    rows = [[ws.cell(r, c).value for c in cols] for r in range(4, 17)]

    Q_BAND_OFFICIAL = {'2026-Q2': '#E4DFEC', '2026-Q3': '#DAEEF3',
                       '2026-Q4': '#E2EFDA', '2026ALL': '#FCE4D6',
                       '2027-Q1': '#FFF2CC'}
    HDR_COLOR = '#FBE5D6'
    TOTAL_COLOR = '#F4B084'

    fig, ax = plt.subplots(figsize=(9, 7.5))
    ax.axis('off')
    n_cols = len(cols)
    if n_cols == 5:
        col_widths = [0.20, 0.17, 0.17, 0.20, 0.26]
    else:
        col_widths = [0.20, 0.19, 0.19, 0.19, 0.23][:n_cols]
    xs = [sum(col_widths[:i]) for i in range(len(cols) + 1)]
    nrows = len(rows) + 1
    row_h = 0.95 / nrows
    y0 = 0.98

    for ci, h in enumerate(header):
        band = HDR_COLOR if ci == 0 else Q_BAND_OFFICIAL.get(h, '#FFFFFF')
        ax.add_patch(patches.Rectangle((xs[ci], y0 - row_h), col_widths[ci], row_h,
                                       facecolor=band, edgecolor='#BFBFBF', linewidth=0.8))
        ax.text(xs[ci] + col_widths[ci] / 2, y0 - row_h / 2, str(h) if h else '',
                ha='center', va='center', fontsize=11, color=BLACK, weight='bold')

    for ri, row in enumerate(rows):
        y = y0 - row_h * (ri + 2)
        dept = row[0]
        is_total = (dept and 'total' in str(dept).lower())
        for ci, val in enumerate(row):
            if is_total:
                band = TOTAL_COLOR if ci == 0 else '#FFFFFF'
            elif ci == 0:
                band = HDR_COLOR
            elif Q_BAND_OFFICIAL.get(str(header[ci] or '')):
                band = Q_BAND_OFFICIAL[str(header[ci])]
            else:
                band = '#FFFFFF'
            ax.add_patch(patches.Rectangle((xs[ci], y), col_widths[ci], row_h,
                                           facecolor=band, edgecolor='#BFBFBF', linewidth=0.5))
            if ci == 0:
                ax.text(xs[ci] + 0.015, y + row_h / 2, str(val) if val else '',
                        ha='left', va='center', fontsize=10, color=BLACK, weight='bold')
            else:
                if val is None:
                    text, text_color, weight = '', BLACK, 'normal'
                elif isinstance(val, (int, float)):
                    if abs(val) < 0.5:
                        text, text_color, weight = '0', BLUE, 'normal'
                    elif val < 0:
                        text, text_color, weight = f'({abs(val):,.0f})', RED, 'bold'
                    else:
                        text, text_color, weight = f'{val:,.0f}', BLUE, 'bold'
                else:
                    text, text_color, weight = str(val), BLACK, 'normal'
                ax.text(xs[ci] + col_widths[ci] - 0.015, y + row_h / 2, text,
                        ha='right', va='center', fontsize=10, color=text_color, weight=weight)

    ax.set_xlim(0, sum(col_widths))
    ax.set_ylim(0, 1)
    plt.title('2026 Q3-Q4 + 2027-Q1 二週差異 BY 部門（單位：標打）', fontsize=12, weight='bold',
              color=F_HDR, pad=12)
    plt.savefig(out_path, dpi=150, bbox_inches='tight', facecolor='white')
    plt.close()


# ========= 4) BY 產區 (完整全塊：二週差異 + 季度差異 + 上週 + 本週 + 調整後季) =========
def render_by_region(ws, out_path):
    """rows 3-23，顯示 2026 Q3+Q4（Jul-Dec）+ 2027-Q1（Jan-Mar），省略 Q2（Apr-Jun）"""
    # 動態掃 row 3，找 Q3+Q4 2026 月份 及 2027-Q1 月份，以及 2026 TOTAL
    _ABBR_2026 = {'JUL': 'Jul', 'AUG': 'Aug', 'SEP': 'Sep',
                  'OCT': 'Oct', 'NOV': 'Nov', 'DEC': 'Dec'}
    _ABBR_2027 = {'JAN': '27-Jan', 'FEB': '27-Feb', 'MAR': '27-Mar'}
    _ORDER = ['Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec', '27-Jan', '27-Feb', '27-Mar']
    _month_col = {}
    _last_col = 31
    for c in range(18, 42):
        h = str(ws.cell(3, c).value or '').strip()
        if '2026' in h and 'TOTAL' in h.upper():
            _last_col = c
            continue
        if h.startswith('2026-') or h.startswith('2026 '):
            mp = h[5:].strip().upper()
            if mp in _ABBR_2026:
                _month_col[_ABBR_2026[mp]] = c
        elif h.startswith('2027-') or h.startswith('2027 '):
            mp = h[5:].strip().upper()
            if mp in _ABBR_2027:
                _month_col[_ABBR_2027[mp]] = c
    MONTH_COLS = [_month_col[m] for m in _ORDER if m in _month_col]
    MONTH_NAMES = [m for m in _ORDER if m in _month_col]
    LAST_COL = _last_col
    _dec_col = _month_col.get('Dec', MONTH_COLS[5] if len(MONTH_COLS) > 5 else 27)

    def C(r, c):
        return ws.cell(r, c).value

    rows = []
    rows.append({'kind': 'header1', 'label': 'EXP MONTH\nR3:AE23', 'last_label': '2026\nTOTAL'})
    for r, label in [(4, 'IND--二週差異 Non-SOLO'), (5, 'IND--二週差異- SOLO'),
                     (6, 'CAB--二週差異'), (7, 'CHN-二週差異'), (8, 'N.Vin-二週差異')]:
        rows.append({'kind': 'diff', 'label': label,
                     'vals': [C(r, c) for c in MONTH_COLS], 'last': C(r, LAST_COL)})
    rows.append({'kind': 'total', 'label': 'TOTAL',
                 'vals': [C(9, c) for c in MONTH_COLS], 'last': C(9, LAST_COL)})
    rows.append({'kind': 'qdiff', 'label': '季度差異',
                 'vals': [C(10, c) for c in MONTH_COLS], 'last': C(10, LAST_COL)})
    rows.append({'kind': 'spacer'})
    rows.append({'kind': 'header2', 'label': 'EXP MONTH', 'last_label': '2026\nTOTAL'})
    for r, label in [(12, '上週 IND Non-SOLO'), (13, '上週 IND SOLO'),
                     (14, '上週 CAB'), (15, '上週 CHN'), (16, '上週 N.Vin')]:
        note_dec = ''
        if r == 12:
            cell_val = C(12, _dec_col)
            if isinstance(cell_val, str) and '本週' in cell_val:
                note_dec = cell_val
        rows.append({'kind': 'lastwk', 'label': label,
                     'vals': [C(r, c) for c in MONTH_COLS], 'last': C(r, LAST_COL),
                     'note_dec': note_dec})
    rows.append({'kind': 'spacer'})
    rows.append({'kind': 'header3', 'label': 'EXP MONTH', 'last_label': '2026\nTOTAL'})
    for r, label in [(18, '本週 IND Non solo'), (19, '本週 IND solo'),
                     (20, '本週 CAB'), (21, '本週 CHN'), (22, '本週 N.Vin')]:
        label_extra = ''
        if r == 21:
            cell_lbl = C(21, 18)
            if isinstance(cell_lbl, str) and '\n' in cell_lbl:
                label_extra = '\n'.join(cell_lbl.split('\n')[1:]).strip()
        rows.append({'kind': 'thiswk', 'label': label,
                     'vals': [C(r, c) for c in MONTH_COLS], 'last': C(r, LAST_COL),
                     'label_extra': label_extra})
    rows.append({'kind': 'adj', 'label': '調整後季',
                 'vals': [C(23, c) for c in MONTH_COLS], 'last': None})

    ncols = 1 + len(MONTH_COLS) + 1
    col_widths = [0.16] + [0.075] * 9 + [0.105]
    xs = [sum(col_widths[:i]) for i in range(ncols + 1)]
    total_w = sum(col_widths)

    row_heights = []
    for row in rows:
        k = row['kind']
        if k == 'spacer':
            row_heights.append(0.6)
        elif k.startswith('header'):
            row_heights.append(1.8)
        elif k == 'lastwk' and row.get('note_dec'):
            row_heights.append(2.6)
        elif k == 'thiswk' and row.get('label_extra'):
            row_heights.append(2.2)
        else:
            row_heights.append(1.4)
    total_h = sum(row_heights)
    unit_h = 0.97 / total_h

    fig, ax = plt.subplots(figsize=(16, 0.30 * total_h + 0.6))
    ax.axis('off')

    def fmt_int(v):
        if v is None or (isinstance(v, str) and not v):
            return ''
        if not isinstance(v, (int, float)):
            return str(v)[:20]
        if abs(v) < 0.5:
            return '0'
        if v < 0:
            return f'({abs(v):,.0f})'
        return f'{v:,.0f}'

    def color_for(v, kind):
        if not isinstance(v, (int, float)):
            return BLACK, 'normal'
        if kind in ('diff', 'total', 'qdiff', 'adj'):
            if abs(v) < 0.5:
                return BLUE, 'normal'
            if v < 0:
                return RED, 'bold'
            return BLUE, 'bold'
        return BLACK, 'normal'

    y_cursor = 0.98
    for ri, row in enumerate(rows):
        h = row_heights[ri] * unit_h
        y = y_cursor - h
        k = row['kind']
        if k == 'spacer':
            y_cursor = y
            continue
        if k.startswith('header'):
            ax.add_patch(patches.Rectangle((xs[0], y), col_widths[0], h,
                                           facecolor=F_HDR, edgecolor='#8FAADC', linewidth=0.8))
            ax.text(xs[0] + col_widths[0] / 2, y + h / 2, row['label'],
                    ha='center', va='center', fontsize=9, color='white', weight='bold')
            for ci_, m in enumerate(MONTH_NAMES):
                ci = ci_ + 1
                band = MONTH_BAND[m]
                ax.add_patch(patches.Rectangle((xs[ci], y), col_widths[ci], h,
                                               facecolor=band, edgecolor='#BFBFBF', linewidth=0.5))
                sub = ''
                ax.text(xs[ci] + col_widths[ci] / 2, y + h / 2, m,
                        ha='center', va='center', fontsize=9, color=BLACK, weight='bold')
            ci = ncols - 1
            last_band = '#F4B6B6' if '平均' in row.get('last_label', '') else '#FCE4D6'
            ax.add_patch(patches.Rectangle((xs[ci], y), col_widths[ci], h,
                                           facecolor=last_band, edgecolor='#BFBFBF', linewidth=0.5))
            ax.text(xs[ci] + col_widths[ci] / 2, y + h / 2, row['last_label'],
                    ha='center', va='center', fontsize=8.5, color=BLACK, weight='bold')
            y_cursor = y
            continue

        lbl_band = '#FFFF00' if k == 'total' else '#FFFFFF'
        ax.add_patch(patches.Rectangle((xs[0], y), col_widths[0], h,
                                       facecolor=lbl_band, edgecolor='#BFBFBF', linewidth=0.5))
        if k == 'thiswk' and row.get('label_extra'):
            ax.text(xs[0] + 0.005, y + h * 0.80, row['label'], ha='left', va='center',
                    fontsize=9, color=BLACK, weight='bold')
            parts = row['label_extra'].split('\n')
            ax.text(xs[0] + 0.005, y + h * 0.40, parts[0], ha='left', va='center',
                    fontsize=7, color=BLUE)
            if len(parts) > 1:
                ax.text(xs[0] + 0.005, y + h * 0.20, ' '.join(parts[1:]).strip(),
                        ha='left', va='center', fontsize=7, color=BLUE)
        else:
            ax.text(xs[0] + 0.005, y + h / 2, row['label'], ha='left', va='center',
                    fontsize=9, color=BLACK,
                    weight='bold' if k in ('total', 'qdiff', 'adj') else 'normal')

        vals = row.get('vals', [None] * 9)
        for ci_, (m, val) in enumerate(zip(MONTH_NAMES, vals)):
            ci = ci_ + 1
            if k == 'total':
                band = '#FFFF00'
            elif k in ('qdiff', 'adj'):
                band = '#FFFFFF'
            else:
                band = MONTH_BAND[m]
            ax.add_patch(patches.Rectangle((xs[ci], y), col_widths[ci], h,
                                           facecolor=band, edgecolor='#BFBFBF', linewidth=0.5))

            if k == 'qdiff':
                text = fmt_int(val) if m in ('Apr', 'Jul', 'Oct') else ''
                text_color, weight = color_for(val if text else None, k)
            else:
                text = fmt_int(val)
                text_color, weight = color_for(val, k)
            if k == 'total':
                weight = 'bold'

            if k == 'lastwk' and m == 'Dec' and row.get('note_dec'):
                ax.text(xs[ci] + col_widths[ci] - 0.003, y + h * 0.78, text,
                        ha='right', va='center', fontsize=9, color=text_color, weight=weight)
                note_lines = row['note_dec'].split('\n')
                ax.text(xs[ci] + col_widths[ci] / 2, y + h * 0.40, note_lines[0][:25],
                        ha='center', va='center', fontsize=6.5, color=RED)
                if len(note_lines) > 1:
                    ax.text(xs[ci] + col_widths[ci] / 2, y + h * 0.18, note_lines[1][:25],
                            ha='center', va='center', fontsize=6.5, color=RED)
            else:
                ax.text(xs[ci] + col_widths[ci] - 0.003, y + h / 2, text,
                        ha='right', va='center', fontsize=9, color=text_color, weight=weight)

        ci = ncols - 1
        if k == 'total':
            last_band = '#FFFF00'
        elif k in ('qdiff', 'adj'):
            last_band = '#FFFFFF'
        elif k in ('lastwk', 'thiswk'):
            last_band = '#F4B6B6'
        else:
            last_band = '#FCE4D6'
        ax.add_patch(patches.Rectangle((xs[ci], y), col_widths[ci], h,
                                       facecolor=last_band, edgecolor='#BFBFBF', linewidth=0.5))
        if k != 'adj':
            text = fmt_int(row.get('last'))
            text_color, weight = color_for(row.get('last'), k)
            if k in ('lastwk', 'thiswk'):
                text_color, weight = BLUE, 'bold'
            ax.text(xs[ci] + col_widths[ci] - 0.003, y + h / 2, text,
                    ha='right', va='center', fontsize=9, color=text_color, weight=weight)
        y_cursor = y

    ax.set_xlim(0, total_w)
    ax.set_ylim(0, 1)
    plt.title('2026 PJT BY 產區 × 月份（單位：標打）', fontsize=12, weight='bold',
              color=F_HDR, pad=8)
    plt.savefig(out_path, dpi=140, bbox_inches='tight', facecolor='white')
    plt.close()


# ========= 主流程 =========
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('pjt_file')
    ap.add_argument('--current', default=None, help='本週 sheet 名 (如 06-02)')
    ap.add_argument('--prev', default=None, help='上週 sheet 名 (備用)')
    ap.add_argument('--out-dir', default=None, help='輸出目錄')
    args = ap.parse_args()

    wb = load_workbook(args.pjt_file, data_only=True)
    sheets = wb.sheetnames
    date_sheets = [s for s in sheets if re.match(r'^\d{2}-\d{2}$', s.strip())]
    cur_name = args.current or (date_sheets[0] if date_sheets else sheets[0])
    print(f'本週 sheet: {cur_name}')

    out_dir = args.out_dir or os.path.dirname(args.pjt_file) or '.'
    os.makedirs(out_dir, exist_ok=True)
    prefix = f'PJT {cur_name}'

    ws = wb[cur_name]

    # 1) Q×產區 矩陣
    p = os.path.join(out_dir, f'{prefix} Q產區矩陣.jpg')
    render_q_region_matrix(ws, p)
    print(f'  {p}')

    # 2) Q3 / Q4 Buffer
    buffer_sheets = [s for s in sheets if 'buffer' in s.lower()]
    if buffer_sheets:
        bws = wb[buffer_sheets[0]]
        # Q3 Buffer 已從 email 移除，且新版 sheet 欄位位置已異動，跳過
        p3 = os.path.join(out_dir, f'{prefix} Q3 Buffer.jpg')
        try:
            render_q3_buffer(bws, p3)
            print(f'  {p3}')
        except Exception:
            print(f'  (skip) Q3 Buffer JPG 略過（sheet 結構異動）')
        p4 = os.path.join(out_dir, f'{prefix} Q4 Buffer.jpg')
        render_q4_buffer(bws, p4)
        print(f'  {p4}')
    else:
        print('  (warning) 找不到 buffer sheet，跳過 Buffer JPG')

    # 3) BY 部門
    p = os.path.join(out_dir, f'{prefix} BY部門.jpg')
    render_by_dept(ws, p)
    print(f'  {p}')

    # 4) BY 產區
    p = os.path.join(out_dir, f'{prefix} BY產區.jpg')
    render_by_region(ws, p)
    print(f'  {p}')

    print('5 張 JPG 全部完成')


if __name__ == '__main__':
    main()
