#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
pjt_email.py — PJT 二週差異 Outlook 草稿產生器（v4 版型）

從 pjt_diff.py 輸出的 RECAP xlsx + 原始 PJT xlsx 讀取數據，
自動產出符合 v4 版型的 HTML 檔 + PowerShell Outlook 草稿腳本。

用法：
  python pjt_email.py <recap.xlsx> ^
    --pjt-file <pjt.xlsx> ^
    --current 06-09 ^
    [--img-dir "C:/Users/JoyceChen/Desktop/Claude"] ^
    [--recipients "config/recipients.json"] ^
    [--out-dir "C:/Users/JoyceChen/Desktop/Claude"] ^
    [--run]

v4 版型結構（確立於 2026-06-10）：
  [Arthur] Dear Arthur
  ★ 全年 PJT 暫定 XXX 萬打 | 二週淨差 ±X.X 萬打 => 結論
  今日較大調動（有重大異動時顯示）
  [Q産区矩陣圖]               ← 在 BY出口期說明 之前
  【BY 出口期說明】Q3/Q4/2027-Q1 bullets（Q2 不顯示）
  【Q3/Q4 Buffer】            ← Arthur 版專用
  【BY 部門】/ 【BY 産區】

安全紅線：
  草稿 .Save()，不 .Send()
  收件人定義在 config/recipients.json，不動腳本
"""

import re
import sys
import json
import argparse
import subprocess
from pathlib import Path

# ── openpyxl 壞圖 patch（PJT 大表含 xl/drawings/NULL 無效參照）─────────────
try:
    import openpyxl.reader.excel as _xlreader
    _orig_fi = _xlreader.find_images
    def _safe_fi(archive, path):
        try:
            return _orig_fi(archive, path)
        except Exception:
            return [], []
    _xlreader.find_images = _safe_fi
except Exception:
    pass

from openpyxl import load_workbook

# ── 色碼常數 ──────────────────────────────────────────────────────────────────
R = "#C00000"    # 減量紅
B = "#0070C0"    # 加量藍
G = "#00B050"    # Buffer 消化綠
H = "#305496"    # 段落標題深藍
FONT = "font-family:'微軟正黑體',Calibri,sans-serif; font-size:11pt;"

REGION_MAP = {'印尼': 'IND', '柬埔寨': 'CAB', '中國': 'CHN', '北越': '北越'}
THRESH = 0.05    # 萬打；低於此值視為持平


# ════════════════════════════════════════════════════════════════════════════════
#  格式化工具
# ════════════════════════════════════════════════════════════════════════════════

def fmt_wan(val: float, sign: bool = True, deco: bool = True, bold: bool = False) -> str:
    """萬打數值 → 帶色碼 HTML span（正值藍/負值紅；deco=底線；bold=粗體）。"""
    if val is None:
        return ''
    color = R if val < -THRESH else (B if val > THRESH else '')
    s = f"{'+' if (sign and val > 0) else ''}{val:.1f}萬打"
    inner = f"<b>{s}</b>" if bold else s
    if color:
        deco_css = " text-decoration:underline;" if deco else ""
        return f"<span style='color:{color};{deco_css}'>{inner}</span>"
    return f"<b>{s}</b>" if bold else s


def extract_reason(annotation: str) -> str:
    """取 annotation 中第一個【…】（多餘空白 normalize）。"""
    if not annotation:
        return ''
    m = re.search(r'【[^】]+】', str(annotation))
    if not m:
        return ''
    return re.sub(r'\s+', ' ', m.group(0))


def wan_conclusion(val: float) -> str:
    if val > THRESH:
        return '加量'
    if val < -THRESH:
        return '減量'
    return '持平'


# ════════════════════════════════════════════════════════════════════════════════
#  讀取 RECAP xlsx
# ════════════════════════════════════════════════════════════════════════════════

def parse_diff_sheet(ws) -> dict:
    """
    解析「異動說明清單」sheet。

    回傳格式：
    {
      'Q2': {
        'net': float,        # 萬打（Q 小計列）
        'conclusion': str,
        'regions': {
          'IND': {
            'net': float, 'conclusion': str, 'annotation': str,
            'depts': [{'name': str, 'diff': float, 'annotation': str}, ...]
          }, ...
        }
      },
      'Q3': {...},
      'Q4': {...}
    }
    """
    result = {}
    current_q = None
    current_region = None

    for row in ws.iter_rows(values_only=True):
        if not any(v is not None for v in row[:6]):
            continue
        cols = [(row[i] if i < len(row) else None) for i in range(6)]
        q_col, reg_col, dept_col, diff_col, conc_col, ann_col = cols
        diff_val = float(diff_col) if isinstance(diff_col, (int, float)) else 0.0
        q_str = str(q_col).strip() if q_col else ''

        # ── Q 小計行（「Q2 小計」「Q3 小計」「Q4 小計」）
        if q_str.endswith('小計'):
            q_key = q_str.split()[0]
            if q_key in ('Q2', 'Q3', 'Q4', 'Q1'):
                current_q = q_key
                current_region = None
                result[q_key] = {
                    'net': diff_val,
                    'conclusion': str(conc_col or '持平'),
                    'regions': {}
                }
            continue

        if current_q is None:
            continue

        # ── Q × 産区 ALL 行（同一季）
        if q_str == current_q and reg_col and dept_col == 'ALL':
            region_short = REGION_MAP.get(str(reg_col).strip(), str(reg_col).strip())
            current_region = region_short
            result[current_q]['regions'][region_short] = {
                'net': diff_val,
                'conclusion': str(conc_col or '持平'),
                'annotation': str(ann_col or ''),
                'depts': []
            }
            continue

        # ── 跨季 ALL 行（其他季的地區行，通常不會發生，安全保護）
        if q_str and q_str != current_q and dept_col == 'ALL':
            current_region = None
            continue

        # ── 部門子行（q_col=None, reg_col=None, dept_col='#XX'）
        if q_col is None and reg_col is None and dept_col and current_q and current_region:
            rg = result[current_q]['regions']
            if current_region in rg:
                rg[current_region]['depts'].append({
                    'name': re.sub(r'\s+', ' ', str(dept_col).strip()),
                    'diff': diff_val,
                    'annotation': str(ann_col or '')
                })

    return result


def parse_qx_sheet(ws) -> tuple:
    """
    解析「Q×産区二週差異」sheet。
    回傳 (net_wan: float, conclusion: str)：産区合計列的全年淨差。
    """
    headers = None
    for row in ws.iter_rows(values_only=True):
        first = str(row[0]).strip() if row[0] else ''
        if first == '季度':
            headers = [str(h) if h else '' for h in row]
            continue
        if headers and ('産' in first or '產' in first) and '合計' in first:  # 産區合計/產區合計
            try:
                idx = next(i for i, h in enumerate(headers) if '合計' in h and '季' in h)
                net = float(row[idx]) if row[idx] is not None else 0.0
            except (StopIteration, TypeError, ValueError):
                net = 0.0
            try:
                cidx = next(i for i, h in enumerate(headers) if '結論' in h)
                conc = str(row[cidx] or wan_conclusion(net))
            except StopIteration:
                conc = wan_conclusion(net)
            return net, conc
    return 0.0, '持平'


# ════════════════════════════════════════════════════════════════════════════════
#  讀取原始 PJT xlsx
# ════════════════════════════════════════════════════════════════════════════════

def read_buffer(pjt_path: str) -> dict:
    """
    讀 Q3+Q4 buffer工作表1 sheet。
    Q4 row31: col12=本週total, col18=差異
    Q3 row19: col30=本週total, col36=差異
    單位：標打；÷10000 = 萬打。
    """
    wb = load_workbook(pjt_path, data_only=True)
    buf_candidates = [s for s in wb.sheetnames if 'buffer' in s.lower()]
    if not buf_candidates:
        print("[WARN] 找不到 buffer sheet，Buffer 數字略過。", file=sys.stderr)
        return {}
    ws = wb[buf_candidates[0]]

    def sv(r, c):
        v = ws.cell(r, c).value
        return float(v) if isinstance(v, (int, float)) else 0.0

    return {
        'q3_total_wan': round(sv(19, 30) / 10000, 1),
        'q3_diff_wan':  round(sv(19, 36) / 10000, 1),  # 萬打（同 Q4）
        'q4_total_wan': round(sv(31, 12) / 10000, 1),
        'q4_diff_wan':  round(sv(31, 18) / 10000, 1),
    }


def read_annual_pjt(pjt_path: str, current: str) -> tuple:
    """
    從 current sheet「調整後季」列讀 PJT 暫定。
    回傳 (annual_2026_wan, annual_q1_wan)：
      annual_2026 = 調整後季 Q2(col19) + Q3(col22) + sum(rows18-22, cols25-27 Oct-Dec Q4)
      annual_q1   = sum(rows18-22, cols28-30 Jan-Mar 2027)
    """
    wb = load_workbook(pjt_path, data_only=True)
    if current not in wb.sheetnames:
        print(f"[WARN] 找不到 sheet '{current}'，全年 PJT 無法讀取。", file=sys.stderr)
        return 0.0, 0.0
    ws = wb[current]

    target_row = None
    for r in range(1, 80):
        v = ws.cell(r, 18).value
        if v and '調整後季' in str(v):
            target_row = r
            break
    if not target_row:
        print("[WARN] 找不到「調整後季」列。", file=sys.stderr)
        return 0.0, 0.0

    def sv(r, c):
        v = ws.cell(r, c).value
        return float(v) if isinstance(v, (int, float)) else 0.0

    # 調整後季 row：季度合計放在各季第一個月欄（Apr=col19, Jul=col22, Oct=col25, Jan27=col28）
    q2_adj = sv(target_row, 19)
    q3_adj = sv(target_row, 22)
    q4_adj = sv(target_row, 25)
    q1_adj = sv(target_row, 28)

    annual_2026 = round((q2_adj + q3_adj + q4_adj) / 10000, 1)
    annual_q1   = round(q1_adj / 10000, 1)
    return annual_2026, annual_q1


# ════════════════════════════════════════════════════════════════════════════════
#  HTML 段落產生器
# ════════════════════════════════════════════════════════════════════════════════

def has_major_move(diff_data: dict, threshold: float = 0.5) -> bool:
    """任一産区單季淨差 >= threshold 萬打 → 顯示「今日較大調動」標題。"""
    for q in ('Q3', 'Q4', 'Q1'):
        for rdata in diff_data.get(q, {}).get('regions', {}).values():
            if abs(rdata['net']) >= threshold:
                return True
    return False


def dept_item_html(dept: dict) -> str:
    """單一部門細項 HTML：#款號【reason】 ±X.X萬打（色碼+底線）。"""
    name = dept['name']
    diff = dept['diff']
    reason = extract_reason(dept['annotation'])
    amount = fmt_wan(diff, sign=True, deco=True)
    return f"{name}{reason}&nbsp;{amount}"


def build_q_html(q: str, q_data: dict) -> str:
    """
    一個季度的 bullet HTML。
    Q2：單行（精簡）；Q3/Q4/Q1：主行 + 縮排地區子行。
    """
    net = q_data['net']
    regions = q_data.get('regions', {})

    # Q1 顯示為 "2027-Q1"
    q_display = '2027-Q1' if q == 'Q1' else q

    # Q 主行
    net_c = R if net < -THRESH else (B if net > THRESH else '')
    if abs(net) > THRESH:
        net_s = f"{'+' if net > 0 else ''}{net:.1f} 萬打"
        if net_c:
            net_html = f"<span style='color:{net_c};'><b>{net_s}</b></span>"
        else:
            net_html = f"<b>{net_s}</b>"
        q_label = f"<b>&#9658; {q_display}&nbsp;{net_html}：</b>"
    else:
        q_label = f"<b>&#9658; {q_display}&nbsp;持平</b>："

    # ── Q2 單行精簡（Q2 已不在標準出口期範圍，保留以備需要）
    if q == 'Q2':
        parts = []
        for reg_short, rd in regions.items():
            active = [d for d in rd.get('depts', []) if abs(d['diff']) > THRESH]
            if not active and abs(rd['net']) <= THRESH:
                continue
            items = '；'.join(dept_item_html(d) for d in active)
            rnet_html = fmt_wan(rd['net'], bold=True) if abs(rd['net']) > THRESH else '<b>持平</b>'
            if items:
                parts.append(f"{reg_short}&nbsp;{rnet_html}（{items}）")
            else:
                parts.append(f"{reg_short}&nbsp;{rnet_html}")
        summary = '；'.join(parts) if parts else '各產區持平'
        return q_label + summary

    # ── Q3 / Q4 多行（地區縮排子行）
    indent = '&nbsp;&nbsp;&nbsp;'
    lines = [q_label]
    for reg_short, rd in regions.items():
        r_net = rd['net']
        active = [d for d in rd.get('depts', []) if abs(d['diff']) > THRESH]
        if abs(r_net) <= THRESH and not active:
            continue   # 純持平且無內部異動 → 略過

        # 地區標頭
        if abs(r_net) > THRESH:
            rc = R if r_net < -THRESH else B
            r_head = f"{indent}{reg_short}&nbsp;<b><span style='color:{rc};'>{'+' if r_net > 0 else ''}{r_net:.1f}萬打</span></b>："
        else:
            r_head = f"{indent}{reg_short}&nbsp;<b>持平</b>："

        # 部門細項分類
        add_items = [dept_item_html(d) for d in active if d['diff'] > THRESH]
        sub_items = [dept_item_html(d) for d in active if d['diff'] < -THRESH]
        detail_parts = []
        if add_items:
            detail_parts.append(f"加量&nbsp;{'，'.join(add_items)}")
        if sub_items:
            detail_parts.append(f"減量&nbsp;{'，'.join(sub_items)}")
        lines.append(r_head + '；'.join(detail_parts))

    return '<br>\n'.join(lines)


def build_byq_section(diff_data: dict) -> str:
    """完整【BY 出口期說明】段落 HTML（Q3/Q4/2027-Q1，不含 Q2）。"""
    blocks = []
    labels = {'Q3': 'Q3', 'Q4': 'Q4', 'Q1': '2027-Q1'}
    for q, q_disp in labels.items():
        if q in diff_data:
            blocks.append(build_q_html(q, diff_data[q]))
        else:
            blocks.append(f"<b>&#9658; {q_disp}&nbsp;持平</b>")
    return (
        f"<div style='{FONT}'>\n"
        f"<p style='font-size:12pt; font-weight:bold; color:{H}; margin:0 0 6px 0;'>"
        f"【BY 出口期說明】</p>\n"
        f"<p style='margin:0; line-height:1.9;'>\n"
        + '<br>\n'.join(blocks)
        + "\n</p>\n</div>\n"
    )


def _img_b64(path: str, max_width: int) -> str:
    """Return <img> HTML with base64-encoded JPEG, or empty string if file missing."""
    import base64
    p = Path(path)
    if not p.exists():
        return ''
    b64 = base64.b64encode(p.read_bytes()).decode()
    return f"<img src='data:image/jpeg;base64,{b64}' style='max-width:{max_width}px;' />"


def build_buffer_section(buf: dict, q4buf_img_path: str = '') -> str:
    """Q4 Buffer 段落（Arthur 版專用；Q3 Buffer 已移除）。"""
    if not buf:
        return ''

    # Q4 Buffer（消化 → 綠；增加 → 紅）
    q4t = f"{buf['q4_total_wan']:.1f}"
    q4d = buf['q4_diff_wan']
    q4conc = '大幅消化' if q4d < -0.3 else ('消化' if q4d < -THRESH else ('增加' if q4d > THRESH else '持平'))
    q4ds = f"{'+' if q4d > 0 else ''}{q4d:.1f}"
    if q4d < -THRESH:
        q4da = f"<b><span style='color:{G}; text-decoration:underline;'>{q4ds}&nbsp;萬打</span></b>"
    elif q4d > THRESH:
        q4da = f"<b><span style='color:{R};'>{q4ds}&nbsp;萬打</span></b>"
    else:
        q4da = f"<b>{q4ds}&nbsp;萬打</b>"

    img_html = _img_b64(q4buf_img_path, 700)
    return (
        f"<div style='{FONT} margin-top:16px;'>\n"
        f"<p style='margin:0 0 4px 0;'><b>【Q4 Buffer】</b>&nbsp; 共計 <b>{q4t}&nbsp;萬打</b>，"
        f"本週{q4conc}&nbsp;{q4da}</p>\n"
        f"<p style='margin:0;'>{img_html}</p>\n"
        f"</div>\n"
    )


def build_html(version: str, annual_2026: float, annual_q1: float,
               net_wan: float, net_2026: float, net_conc: str,
               diff_data: dict, buf: dict, current: str,
               matrix_img_path: str = '',
               q4buf_img_path: str = '',
               dept_img_path: str = '',
               region_img_path: str = '',
               q1_note: str = '') -> str:
    """
    完整 email HTML。所有圖片以 base64 data URI 內嵌，不走 CID 附件。
    version: 'arthur' | 'dept'
    雙 ★ header：第一行 2026 全年、第二行 2027-Q1
    """
    q1_net = net_wan - net_2026   # Q1 單季淨差

    # 2026 net span
    net_2026_c = R if net_2026 < -THRESH else (B if net_2026 > THRESH else '')
    net_2026_s = f"{'+' if net_2026 > 0 else ''}{net_2026:.1f}"
    net_2026_html = (f"<span style='color:{net_2026_c};'>{net_2026_s}&nbsp;萬打</span>"
                     if net_2026_c else f"{net_2026_s}&nbsp;萬打")

    # Q1 net span
    net_q1_c = R if q1_net < -THRESH else (B if q1_net > THRESH else '')
    net_q1_s = f"{'+' if q1_net > 0 else ''}{q1_net:.1f}"
    net_q1_html = (f"<span style='color:{net_q1_c};'>{net_q1_s}&nbsp;萬打</span>"
                   if net_q1_c else f"{net_q1_s}&nbsp;萬打")
    q1_action = '淨增' if q1_net > THRESH else ('淨減' if q1_net < -THRESH else '持平')
    q1_note_html = f"&nbsp;&nbsp;&nbsp;{q1_note}" if q1_note else ''

    line_2026 = (
        f"<p style='{FONT} margin:0 0 4px 0;'>"
        f"<b>★2026 全年 PJT 暫定 {annual_2026}&nbsp;萬打 &nbsp;|&nbsp; "
        f"二週淨差 {net_2026_html}</b></p>\n"
    )
    line_q1 = (
        f"<p style='{FONT} margin:0 0 10px 0;'>"
        f"<b>★2027 Q1 暫定 {annual_q1}&nbsp;萬打 &nbsp;|&nbsp; "
        f"二週{q1_action} {net_q1_html}{q1_note_html}</b></p>\n"
    )

    if version == 'arthur':
        major = (
            f"<p style='{FONT} margin:0 0 14px 0;'><b>今日較大調動</b></p>\n"
            if has_major_move(diff_data) else ''
        )
        header = (
            f"<p style='{FONT} margin:0 0 6px 0;'><b>Dear Arthur</b></p>\n"
            f"{line_2026}{line_q1}{major}"
        )
    else:
        header = f"{line_2026}{line_q1}"

    # 所有圖片 base64 內嵌（不走 CID，不在附件列顯示）
    matrix = ''
    if matrix_img_path:
        img = _img_b64(matrix_img_path, 700)
        if img:
            matrix = f"<div style='margin:0 0 14px 0;'>{img}</div>\n"

    # v4 結構：矩陣圖(base64) → BY出口期 → Buffer(Arthur) → BY部門 → BY産区
    byq = build_byq_section(diff_data)
    buf_html = build_buffer_section(buf, q4buf_img_path) if version == 'arthur' else ''
    dept_img = _img_b64(dept_img_path, 680) if dept_img_path else ''
    region_img = _img_b64(region_img_path, 780) if region_img_path else ''
    by_dept = (
        f"<div style='{FONT} margin-top:16px;'>"
        f"<p style='margin:0 0 4px 0;'><b>【BY 部門】</b></p>"
        f"<p style='margin:0;'>{dept_img}</p>"
        f"</div>\n"
    )
    by_region = (
        f"<div style='{FONT} margin-top:16px;'>"
        f"<p style='margin:0 0 4px 0;'><b>【BY 産區】</b></p>"
        f"<p style='margin:0;'>{region_img}</p>"
        f"</div>\n"
    )
    sig = (
        f"<div style='{FONT} margin-top:18px; padding-top:10px; "
        f"border-top:1px solid #BFBFBF; font-size:9pt; color:#595959;'>"
        f"Joyce Chen &nbsp;|&nbsp; 業務六處 &nbsp;|&nbsp; Makalot Industrial Co., Ltd."
        f"</div>\n"
    )

    body = header + matrix + byq + buf_html + by_dept + by_region + sig
    return f"<html><body style='{FONT}'>\n{body}</body></html>\n"


# ════════════════════════════════════════════════════════════════════════════════
#  PS1 產生器
# ════════════════════════════════════════════════════════════════════════════════

def build_ps1(current: str,
              html_arthur_path: str, html_dept_path: str,
              recipients: dict, out_ps1: str) -> None:
    """
    產生 PowerShell Outlook 雙草稿腳本。
    圖片已 base64 內嵌於 HTML，PS1 不需要處理附件。
    """
    month = str(int(current.split('-')[0]))
    day   = str(int(current.split('-')[1]))
    subject = f"RE: 2026+2027-Q1 Weekly PJT Review -{month}/{day}"

    arthur_to = (
        recipients.get('arthur', {}).get('to_resolved')
        or recipients.get('arthur', {}).get('to', '')
    )
    arthur_cc_list = recipients.get('arthur', {}).get('cc', [])
    arthur_cc = '; '.join(arthur_cc_list) if isinstance(arthur_cc_list, list) else str(arthur_cc_list)

    dept_to_list = recipients.get('dept', {}).get('to', [])
    dept_to = '; '.join(dept_to_list) if isinstance(dept_to_list, list) else str(dept_to_list)

    ps1 = f"""# PJT {current} Outlook 雙草稿 — Auto-generated by pjt_email.py
# .Save() only — 確認內容後請手動送出，切勿直接 Send
# 圖片已 base64 內嵌於 HTML，無需處理附件
# ─────────────────────────────────────────────────────────────
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$subject   = "{subject}"
$arthurTo  = "{arthur_to}"
$arthurCC  = "{arthur_cc}"
$deptTo    = "{dept_to}"

# HTML 從外部檔讀入（圖片已內嵌，避免 PS1 轉義問題）
$htmlArthur = Get-Content "{html_arthur_path}" -Encoding UTF8 -Raw
$htmlDept   = Get-Content "{html_dept_path}"   -Encoding UTF8 -Raw

# ── 建立草稿函數
function New-PjtDraft {{
    param(
        [object]$ol,
        [string]$To,
        [string]$CC,
        [string]$Subj,
        [string]$Html,
        [string]$Category
    )
    $mail = $ol.CreateItem(0)
    $mail.Subject    = $Subj
    $mail.To         = $To
    if ($CC) {{ $mail.CC = $CC }}
    $mail.BodyFormat = 2  # olFormatHTML
    $mail.HTMLBody   = $Html
    $mail.Categories = $Category
    $mail.Save()
    return $mail
}}

# ── 主程式
try {{
    $ol = New-Object -ComObject Outlook.Application

    Write-Host "建立 PJT-Arthur 草稿..."
    $m1 = New-PjtDraft $ol $arthurTo $arthurCC $subject $htmlArthur "PJT-Arthur"
    Write-Host "  ✓ 已存：$($m1.Subject)"

    Write-Host "建立 PJT-Dept 草稿..."
    $m2 = New-PjtDraft $ol $deptTo "" $subject $htmlDept "PJT-Dept"
    Write-Host "  ✓ 已存：$($m2.Subject)"

    Write-Host ""
    Write-Host "完成 — 草稿匣已有 PJT-Arthur / PJT-Dept，確認後手動送出。"
}} catch {{
    Write-Host "ERROR: $_"
    exit 1
}}
"""

    with open(out_ps1, 'w', encoding='utf-8') as f:
        f.write(ps1)
    print(f"[OK] PS1 已寫入：{out_ps1}")


# ════════════════════════════════════════════════════════════════════════════════
#  主程式
# ════════════════════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser(description='PJT 二週差異 Outlook 草稿產生器 v4')
    ap.add_argument('recap_xlsx',
                    help='pjt_diff.py 輸出的 RECAP xlsx')
    ap.add_argument('--pjt-file', required=True,
                    help='原始 PJT xlsx（讀 Buffer + 全年暫定）')
    ap.add_argument('--current', default='',
                    help='本週日期 MM-DD（預設從 recap 檔名推斷）')
    ap.add_argument('--img-dir', default='',
                    help='JPG 目錄（預設與 recap 同目錄）')
    ap.add_argument('--recipients', default='',
                    help='recipients.json 路徑（預設 skill config/）')
    ap.add_argument('--out-dir', default='',
                    help='輸出目錄（預設與 recap 同目錄）')
    ap.add_argument('--q1-note', default='',
                    help='2027-Q1 行尾附加說明（選填，例如「各部門依27SS進度第一波調動」）')
    ap.add_argument('--run', action='store_true',
                    help='產出後立即執行 PS1 建立 Outlook 草稿')
    args = ap.parse_args()

    recap_path = Path(args.recap_xlsx)
    if not recap_path.exists():
        sys.exit(f"[ERROR] RECAP 檔不存在：{recap_path}")
    pjt_path = Path(args.pjt_file)
    if not pjt_path.exists():
        sys.exit(f"[ERROR] PJT 檔不存在：{pjt_path}")

    # 推斷日期
    current = args.current
    if not current:
        m = re.search(r'(\d{2}-\d{2})', recap_path.stem)
        current = m.group(1) if m else '00-00'
    print(f"[INFO] 本週日期：{current}")

    # 目錄
    img_dir = args.img_dir or str(recap_path.parent)
    out_dir = Path(args.out_dir or recap_path.parent)
    out_dir.mkdir(parents=True, exist_ok=True)

    # 收件人
    skill_root = Path(__file__).parent.parent  # .../pjt-biweekly-diff/
    default_rec = skill_root / 'config' / 'recipients.json'
    rec_path = Path(args.recipients) if args.recipients else default_rec
    if rec_path.exists():
        with open(rec_path, encoding='utf-8') as f:
            recipients = json.load(f)
        print(f"[INFO] 收件人：{rec_path}")
    else:
        print(f"[WARN] recipients.json 不存在，收件人欄位留空。", file=sys.stderr)
        recipients = {'arthur': {'to': '', 'cc': []}, 'dept': {'to': []}}

    # 讀 RECAP
    print("[INFO] 讀 RECAP xlsx...")
    wb = load_workbook(recap_path, data_only=True)

    diff_sheet = next((s for s in wb.sheetnames if '異動說明' in s), None)
    qx_sheet   = next((s for s in wb.sheetnames if '二週差異' in s and '部門' not in s), None)
    if not diff_sheet:
        sys.exit("[ERROR] 找不到「異動說明清單」sheet。")
    print(f"  差異說明 sheet：{diff_sheet}")
    print(f"  Q×産区 sheet  ：{qx_sheet}")

    diff_data = parse_diff_sheet(wb[diff_sheet])
    if qx_sheet:
        net_wan, net_conc = parse_qx_sheet(wb[qx_sheet])
    else:
        net_wan = sum(q.get('net', 0.0) for q in diff_data.values())
        net_conc = wan_conclusion(net_wan)

    print(f"[INFO] 二週淨差：{net_wan:+.1f} 萬打（{net_conc}）")

    # 讀 Buffer + 全年 PJT
    print("[INFO] 讀 Buffer...")
    buf = read_buffer(str(pjt_path))
    if buf:
        print(f"  Q3 Buffer：{buf['q3_total_wan']:.1f} 萬打，差異 {buf['q3_diff_wan']:+.1f} 萬打")
        print(f"  Q4 Buffer：{buf['q4_total_wan']:.1f} 萬打，差異 {buf['q4_diff_wan']:+.1f} 萬打")

    print("[INFO] 讀全年 PJT 暫定...")
    annual_2026, annual_q1 = read_annual_pjt(str(pjt_path), current)
    print(f"  2026 全年 PJT 暫定：{annual_2026} 萬打")
    print(f"  2027 Q1 暫定：{annual_q1} 萬打")

    # 2026 淨差 = 全部 - Q1；Q1 淨差單獨
    q1_net = diff_data.get('Q1', {}).get('net', 0.0)
    net_2026 = round(net_wan - q1_net, 1)

    # 產生 HTML（所有圖片 base64 內嵌）
    print("[INFO] 產生 HTML...")
    img_p = Path(img_dir)
    matrix_jpg  = str(img_p / f"PJT {current} Q產區矩陣.jpg")
    q4buf_jpg   = str(img_p / f"PJT {current} Q4 Buffer.jpg")
    dept_jpg    = str(img_p / f"PJT {current} BY部門.jpg")
    region_jpg  = str(img_p / f"PJT {current} BY產區.jpg")
    q1_note     = getattr(args, 'q1_note', '') or ''

    html_a = build_html('arthur', annual_2026, annual_q1, net_wan, net_2026, net_conc,
                        diff_data, buf, current,
                        matrix_img_path=matrix_jpg,
                        q4buf_img_path=q4buf_jpg,
                        dept_img_path=dept_jpg,
                        region_img_path=region_jpg,
                        q1_note=q1_note)
    html_d = build_html('dept',   annual_2026, annual_q1, net_wan, net_2026, net_conc,
                        diff_data, {},  current,
                        matrix_img_path=matrix_jpg,
                        dept_img_path=dept_jpg,
                        region_img_path=region_jpg,
                        q1_note=q1_note)

    path_a = out_dir / f"pjt_email_arthur_{current}.html"
    path_d = out_dir / f"pjt_email_dept_{current}.html"
    path_a.write_text(html_a, encoding='utf-8')
    path_d.write_text(html_d, encoding='utf-8')
    print(f"[OK] HTML 已寫入：")
    print(f"     Arthur：{path_a}")
    print(f"     Dept  ：{path_d}")

    # 產生 PS1（圖片已內嵌於 HTML，無附件邏輯）
    out_ps1 = out_dir / f"build_pjt_draft_{current}.ps1"
    build_ps1(
        current=current,
        html_arthur_path=str(path_a),
        html_dept_path=str(path_d),
        recipients=recipients,
        out_ps1=str(out_ps1),
    )

    # 執行 PS1（--run）
    if args.run:
        print(f"[INFO] 執行 PS1：{out_ps1}")
        cmd = f"Get-Content '{out_ps1}' -Encoding UTF8 | Out-String | Invoke-Expression"
        r = subprocess.run(
            ['powershell', '-NonInteractive', '-Command', cmd],
            capture_output=True, text=True
        )
        if r.stdout:
            print(r.stdout)
        if r.stderr:
            print(f"[WARN] {r.stderr}", file=sys.stderr)
        if r.returncode != 0:
            sys.exit(f"[ERROR] PS1 執行失敗（exit {r.returncode}）")
    else:
        print(f"\n[NEXT] 確認 HTML 後執行：")
        print(f"  Get-Content '{out_ps1}' -Encoding UTF8 | Out-String | Invoke-Expression")

    print("\n[DONE]")


if __name__ == '__main__':
    main()
