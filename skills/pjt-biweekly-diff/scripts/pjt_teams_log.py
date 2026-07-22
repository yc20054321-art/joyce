#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
pjt_teams_log.py — 將 Teams PJT大表群組更新訊息整理成追蹤 Excel

Claude 先用 Teams MCP 搜尋訊息，解析後寫入 data JSON，再呼叫本腳本產 Excel。

用法：
  python pjt_teams_log.py --json-file <messages.json> --week 06-30 \
    [--out-dir "C:/Users/JoyceChen/Desktop/Claude"]

data JSON 格式：
  {
    "week": "06-30",          # PJT 週次 (MM-DD)
    "messages": [
      {
        "time":   "2026-07-01 13:10",   # 訊息時間（Taipei）
        "sender": "Joanne Wu / 吳欣容", # 發訊者
        "dept":   "#32-solo",           # 部門（#22/#23/#27/#32-solo…）
        "region": "印尼",               # 産区（空字串=不明）
        "style":  "60327H092A",         # 款號
        "month":  "10月",              # 出口月
        "qty":    99,                   # 標打量（正=加，負=減）
        "reason": "昨日+123→今日+99",  # 動作原因（精簡）
        "raw":    "SOLO 60327H092A …"  # 原文摘要
      },
      ...
    ]
  }

部門列表（固定，與 PJT 大表 BY部門截圖一致）：
  #22 / #22 solo / #23 / #24 / #27 / #kids /
  #32-non solo / #32-solo / #34 / #33 / #37 NON solo / #37 Solo
"""

import json
import sys
import argparse
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── 色碼 ──────────────────────────────────────────────────────────────────────
HDR   = PatternFill("solid", fgColor="305496")  # 深藍 header
POS   = PatternFill("solid", fgColor="DEEAF1")  # 淡藍 加量
NEG   = PatternFill("solid", fgColor="FCE4D6")  # 淡橘 減量
ZERO  = PatternFill("solid", fgColor="F2F2F2")  # 灰 持平
DEPT  = PatternFill("solid", fgColor="D6DCE4")  # 部門列底色
TOTAL = PatternFill("solid", fgColor="BDD7EE")  # 合計列
WH    = Font(color="FFFFFF", bold=True)
RED   = Font(color="C00000", bold=True)
BLU   = Font(color="0070C0", bold=True)
BLD   = Font(bold=True)
thin  = Side(style="thin", color="BFBFBF")
BD    = Border(left=thin, right=thin, top=thin, bottom=thin)

# 固定部門順序（與 PJT 大表 BY部門截圖一致）
DEPT_ORDER = [
    "#22", "#22 solo", "#23", "#24", "#27", "#kids",
    "#32-non solo", "#32-solo", "#34", "#33", "#37 NON solo", "#37 Solo"
]

# 月份 → 季度
def mo2q(month: str) -> str:
    try:
        n = int(str(month).replace("月", "").strip())
        if n in (4, 5, 6):  return "Q2"
        if n in (7, 8, 9):  return "Q3"
        if n in (10, 11, 12): return "Q4"
        if n in (1, 2, 3):  return "Q1"
    except ValueError:
        pass
    return "?"


# ── Cell 輔助 ─────────────────────────────────────────────────────────────────

def cell_hdr(ws, r, c, v, w=None):
    cell = ws.cell(r, c, v)
    cell.font = WH; cell.fill = HDR; cell.border = BD
    cell.alignment = Alignment(horizontal="center", vertical="center")
    if w:
        ws.column_dimensions[get_column_letter(c)].width = w
    return cell

def cell_put(ws, r, c, v, fill=None, font=None, align="left", wrap=False):
    cell = ws.cell(r, c, v)
    if fill:  cell.fill = fill
    if font:  cell.font = font
    cell.border = BD
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    return cell


# ── Sheet 1 — 異動明細 ────────────────────────────────────────────────────────

DETAIL_COLS = ["週次", "訊息時間", "發訊者", "部門", "産区",
               "款號", "出口月", "出口季", "標打量(±)", "增/減", "動作原因", "原文摘要"]
DETAIL_W    = [7, 17, 20, 12, 6, 16, 7, 7, 10, 6, 36, 50]


def build_detail(wb, week: str, messages: list):
    ws = wb.active
    ws.title = "異動明細"
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 20

    for c, (h, w) in enumerate(zip(DETAIL_COLS, DETAIL_W), 1):
        cell_hdr(ws, 1, c, h, w)

    for i, msg in enumerate(messages, 2):
        qty = msg.get("qty", 0)
        q   = mo2q(msg.get("month", ""))
        fill = POS if qty > 0 else (NEG if qty < 0 else ZERO)
        sign = "加量" if qty > 0 else ("減量" if qty < 0 else "持平")

        vals = [week, msg.get("time",""), msg.get("sender",""),
                msg.get("dept",""), msg.get("region",""), msg.get("style",""),
                msg.get("month",""), q, qty, sign,
                msg.get("reason",""), msg.get("raw","")]

        for c, val in enumerate(vals, 1):
            font_ = (BLU if qty > 0 else RED) if c in (9, 10) else None
            align = "center" if c in (9, 10) else ("left" if c != len(DETAIL_COLS) else "left")
            cell_put(ws, i, c, val, fill=fill, font=font_,
                     align=align, wrap=(c == len(DETAIL_COLS)))
        ws.row_dimensions[i].height = 28

    # 合計列
    last = len(messages) + 2
    cell_put(ws, last, 1, "合計", font=BLD)
    total = sum(m.get("qty", 0) for m in messages)
    tc = ws.cell(last, 9, total)
    tc.font = Font(bold=True, color="0070C0" if total > 0 else "C00000")
    tc.border = BD; tc.alignment = Alignment(horizontal="center")

    ws.freeze_panes = "A2"


# ── Sheet 2 — BY部門彙總 ──────────────────────────────────────────────────────

def build_summary(wb, week: str, messages: list):
    from collections import defaultdict
    ws = wb.create_sheet("BY部門彙總")
    ws.sheet_view.showGridLines = False

    # Header
    hdrs = ["部門", "2026-Q2", "2026-Q3", "2026-Q4", "2026 ALL", "說明（Teams異動訊息）"]
    ws.row_dimensions[1].height = 20
    for c, (h, w) in enumerate(zip(hdrs, [14, 10, 10, 10, 10, 72]), 1):
        cell_hdr(ws, 1, c, h, w)

    # 彙總：每部門 Q2/Q3/Q4 合計 + 說明清單
    agg   = defaultdict(lambda: {"Q2": 0, "Q3": 0, "Q4": 0})
    notes = defaultdict(list)
    for msg in messages:
        dept = msg.get("dept", "")
        qty  = msg.get("qty", 0)
        q    = mo2q(msg.get("month", ""))
        if q in ("Q2", "Q3", "Q4"):
            agg[dept][q] += qty
        # 說明文字：款號 + 月份 + ±量 + 原因
        style = msg.get("style", "")
        month = msg.get("month", "")
        note  = f"{style} {month} {'+' if qty >= 0 else ''}{qty}標打 ─ {msg.get('reason','')}"
        if note not in notes[dept]:
            notes[dept].append(note)

    # 寫入部門列
    r = 2
    tot = {"Q2": 0, "Q3": 0, "Q4": 0}
    for dept in DEPT_ORDER:
        d    = agg.get(dept, {"Q2": 0, "Q3": 0, "Q4": 0})
        q2, q3, q4 = d["Q2"], d["Q3"], d["Q4"]
        all_ = q2 + q3 + q4
        for k in ("Q2", "Q3", "Q4"):
            tot[k] += d[k]

        # 部門名稱
        dc = ws.cell(r, 1, dept)
        dc.fill = DEPT; dc.font = BLD; dc.border = BD
        dc.alignment = Alignment(horizontal="center", vertical="center")

        # Q2/Q3/Q4/ALL 數值
        for ci, val in enumerate([q2, q3, q4, all_], 2):
            display = val if val != 0 else ""
            c2 = cell_put(ws, r, ci, display,
                          fill=POS if val > 0 else (NEG if val < 0 else ZERO),
                          font=BLU if val > 0 else (RED if val < 0 else None),
                          align="center")

        # 說明欄（多行）
        note_txt = "\n".join(notes.get(dept, []))
        nc = ws.cell(r, 6, note_txt)
        nc.border = BD
        nc.alignment = Alignment(vertical="top", wrap_text=True)
        if note_txt:
            nc.fill = POS if all_ > 0 else (NEG if all_ < 0 else ZERO)

        # 行高依說明行數
        line_count = max(1, len(notes.get(dept, [""])))
        ws.row_dimensions[r].height = max(30, 16 * line_count)
        r += 1

    # 合計列
    tot_all = tot["Q2"] + tot["Q3"] + tot["Q4"]
    tc = ws.cell(r, 1, "total")
    tc.fill = TOTAL; tc.font = BLD; tc.border = BD
    tc.alignment = Alignment(horizontal="center", vertical="center")
    for ci, val in enumerate([tot["Q2"], tot["Q3"], tot["Q4"], tot_all], 2):
        c3 = ws.cell(r, ci, val if val != 0 else "")
        c3.font = Font(bold=True,
                       color="0070C0" if val > 0 else ("C00000" if val < 0 else "000000"))
        c3.fill = TOTAL; c3.border = BD
        c3.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "B2"

    # 備註
    ws.cell(r + 2, 1, "※ 資料來源：Teams PJT大表群組  |  Q2=4-6月  Q3=7-9月  Q4=10-12月").font = \
        Font(italic=True, color="7F7F7F", size=9)
    ws.cell(r + 3, 1, "※ 標打量：正數(藍)=加量  負數(紅)=減量  括號=減量").font = \
        Font(italic=True, color="7F7F7F", size=9)


# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    ap = argparse.ArgumentParser(description="PJT Teams 異動記錄 Excel 產生器")
    ap.add_argument("--json-file", required=True,
                    help="包含 parsed messages 的 JSON 檔案路徑")
    ap.add_argument("--week", default="",
                    help="PJT 週次 MM-DD（預設從 JSON 的 week 欄位讀取）")
    ap.add_argument("--out-dir", default=r"C:\Users\JoyceChen\Desktop\Claude",
                    help="輸出目錄")
    args = ap.parse_args()

    json_path = Path(args.json_file)
    if not json_path.exists():
        sys.exit(f"[ERROR] 找不到 JSON 檔：{json_path}")

    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    week     = args.week or data.get("week", "00-00")
    messages = data.get("messages", [])
    if not messages:
        sys.exit("[ERROR] JSON 內沒有 messages 資料。")

    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / f"PJT_Teams異動記錄_{week}.xlsx"

    wb = Workbook()
    build_detail(wb, week, messages)
    build_summary(wb, week, messages)
    wb["異動明細"].sheet_properties.tabColor   = "305496"
    wb["BY部門彙總"].sheet_properties.tabColor = "4472C4"
    wb.save(str(out_path))

    total = sum(m.get("qty", 0) for m in messages)
    print(f"[OK] {out_path}")
    print(f"     訊息數：{len(messages)} 筆")
    print(f"     淨差：{total:+,d} 標打")


if __name__ == "__main__":
    main()
