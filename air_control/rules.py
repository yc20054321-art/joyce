"""大貨主副料空/陸運控管表 —— 反黃區判斷規則。

規則來源：2026-08 大貨相關主副料空運快遞費用 mail thread
  檢視條件 ① 成衣量 > 10,000 件
           ② 決策 due day = 空/陸運 ETC 前兩週
  決策人員 成衣量 > 1 萬件 → 產區；其餘 → 主採 / 大貨 MD
  報價標準 $0.24 / 實打，超出即列入超標追蹤
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, timedelta

# ---- 可調參數（改這裡就好，不要動下面的邏輯）------------------------------
QTY_THRESHOLD = 10_000          # 成衣件數門檻（件）
DECISION_LEAD_DAYS = 14         # 決策 due day = ETC 前幾天
QUOTED_RATE = 0.24              # 報價已含的標準空運單價（USD / 實打）
URGENT_WINDOW_DAYS = 3          # 決策日剩幾天內視為「急」

OWNER_LARGE = "產區"
OWNER_NORMAL = "主採/大貨MD"

CONFIRM_BY_MODE = {
    "空運": "是否需全量走AIR?",
    "陸運": "是否需全量走卡車?",
    "海運": "",
}


@dataclass
class Row:
    """一筆料號（主布或配布）的控管資料。"""

    產區: str = ""
    DIV: str = ""
    STYLE: str = ""
    料別: str = ""
    成衣件數: int | None = None
    面料數量_yds: float | None = None
    運輸方式: str = ""
    運送數量_yds: float | None = None
    CRFP: date | None = None
    成衣EXP: date | None = None
    空陸運ETC: date | None = None
    空陸運IN_FTY: date | None = None
    海運IN_FTY: date | None = None
    報價已含: str = ""
    空運單價: float | None = None
    實打: float | None = None

    # 以下為計算欄位
    決策時間點: date | None = None
    決策人員: str = ""
    主要確認: str = ""
    結論: str = ""
    超標金額_USD: float | None = None
    狀態: str = ""
    反黃: bool = False
    警示: list[str] = field(default_factory=list)


def _effective_qty(row: Row, group_qty: int | None) -> int:
    """配布列通常不重複填成衣件數，沿用同 STYLE 主布的件數。"""
    return row.成衣件數 if row.成衣件數 is not None else (group_qty or 0)


def apply_rules(row: Row, today: date, group_qty: int | None = None) -> Row:
    """就地計算一列的所有衍生欄位，回傳同一個 Row。"""
    qty = _effective_qty(row, group_qty)

    # ① 決策時間點 = 空/陸運 ETC 前兩週
    if row.空陸運ETC:
        row.決策時間點 = row.空陸運ETC - timedelta(days=DECISION_LEAD_DAYS)

    # ② 決策人員：超過 1 萬件必須上到產區
    row.決策人員 = OWNER_LARGE if qty > QTY_THRESHOLD else OWNER_NORMAL

    # ③ 主要確認語句依運輸方式
    row.主要確認 = CONFIRM_BY_MODE.get(row.運輸方式, "")

    # ④ 結論：海運進廠已來不及趕上成衣 EXP → 只能走空/陸運
    if row.海運IN_FTY and row.成衣EXP and row.海運IN_FTY > row.成衣EXP:
        row.結論 = "無法海運出布(成衣交期急迫)"
    elif row.海運IN_FTY and row.成衣EXP:
        buffer_days = (row.成衣EXP - row.海運IN_FTY).days
        if buffer_days >= DECISION_LEAD_DAYS:
            row.結論 = f"可評估改海運(海運進廠尚有{buffer_days}天餘裕)"

    # ⑤ 超出報價標準的空運金額
    if row.運輸方式 == "空運" and row.空運單價 and row.實打:
        excess = (row.空運單價 - QUOTED_RATE) * row.實打
        if excess > 0:
            row.超標金額_USD = round(excess, 2)

    # ⑥ 反黃：量大 + 走空/陸運 = 一定要有人決策
    row.反黃 = qty > QTY_THRESHOLD and row.運輸方式 in ("空運", "陸運")

    # ⑦ 時效狀態
    if row.決策時間點:
        days_left = (row.決策時間點 - today).days
        if days_left < 0:
            row.狀態 = f"逾期{abs(days_left)}天"
        elif days_left <= URGENT_WINDOW_DAYS:
            row.狀態 = f"剩{days_left}天"
        else:
            row.狀態 = f"剩{days_left}天"
    else:
        row.狀態 = "缺ETC"

    # ⑧ 資料完整性警示 —— 沒有這些欄位就無法判斷
    if not row.空陸運ETC:
        row.警示.append("缺空/陸運ETC，無法算決策日")
    if not row.成衣EXP:
        row.警示.append("缺成衣EXP，無法判斷可否改海運")
    if row.運輸方式 == "空運" and not row.海運IN_FTY:
        row.警示.append("走空運但無海運IN FTY可比對")
    if row.運輸方式 == "空運" and not row.空運單價:
        row.警示.append("走空運但無單價，無法算超標")

    return row


def needs_escalation(row: Row, today: date) -> bool:
    """是否該進今日的催辦名單。"""
    if not row.反黃 or not row.決策時間點:
        return False
    return (row.決策時間點 - today).days <= URGENT_WINDOW_DAYS
