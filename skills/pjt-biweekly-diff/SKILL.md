---
name: pjt-biweekly-diff
description: 聚陽業務六處每週 PJT review 的「二週差異分析」固定流程。比對 PJT 大表內本週 sheet 與上週 sheet (如 05-26 vs 05-19) 的款號/部門/產區/季度數量變化，產出標準 RECAP Excel (Q×產區二週差異 / 異動說明清單 / 部門產區季度三個 sheet)。當使用者提到「PJT 二週差異」「PJT review」「比較這週跟上週 PJT」「跑 RECAP」「PJT 加量減量」「款號/部門/產區 量的變化」「出一份 PJT 結論」「每週 PJT 表格」時，務必使用本 skill。即使使用者只說「比一下 5-XX 跟 5-XX」或給一個 PJT.xlsx 路徑要看差異，也應觸發。
---

# PJT 二週差異分析

把聚陽業務六處每週 PJT review 的差異分析流程固化。核心是「本週 vs 上週」的款號層級量變化，再彙總到 **季度 × 產區 × 部門**，並標出加量／減量／持平與異動原因。

## 什麼時候用

每週收到新的 `YYYY-MM-DD PJT.xlsx` 後，要回答這類問題：
- 本週相對上週，哪些產區/部門/款號加量、減量？
- 各季度 (Q1–Q4) × 各產區 (印尼/柬埔寨/中國/北越) 的二週淨差是多少？
- 異動原因是什麼 (DROP / NEW ADD / MOVE 移轉 / REDUCE)？
- 出一份對外可用的 RECAP。

## 重要背景知識（這些是踩過坑校正出來的，務必遵守）

1. **檔案結構**：一個 PJT xlsx 內就含多週 sheet，sheet 名是 `MM-DD`（如 `05-26`、`05-19`）。**本週 vs 上週是同一個檔案內的兩個 sheet**，不是兩個檔案。第一個 date sheet 通常是本週，第二個是上週。

2. **「已進單」= 灰底白字**：這種儲存格代表已進系統的確認單，**不算入「開單量」**，做差異時要排除。有兩種 fill 變體都要抓：
   - theme 0 + 負 tint (< -0.15)
   - theme 1 + 正 tint (> 0.3)
   - 或純灰 RGB
   字色為白 (theme 0 或 #FFFFFF)。`scripts/pjt_diff.py` 的 `is_confirmed()` 已實作，**不要簡化成只看一種**，否則會漏抓造成數字暴增。

3. **官方口徑 vs 逐款比對是兩套數字**：
   - PJT 大表本週 sheet 的 **cols 44-60 / rows 4-16** 有官方「部門 × 產區 × 季度」表，**對外報告以這個為準**。
   - 逐款 (style-by-style) 比對是內部追因用，數字會跟官方略有出入（官方含 BUFFER / 手 KEY baseline）。兩者都要呈現，但別把它們混為一談。

4. **noise 列**：`TOTAL / Total / DIFF / revised / 上週 PJT / BUFFER / EXP MONTH / 已進單交期 / 安進產能` 在款號層級要過濾。注意 `Total - 待確認PJT` 是大小寫混合，過濾要 case-insensitive（用 `.upper()`），否則 section 加總列會被當成款號混進去，造成數字錯得離譜。

5. **單位**：標打 ÷ 10000 = 萬打，預設輸出用萬打、1 位小數。加量 > +0.05 萬用藍字、減量 < -0.05 萬用紅字、其餘持平用黑字。即使四捨五入成 ±0.0，產區×季度的小計列仍要列出並標「持平」。

6. **款號顯示**：對外只 show 去掉前 3 碼年份的部分（227N041A → N041A、226F073A → F073A）。

7. **產區順序**：印尼(IND) → 柬埔寨(CAB) → 中國(CHN) → 北越(N.Vin)。季度順序 Q1→Q2→Q3→Q4。

## 標準流程

### Step 1 — 確認來源檔
若使用者沒給路徑，最新 PJT 通常在：
```
\\Nt_pdc\業務六處\外部使用檔\V6\00.每週PJT Review\每週PJT表格 \YYYY-MM-DD PJT.xlsx
```
（注意資料夾「每週PJT表格 」結尾有一個空格。）先 `Copy-Item` 到本地暫存再讀，避免鎖檔。

讀網路檔可能因壞掉的圖片參照報 `KeyError: ... xl/drawings/NULL`；遇到時改用 `load_workbook(path, data_only=True, read_only=True)` 規避。

### Step 2 — 跑分析腳本
```bash
python "C:/Users/JoyceChen/.claude/skills/pjt-biweekly-diff/scripts/pjt_diff.py" \
  "<本地暫存 PJT.xlsx>" --current 05-26 --prev 05-19 \
  --out "<輸出.xlsx>"
```
- 不給 `--current/--prev` 會自動抓前兩個 `MM-DD` sheet。
- 不給 `--out` 預設存在來源檔同目錄。
- 腳本會自動：解析兩週 sheet、排除已進單、配對款號、讀官方表、抓 textbox 異動標註、產出 3-sheet RECAP，並在 stdout 印出各季度淨差摘要。

### Step 3 — 回報
把腳本 stdout 的季度摘要整理給使用者（全年淨差 + 各 Q 加/減/持平），並指出輸出檔路徑。如果使用者要更細的款號明細，RECAP 第 2 個 sheet「異動說明清單」展開部門列即可看到。

### Step 4 — 產 JPG（5 張附圖）
```bash
python "C:/Users/JoyceChen/.claude/skills/pjt-biweekly-diff/scripts/pjt_jpgs.py" \
  "<本地 PJT.xlsx>" --current 06-09 --out-dir "C:/Users/JoyceChen/Desktop/Claude"
```
產出：Q産區矩陣 / Q3 Buffer / Q4 Buffer / BY部門 / BY産區 共 5 張 jpg。

### Step 5 — 建 Outlook 草稿
```bash
python "C:/Users/JoyceChen/.claude/skills/pjt-biweekly-diff/scripts/pjt_email.py" \
  "<RECAP.xlsx>" \
  --pjt-file "<本地 PJT.xlsx>" \
  --current 06-09 \
  --img-dir "C:/Users/JoyceChen/Desktop/Claude" \
  --out-dir "C:/Users/JoyceChen/Desktop/Claude"
```
腳本會：
1. 從 RECAP xlsx 讀「異動說明清單」「Q×産区二週差異」
2. 從 PJT xlsx 讀 Buffer 數字（row31/19）與全年暫定（調整後季列）
3. 產出 `pjt_email_arthur_MM-DD.html` / `pjt_email_dept_MM-DD.html`
4. 產出 `build_pjt_draft_MM-DD.ps1`（HTML 從外部 .html 檔讀入）
5. 印出下一步執行指令

執行 PS1（建草稿）：
```powershell
Get-Content 'C:/Users/JoyceChen/Desktop/Claude/build_pjt_draft_06-09.ps1' -Encoding UTF8 | Out-String | Invoke-Expression
```
或加 `--run` 讓 pjt_email.py 直接執行：
```bash
python pjt_email.py ... --run
```

### Step 6 — 週三 Teams PJT異動追蹤（可選）

週二晚 Joyce Fix 版後，週三整天各部門會在 Teams **PJT大表群** 通知量的異動。
本 Step 把 Teams 訊息整理成 BY部門彙總表，補充 Step 3 / Step 5 的數字脈絡。

**觸發時機**：週三 17:30 前，說「整理本週 Teams PJT大表群的更新訊息」。

**Claude 執行流程**：
1. 用 Teams MCP 搜尋 PJT大表群 週三的訊息：
   ```
   chatId: 19:a07d3bed71da40af95de613c8bb3c893@thread.v2
   afterDateTime: <週二 22:00>  beforeDateTime: <週三 17:30>
   ```
2. 解析每則訊息 → 部門 / 産区 / 款號 / 出口月 / 標打量(±) / 原因
3. 寫入 JSON 暫存檔（格式詳見腳本 docstring）
4. 執行：
   ```bash
   python "C:/Users/JoyceChen/.claude/skills/pjt-biweekly-diff/scripts/pjt_teams_log.py" \
     --json-file "<messages.json>" --week <MM-DD> \
     --out-dir "C:/Users/JoyceChen/Desktop/Claude"
   ```
5. 產出 `PJT_Teams異動記錄_MM-DD.xlsx`（2 sheets：異動明細 / BY部門彙總）

**JSON 格式（每筆 message）**：
```json
{
  "week": "06-30",
  "messages": [
    { "time": "2026-07-01 13:10", "sender": "Joanne Wu", "dept": "#32-solo",
      "region": "印尼", "style": "60327H092A", "month": "10月",
      "qty": 99, "reason": "昨+123今拆", "raw": "原文摘要..." }
  ]
}
```

**部門對應規則**：
| Teams 訊息關鍵字 | 部門欄 |
|---|---|
| `22DIV` / `22Div` | `#22` |
| `SOLO 6032x…` (非 22DIV) | `#32-solo` |
| `23DIV` / `23Div` | `#23` |
| `27DIV` / `27Div` | `#27` |
| `#kids` / `KIDS` | `#kids` |
| (依實際訊息補充) | |

**月份 → 季度**：Q2=4-6月、Q3=7-9月、Q4=10-12月

## 輸出檔結構（scripts/pjt_diff.py 產出）

| Sheet | 內容 |
|---|---|
| **Q×產區二週差異** | 4 季 × 4 產區矩陣 + 季度合計 + 產區合計 + 結論（對外一頁看完）|
| **異動說明清單** | 各 Q 小計 → 各產區 ALL（含款號細項與合計）→ 部門細項預設收起可展開 |
| **部門產區季度 RECAP** | 完整 12 部門 × 16 (Q×產區) 矩陣 + 產區小計 + 季度小計，cell 含 hover comment |

## 字體與配色（符合公司規範）
- 中文：微軟正黑體；英文/數字：Calibri
- 加量=藍 (0070C0)；減量=紅 (C00000)；持平=黑
- 季度分色：Q1 黃 / Q2 紫 / Q3 藍 / Q4 綠

## 對外信件附圖（5 張 JPG）

使用者要寄 PJT review 信時，會用以下 5 張 JPG 內嵌到 Outlook 草稿。由 `scripts/pjt_jpgs.py` 統一產出（命名以本週日期為前綴 `PJT MM-DD ...`）：

| JPG | 內容 | 用在信中段落 |
|---|---|---|
| `PJT MM-DD Q產區矩陣.jpg` | Q2-Q4 × 印尼/柬埔寨/中國/北越 矩陣 + 結論 | 【BY 出口期說明】首張 |
| `PJT MM-DD Q3 Buffer.jpg` | Q3 本週數量 / 二週差異 並排 | 【BY 出口期說明】Q3 段 |
| `PJT MM-DD Q4 Buffer.jpg` | Q4 本週數量 / 二週差異 並排 | 【BY 出口期說明】Q4 段 |
| `PJT MM-DD BY部門.jpg` | 12 部門 × Q2/Q3/Q4/ALL（省略 Q1） | 【BY 部門】 |
| `PJT MM-DD BY產區.jpg` | 二週差異 + 季度差異 + 上週 + 本週 + 調整後季 全塊（含 textbox 註解）| 【BY 產區】 |

執行方式：
```bash
python "C:/Users/JoyceChen/.claude/skills/pjt-biweekly-diff/scripts/pjt_jpgs.py" \
  "<本地 PJT.xlsx>" --current 06-02 --out-dir "C:/Users/JoyceChen/Desktop/Claude"
```
不給 `--current` 會自動抓最前面的 `MM-DD` sheet。
不給 `--out-dir` 預設存在 PJT 檔同目錄。

Buffer 數據來自 PJT 檔的 `Q3+Q4 buffer工作表1` sheet；其他來自本週 sheet 的官方公式格。

## Outlook 草稿與收件人清單

收件人定義在 `config/recipients.json`，**修改名單只改這個檔案，不用動腳本**。

```json
{
  "arthur": { "to": "Arthur", "cc": ["..."] },
  "dept":   { "to": ["..."], "cc": [] }
}
```

雙草稿流程：
1. **Arthur 版**（完整，含 Q3/Q4 Buffer，CC 17 人）— 主管確認用
2. **各部門版**（精簡，**不含 Q3/Q4 Buffer**，To 5 人，無 CC）— Arthur 確認後再發

兩封草稿都用 Outlook **Categories** 標籤區分（`PJT-Arthur` / `PJT-Dept`），同主旨並列不互干擾。

建草稿用 `scripts/pjt_email.py`（v4 版型，確立於 2026-06-10）：
- 讀 RECAP xlsx + 原始 PJT xlsx，自動算出全年暫定、季度淨差、Buffer 數字
- 產出 `pjt_email_arthur_MM-DD.html` / `pjt_email_dept_MM-DD.html`（分離式，避免 PS1 轉義）
- 產出 `build_pjt_draft_MM-DD.ps1`，在 PS1 中用 `Get-Content ... -Raw` 讀入 HTML
- v4 版型結構：矩陣圖 → 【BY出口期說明】→ Buffer(Arthur 版) → 【BY部門】→ 【BY産區】
- 收件人從 `config/recipients.json` 讀，修名單只改 JSON，不動 py 或 ps1

**紅線：草稿 `.Save()` 不 `.Send()`**。Joyce 自己檢查補 To 按送出。

## 常見變化題
- **比對官方版 vs 個人(joyce)版**：兩個不同檔案、同一個 `05-26` sheet。改成讀兩個檔的同名 sheet，比款號的「開單量」與「已進單量」各自差異。多數差異會是「已進單↔開單格式轉換」(淨 0)，要留意真正的 IE 上修或漏列款。
- **只要某一季或某一部門**：跑完整腳本後，從輸出 sheet 篩選即可，不必改腳本。
