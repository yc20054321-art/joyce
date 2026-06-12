#!/usr/bin/env python3
"""
Daily Email Digest for joycechen@makalot.com.tw
Reads unread inbox via Microsoft Graph API, analyzes with Claude, sends digest email.

Required environment variables (set as GitHub Secrets):
  MS_CLIENT_ID      - Azure AD app (client) ID
  MS_CLIENT_SECRET  - Azure AD client secret
  MS_TENANT_ID      - Azure AD tenant ID
  ANTHROPIC_API_KEY - Anthropic Claude API key
  USER_EMAIL        - (optional) defaults to joycechen@makalot.com.tw
"""

import os
import re
from datetime import datetime, timezone, timedelta

import msal
import requests
import anthropic

CLIENT_ID = os.environ["MS_CLIENT_ID"]
CLIENT_SECRET = os.environ["MS_CLIENT_SECRET"]
TENANT_ID = os.environ["MS_TENANT_ID"]
USER_EMAIL = os.environ.get("USER_EMAIL", "joycechen@makalot.com.tw")
ANTHROPIC_API_KEY = os.environ["ANTHROPIC_API_KEY"]

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
SCOPES = ["https://graph.microsoft.com/.default"]
MAX_EMAILS = 50
TST = timezone(timedelta(hours=8))


def get_access_token() -> str:
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
    )
    result = app.acquire_token_silent(SCOPES, account=None) or app.acquire_token_for_client(scopes=SCOPES)
    if "access_token" not in result:
        raise RuntimeError(f"Graph token error: {result.get('error_description', result)}")
    return result["access_token"]


def get_unread_emails(token: str) -> list[dict]:
    headers = {"Authorization": f"Bearer {token}"}
    url = (
        f"{GRAPH_BASE}/users/{USER_EMAIL}/mailFolders/inbox/messages"
        "?$filter=isRead eq false"
        "&$orderby=receivedDateTime desc"
        f"&$top={MAX_EMAILS}"
        "&$select=id,subject,from,receivedDateTime,bodyPreview,importance,hasAttachments"
    )
    resp = requests.get(url, headers=headers, timeout=30)
    resp.raise_for_status()
    return resp.json().get("value", [])


def format_emails_for_claude(emails: list[dict]) -> str:
    if not emails:
        return "（今日無未讀信件）"

    parts = []
    for i, e in enumerate(emails, 1):
        sender = e.get("from", {}).get("emailAddress", {})
        try:
            dt = datetime.fromisoformat(e["receivedDateTime"].replace("Z", "+00:00")).astimezone(TST)
            date_str = dt.strftime("%Y-%m-%d %H:%M")
        except Exception:
            date_str = e.get("receivedDateTime", "")[:16]

        importance_tag = "【標記重要】" if e.get("importance") == "high" else ""
        attach_tag = "【含附件】" if e.get("hasAttachments") else ""

        parts.append(
            f"[{i}] {importance_tag}{attach_tag}{e.get('subject', '(無主旨)')}\n"
            f"    寄件人: {sender.get('name', '')} <{sender.get('address', '')}>\n"
            f"    時間(TST): {date_str}\n"
            f"    內文摘錄: {e.get('bodyPreview', '')[:250]}"
        )
    return "\n\n".join(parts)


def analyze_with_claude(email_text: str) -> str:
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    today = datetime.now(TST).strftime("%Y-%m-%d")

    prompt = f"""你是 Makalot Industrial Co., Ltd. 的 AI 助理，今天是 {today}（台灣時間）。
以下是 {USER_EMAIL} 收件匣目前的未讀信件清單（摘錄，非完整信件內容）：

{email_text}

請以繁體中文完成以下分析，輸出結構如下：

---

## 一、需立即處理：決策 & 行動優先序

針對需要回覆、核准、決定或採取行動的信件，依緊急程度排序：

| 優先 | 信件編號 & 主旨 | 所需行動 | 建議期限 |
|------|----------------|----------|----------|

若有法務、財務、合約或 ESG 相關事項，請在「所需行動」欄標示「⚠️ 需相關部門確認」。

---

## 二、重要資訊（知悉即可）

列出重要通知、狀態更新等不需立即行動的信件（每項一行）。

---

## 三、可稍後處理

列出電子報、系統通知、廣告等低優先信件（每項一行）。

---

## 四、今日待辦清單（總整理）

以條列式整理今日需完成的具體行動，按優先順序排列。

---

## 備注

- 以上分析基於信件摘錄，AI 推斷僅供參考，請自行複核原始信件
- 敏感資料（報價、成本、個資）已避免於此摘要中重複呈現
"""

    msg = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt}],
    )
    return msg.content[0].text


def md_to_html(text: str) -> str:
    """Minimal markdown → HTML for email rendering."""
    lines = text.split("\n")
    out = []
    in_ul = False
    in_table = False

    for line in lines:
        # Table rows
        if line.startswith("|"):
            if not in_table:
                out.append('<table style="border-collapse:collapse;width:100%;font-size:14px;">')
                in_table = True
            cells = [c.strip() for c in line.strip("|").split("|")]
            is_sep = all(re.match(r"^[-: ]+$", c) for c in cells if c)
            if is_sep:
                continue
            tag = "th" if not any("<td" in r for r in out[-5:]) else "td"
            style = (
                "background:#1a5276;color:white;padding:6px 10px;text-align:left;"
                if tag == "th"
                else "padding:6px 10px;border-bottom:1px solid #ddd;vertical-align:top;"
            )
            row = "".join(f'<{tag} style="{style}">{c}</{tag}>' for c in cells)
            out.append(f"<tr>{row}</tr>")
            continue
        else:
            if in_table:
                out.append("</table>")
                in_table = False

        if line.startswith("## "):
            if in_ul:
                out.append("</ul>")
                in_ul = False
            out.append(
                f'<h2 style="color:#1a5276;border-bottom:2px solid #aed6f1;padding-bottom:4px;">'
                f'{line[3:]}</h2>'
            )
        elif line.startswith("### "):
            if in_ul:
                out.append("</ul>")
                in_ul = False
            out.append(f'<h3 style="color:#1f618d;">{line[4:]}</h3>')
        elif line.startswith("- ") or line.startswith("* "):
            if not in_ul:
                out.append('<ul style="padding-left:20px;">')
                in_ul = True
            content = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", line[2:])
            out.append(f"<li>{content}</li>")
        elif line.strip() in ("", "---"):
            if in_ul:
                out.append("</ul>")
                in_ul = False
            out.append('<div style="margin:8px 0;"></div>')
        else:
            if in_ul:
                out.append("</ul>")
                in_ul = False
            content = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", line)
            out.append(f'<p style="margin:4px 0;">{content}</p>')

    if in_ul:
        out.append("</ul>")
    if in_table:
        out.append("</table>")

    return "\n".join(out)


def send_digest_email(token: str, analysis: str, email_count: int) -> None:
    today = datetime.now(TST).strftime("%Y-%m-%d")
    subject = f"📬 每日信件摘要 & 決策優先序 — {today}"

    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<body style="font-family:Arial,'Microsoft JhengHei',sans-serif;max-width:800px;margin:0 auto;padding:20px;color:#2c3e50;background:#f9f9f9;">
  <div style="background:#1a5276;color:white;padding:18px 28px;border-radius:8px 8px 0 0;">
    <h1 style="margin:0;font-size:20px;">📬 每日信件摘要 &amp; 決策優先序</h1>
    <p style="margin:6px 0 0;opacity:0.85;font-size:14px;">{today}（台灣時間）｜未讀信件：{email_count} 封</p>
  </div>
  <div style="background:#eaf4fb;padding:10px 28px;border-left:4px solid #2980b9;">
    <small style="color:#555;">⚠️ 此摘要由 AI 自動產生，基於信件摘錄內容，<strong>僅供參考</strong>，請自行複核原始信件後再行動。</small>
  </div>
  <div style="background:white;padding:20px 28px;border:1px solid #e0e0e0;border-top:none;">
    {md_to_html(analysis)}
  </div>
  <div style="padding:12px 28px;background:#f0f0f0;border-radius:0 0 8px 8px;border:1px solid #e0e0e0;border-top:none;">
    <small style="color:#999;">Makalot Daily Email Digest — AI 分析結果，需人工審核 | 此郵件由 GitHub Actions 自動發送</small>
  </div>
</body>
</html>"""

    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": [{"emailAddress": {"address": USER_EMAIL}}],
        },
        "saveToSentItems": False,
    }
    resp = requests.post(
        f"{GRAPH_BASE}/users/{USER_EMAIL}/sendMail",
        headers=headers,
        json=payload,
        timeout=30,
    )
    resp.raise_for_status()
    print(f"Digest email sent to {USER_EMAIL}.")


def main() -> None:
    print(f"[{datetime.now(TST).strftime('%Y-%m-%d %H:%M')} TST] Starting daily email digest...")
    token = get_access_token()
    emails = get_unread_emails(token)
    print(f"Fetched {len(emails)} unread email(s).")

    email_text = format_emails_for_claude(emails)
    print("Analyzing with Claude...")
    analysis = analyze_with_claude(email_text)

    send_digest_email(token, analysis, len(emails))
    print("Done.")


if __name__ == "__main__":
    main()
