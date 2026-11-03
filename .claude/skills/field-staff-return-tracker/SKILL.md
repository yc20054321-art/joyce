---
name: field-staff-return-tracker
description: Record and maintain the tracker for 駐廠業務 (field-stationed sales staff) return-to-Taiwan trips — their 返台日期, 述職日 (office debrief day), and 返回產區日. Use this skill whenever the user pastes a screenshot of the internal "返台暨休假日程規劃表" or "休假紀錄查詢" system page, mentions 駐廠/派駐/述職/返台紀錄, asks to log or update a staff member's return trip, asks about "this tracker" or "the tool on my desktop", or wants a 述職日 calendar reminder set up. Trigger even if the user just pastes a screenshot with no explanation — that is the primary way this skill gets used.
---

# Field staff return-to-Taiwan tracker

A live tool (published Claude Artifact) plus an Outlook calendar-reminder habit for
tracking when 駐廠業務 (staff stationed at overseas factories) fly back to Taiwan,
when they debrief at the office, and when they head back to site.

**Live tool URL:** https://claude.ai/code/artifact/6d805a94-7fe1-484e-a9b0-a8279cbb5271

The data lives in that Artifact's shared database (collection `records`), accessed via
the Artifact tool's `read_db` / `write_db` actions — not by editing any file. A copy of
the tool's HTML source is kept at `field_staff_return_log.html` in this repo for version
history; only touch that file when the tool's UI/behavior itself needs to change (adding
a field, changing a formula, restyling). Adding, editing, or querying a trip record never
touches that file.

## Data model

Each document in the `records` collection:

| Field | Type | Meaning |
|---|---|---|
| `staffName` | string | Keep both Chinese and English name as shown, e.g. `"張念平 Louis Zhai"` |
| `site` | string | 駐廠地點/廠區 |
| `returnDate` | `YYYY-MM-DD` | 離開產區日 — the day they leave site / head to Taiwan |
| `debriefDate` | `YYYY-MM-DD` \| `null` | 述職日 — the day they report to the Taipei office. Often not on the screenshot; leave `null` when absent |
| `leaveEndDate` | `YYYY-MM-DD` \| `null` | 返回產區日 — the day they return to site |
| `note` | string | Freeform: department, title, phone numbers, special leave (病假/事假) details, anything that doesn't have its own field |
| `isSample` | boolean | Demo data flag (should be `false` for anything real) |
| `reminderCreated` | boolean | Whether the 述職日 Outlook reminder has already been created for this record |
| `createdAt` | ISO datetime | When the record was added |

## Recording a trip from a screenshot

Two kinds of screenshot come in — handle both, and tell the two apart by what's on
screen rather than waiting for the user to label it:

**1. The full form** — the internal "返台暨休假日程規劃表" form or the "休假紀錄查詢"
system page. It may be cropped — read what's there and say so if a field is missing
rather than guessing. Map fields as follows:

- 姓名 → `staffName` (keep Chinese + English exactly as shown)
- 廠區 → `site`
- 職稱 (only when it's NOT the default "業務" — e.g. 副理/資深經理), 返台手機,
  返台市內電話 → fold into `note`, semicolon-separated. **Skip 部門 entirely**: every
  record so far is "台北業務六處C部", so writing it into every note just repeats the
  same boilerplate string down the whole column and pushes out the details that
  actually differ row to row (the table's 備註 column truncates, so the first thing in
  the string is what the user actually sees at a glance). If a department ever shows up
  that ISN'T 台北業務六處C部, that's worth a note since it'd be the exception.
- 起訖日期「離開產區&返回產區日」 → start is `returnDate`, end is `leaveEndDate`
- 述職日 → `debriefDate`. **If the screenshot doesn't show this field (cropped, or
  the form section just isn't filled in), leave it `null` and say so plainly** — do not
  treat a missing 述職日 as an error or ask the user to re-send the screenshot unless
  they want to.
- 不進辦公室原因說明 (當地休假/他國休假/其它, e.g. 家庭因素, 喪葬事宜) → when
  checked, this is *why* `debriefDate` is null — fold the reason into `note` and tell
  the user plainly that no debrief day is expected this trip, not that one is missing.
- 特殊假別 (病假/事假/喪假 etc.) → fold into `note`. The tool's 休假天數 formula
  only subtracts Sundays and the debrief day — it does NOT yet subtract special leave
  days. If the form's stated "本次休假X天" doesn't match what the tool will compute,
  say so to the user rather than silently forcing one number to match the other.

**2. The system's own approval-notification email** (sender `sysAlert@makalot.com.tw`,
displayed name "派駐人員休假系統通知"), screenshotted straight from Outlook. Subject
reads「休假申請單 - 核准通知」or「派駐休假交通報支申請單 - 核准通知」; the body is a
fixed template: "申請人 {staffName} 於 {填表時間} 填寫「休假申請單」... 休假N天[+病假
M天]期間為 {start} ~ {end}". This only gives `staffName`, `returnDate` (start),
`leaveEndDate` (end), and whatever leave-type note is in the subject line — there is no
廠區, 述職日, 職稱, or phone number in it. Record what it gives you (`debriefDate:
null`, since the email genuinely doesn't contain it — this is a data-availability gap,
not the "normal, no debrief expected" case above) and say plainly that site/debrief/
contact details are still pending a fuller screenshot. The user has chosen to forward
these to Claude as soon as they arrive (see "Known gap" below for why) — treat that as
the normal, expected way this skill receives new trips, not a stopgap.

**Doc id:** `{returnDate}-{english-first-name}-{english-last-or-initial}`, lowercase,
hyphenated (e.g. `2026-10-06-yuffie-yu`). Same staff name + same `returnDate` = update
the existing doc (it's the same trip, more detail arrived — this is exactly what
happens when the approval-email screenshot above is later followed by the full form).
Same staff name + a *different* `returnDate` = a new trip (this is a recurring
rotation — the same person returns many times over the year); use a distinguishing
suffix if a doc id would otherwise collide (e.g. `-2`).

Write with the Artifact tool's `write_db` action (`set` for a new doc, `update` for
filling in more detail on an existing one) against the URL above, collection `records`.
`update` calls on this contract require `if_version` — `read_db get` the doc first if
you don't already have its current version in hand from this conversation.

## Policy rules — established with the user, don't relitigate

- **No 述職日 is normal, not a warning.** A trip where the staff member never comes
  into the office is a completely ordinary outcome. The tool already renders this as a
  neutral "未進公司述職" badge — never describe it as overdue, late, or a problem, and
  never suggest adding red/urgent styling for it.
- **Table sorts by `returnDate` ascending** (soonest trip first). Past-vs-upcoming and
  next-week coloring is handled by the tool's own CSS/JS — nothing for this skill to
  compute when just adding records.
- **休假天數 is auto-computed by the tool**, not by this skill: every day from
  `returnDate` to `leaveEndDate` inclusive, minus Sundays, minus the debrief day itself
  when it falls in range. See the caveat above about special leave days.

## 述職日 calendar reminder

Whenever a record ends up with a non-null `debriefDate` and `reminderCreated` is not
`true` (whether from a fresh screenshot or filling in a date that was previously blank),
create an Outlook event for the manager (via the Microsoft 365 connector,
`outlook_create_event`):

- **Date:** the previous business day before `debriefDate` (Mon–Fri only — if
  `debriefDate` is a Monday, the reminder goes on the preceding Friday).
- **Time:** `09:00`–`09:30`, `timeZone: "Asia/Taipei"`.
- **Subject — exact format, do not vary it:** `"{M}/{D}  {EnglishFirstName} 進公司"`
  (two spaces before the name), using the *debrief* date's month/day with no leading
  zeros, and just the English first name — no last name, no Chinese name. Example:
  `debriefDate` `2026-09-14` for Sophie → `"9/14  Sophie 進公司"`.
- **Body:** plain text with `staffName`, `returnDate`, `debriefDate`, `leaveEndDate`.
- After creating the event, `write_db update` the record with `reminderCreated: true`
  so a later pass never duplicates it.
- If a `debriefDate` on an existing record *changes* after a reminder was already made,
  update that same calendar event (`outlook_update_event`) rather than creating a
  second one. This skill doesn't persist the Outlook event id anywhere — if you can't
  find the original event to update (e.g. a fresh session with no memory of its id),
  it's fine to create a new one and just tell the user there may now be two reminders
  for that debrief so they can delete the stale one.

## Known gap: can't backfill from the user's mailbox after the fact

The company's leave system sends approval emails (sender `sysAlert@makalot.com.tw`,
subject containing `休假申請單` or `派駐休假交通報支申請單` + `核准通知`) — see the
email-screenshot format above. Two things rule out quietly catching these on our own
after they've arrived, rather than relying on the user to screenshot them:

1. **A recurring automated mailbox-scan trigger is not set up.** It was discussed and
   an attempted scheduled trigger was blocked by a permission classifier (recurring,
   unattended mailbox access + calendar writes need explicit user sign-off). Don't
   assume one is running; if the user asks for a weekly/periodic sweep, treat it as a
   fresh request to confirm and set up.
2. **Even an on-demand search can come up empty for emails that are genuinely still
   there** — just not reachable. The user files read notifications into a folder
   named "012 驗布報告" that lives under "101個人資料夾-BAK" in their Outlook folder
   pane. That naming ("個人資料夾" = Personal Folders) means it's a local `.pst`
   archive file on their machine, not a folder in the Exchange Online mailbox — the
   Microsoft 365 connector (Graph API) cannot see into a local PST at all, by
   `folderName` lookup or by content search, because the mail has left the server
   mailbox entirely once moved there. If `outlook_email_search` comes up empty for a
   date range the user expects results in, check whether they've already filed those
   emails away before concluding there's nothing to find.

Given both of those, the user's chosen workflow is: **they forward/screenshot each
approval notification to Claude as soon as it arrives**, before it gets filed away.
Treat that as the normal, reliable path for this input type — not a workaround.
