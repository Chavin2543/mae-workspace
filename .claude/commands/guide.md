---
description: Show a simple guide to everything this workspace can do
---

Greet the user warmly (per the greet-mae skill if they haven't been greeted yet).
Then show this guide, nicely formatted, in simple non-technical English with the
short Thai hints kept as-is. Do not add technical jargon.

# What this workspace can do / พื้นที่ทำงานนี้ทำอะไรได้บ้าง

**Reconcile the half-year report with LS8** (ตรวจกระทบยอดกับไฟล์ LS8)
- Upload the LS8 file and the Segment Half-year file, then type: `/reconcile-ls8`
- Claude fixes the numbers in the LYF tab to match LS8, adds an audit tab
  showing every change, gives you the corrected Excel file, a simple visual
  report, and saves everything to GitHub.

**Just check, don't change anything** (ตรวจอย่างเดียว ไม่แก้ไฟล์)
- Type: `/check-ls8` — you get a list of every number that doesn't match, but
  the file is left untouched.

**See the audit report again** (ขอดูรายงานสรุปอีกครั้ง)
- Type: `/audit-report` — Claude rebuilds the visual summary from the latest
  reconciled file and shows it.

**Save my work to GitHub** (บันทึกงานขึ้น GitHub)
- Type: `/sync` — Claude saves everything so your other computer sees the
  latest files too. This usually happens automatically after each task, so
  you only need it if you want to be sure.

**See where things stand** (ดูสถานะงานตอนนี้)
- Type: `/status` — what's in progress, the newest files, and recent decisions.

**Keep work organized** (จัดระเบียบงาน)
- `/new-task` starts a tracked piece of work, `/task-done` closes it, and
  `/log-decision` writes down a decision so Claude remembers it next time.
- These mostly run themselves — Claude uses them automatically; you only need
  them if you want to check or record something yourself.

**Tips**
- You can always just describe what you want in plain words (English or Thai) —
  the commands are shortcuts, not requirements.
- Files you upload appear automatically; you don't need to tell Claude where
  they are.
- Everything Claude changes is listed in the "Recon LS8" tab inside the Excel
  file, so you can always see before/after values.

After showing the guide, ask if they'd like to start one of these now.
