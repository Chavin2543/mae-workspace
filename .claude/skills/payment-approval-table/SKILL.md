---
name: payment-approval-table
description: Read uploaded payment vouchers (PV) — PDFs, images, or screenshots — and summarize them into copy-pasteable approval-email tables, one table per paying company, with No. / Detail / Payment Voucher columns (Description, Pay to, Amount in THB). Use whenever Mae uploads payment vouchers and says anything like "payment approval", "summarize the PV", "ทำตาราง PV", "payment table", "ขออนุมัติจ่าย", or asks to prepare payments for email approval — even if she just uploads voucher files and says "do the payment".
---

## What this is

Mae's recurring payment-approval step (first set up July 2026): each round she
receives ~4 payment vouchers (PVs) from the accounting system (Ananda
Development format), each for a **different company** in the portfolio. She
needs a summary table per company that she can **copy-paste straight into an
approval email**. The email approver reads the table, sees the voucher image
below/next to it, and approves.

## Inputs

- The uploaded PVs — PDF, image, or screenshot. One PV can be 1–2 pages.
- If uploaded as PDF files, save them to `data/pdf/` with clean names
  (`PV_<company>_<doc-no>.pdf`). Chat-pasted images need no filing.
- Read every page (use the pdf skill for PDFs). If a voucher is unreadable or
  a field is missing, ask Mae — never guess an amount.

## What to extract from each voucher

| Field | Where on the Ananda PV | Notes |
|---|---|---|
| **Company** | Header, e.g. "บริษัท เอเอ็มเอช รัชดา จำกัด" | Use the English name (e.g. AMH Ratchada Co., Ltd.). One table per company. |
| **Description** | Journal Description / Item Text / Payment Description | Expand Thai tax shorthand to plain English: `ภพ.30` → Value Added Tax Return (P.P. 30); `ภงด.1/3/53` → Withholding Tax Return (P.N.D. 1/3/53); `ภธ.40` → Specific Business Tax (P.T. 40). Thai years are Buddhist Era: `มิ.ย. 69` = June 2026. |
| **Pay to** | Recipient name / Vendor | e.g. กรมสรรพากร → Revenue Department. |
| **Amount** | Payment details → Amount Settled (THB) total | Show positive, 2 decimals, thousand separators, suffix ` THB`. |

Cross-checks before presenting: debit total = credit total on the accounting
block, and the amount you show = the Payment details total. If they disagree,
flag it to Mae instead of picking one.

## Output format

One table per company, in chat, ready to copy. Company name bold above its
table. Number the vouchers 1..n within each company. Exactly this shape:

**AMH Ratchada Co., Ltd.**

| No. | Detail | Payment Voucher |
|---|---|---|
| 1 | Description: Value Added Tax Return (P.P. 30) for June 2026<br>Pay to: Revenue Department<br>Amount: 1,297,598.79 THB | *(paste voucher image here)* |

- The **Payment Voucher** column stays as the placeholder text — Mae inserts
  the voucher picture herself in the email. Remind her once, briefly.
- After all tables, add one line with the grand total across all companies so
  Mae can sanity-check (this line is for her, not part of the tables).
- Keep the reply short: the tables, the total, one reminder line. No analysis.

## After delivering

- Nothing in the repo changes for a normal run (output lives in chat), but if
  PDFs were saved to `data/pdf/`, commit them with a WORKLOG entry and land on
  `main` per CLAUDE.md.
- If Mae later asks for the tables as a file (Word/HTML) or wants the email
  drafted and sent, that's an extension — log a decision when the format
  changes.
