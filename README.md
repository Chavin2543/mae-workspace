# Mae's Workspace — Hotel Revenue Workbook Reconciliation

A Claude Code workspace for the recurring Excel work of the **Mitsui Thailand
portfolio hotels** — Somerset Rama 9 (SR9), Ascott Embassy Sathorn (AES),
lyf Sukhumvit 8 (LYF), and Somerset Pattaya (SP).

Mae uploads the monthly/half-year workbooks, Claude reconciles them against
the authoritative sources, and every change is saved here with a full audit
trail — so any month's work can be reproduced from this repo alone.
(พื้นที่ทำงานสำหรับตรวจกระทบยอดไฟล์ Excel ของโรงแรมในเครือ — อัปโหลดไฟล์
แล้วให้ Claude จัดการ พร้อมบันทึกทุกการแก้ไขค่ะ)

## How to use it (Mae)

Open Claude Code in this repository and just say what you want, in English or
Thai — or use a shortcut:

| Command | What it does |
|---|---|
| `/guide` | Menu of everything this workspace can do |
| `/reconcile-ls8` | Fix the half-year workbook to match LS8 + audit tab + visual report |
| `/check-ls8` | Compare only — show differences, change nothing |
| `/audit-report` | Rebuild and show the visual audit report |
| `/status` | What's in progress, newest files, recent decisions |
| `/sync` | Save everything to GitHub so every machine sees the latest work |

Uploaded files appear automatically; there's no need to explain where they
are. Every automated change is listed in a "Recon …" tab inside the delivered
Excel file, with before/after values.

## What's inside

```
data/source/     Uploaded Excel originals (read-only, never edited in place)
data/pdf/        Uploaded PDF snapshots, e.g. History & Forecast (read-only)
output/          Deliverables: reconciled workbooks + their HTML audit reports
scripts/         Reusable Python tools (openpyxl) — one per recurring task
docs/decisions/  Decision log — one dated file per ruling Mae has made
tasks/           Work tracking — open tasks and finished ones with outcomes
WORKLOG.md       What each session did, newest first
CLAUDE.md        Instructions for Claude: task anatomy, mappings, rules
.claude/         Commands, skills, and enforcement hooks
resources/       Reference material (gitignored; restored automatically)
```

## How it stays reliable

- **One branch.** Everything lives on `main` — every machine and session
  reads and writes the same history.
- **Originals are untouchable.** Files in `data/` are never edited, moved, or
  deleted; work happens on copies in `output/`.
- **Every change is audited.** Reconciled workbooks carry an audit sheet, and
  a bilingual HTML report is built for each run.
- **Memory that survives sessions.** Decisions land in `docs/decisions/`,
  session notes in `WORKLOG.md`, work items in `tasks/` — a new session can
  catch up from the repo alone.
- **Hooks enforce the rules.** Scripts in `.claude/hooks/` block edits to
  protected files, block stray git branches, auto-commit before new tasks,
  and refuse to end a session with unsaved work. The rules aren't advisory.

## For technical readers

Scripts require Python 3 with `openpyxl` (installed automatically at session
start). Example — the LS8 half-year reconciliation:

```bash
python3 scripts/reconcile_ls8.py \
    --ls8 "data/source/<LS8 file>.xlsx" \
    --halfyear "data/source/<half-year file>.xlsx" \
    --out "output/<half-year file>_LS8-reconciled.xlsx"   # add --dry-run to only compare

python3 scripts/audit_report.py "output/<reconciled file>.xlsx"  # builds the HTML report
```

Scripts assert the workbook layout before writing, change only hardcoded
input cells (never formulas), and append an audit sheet listing every change.
Full task anatomy, segment mappings, verification checklists, and standing
rules are in [CLAUDE.md](CLAUDE.md).
