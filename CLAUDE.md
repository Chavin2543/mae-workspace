# Mae's workspace — hotel revenue workbook reconciliation

This repo holds Mae's recurring Excel reconciliation work for the Mitsui Thailand
portfolio hotels (SR9, AES, LYF, SP). The recurring task documented here is:
**"Reconcile half year according to LS8"** — done here for the first time in
July 2026 and expected to repeat.

**Mae is a non-technical user.** Keep every reply short and plain — no jargon,
no raw tracebacks (summarize problems in one sentence and say what you'll do
about it). Thai hints are welcome. When she uploads files and asks to
"reconcile", assume the `/reconcile-ls8` flow below.

## Command suite (non-tech entry points)

| Command | What it does |
|---|---|
| `/guide` | Plain-language menu of everything this workspace can do |
| `/reconcile-ls8` | Full run: fix workbook to match LS8 + audit tab + visual report + commit/push |
| `/check-ls8` | Compare only (`--dry-run`) — show differences, change nothing |
| `/audit-report` | Rebuild + show the visual audit report from the latest reconciled file |
| `/status` | What's in the workspace: open tasks, latest files, recent decisions |
| `/new-task` | Start a tracked task (commits pending work first) |
| `/task-done` | Close the open task: outcome, move to done, commit & push |
| `/log-decision` | Record a decision so future sessions follow it |

Command definitions live in `.claude/commands/`. SessionStart hooks
(`.claude/settings.json`) install python deps, restore `resources/`, inject
workspace state (open tasks, recent decisions), and point the user to `/guide`.

## Directory layout

```
.claude/           Commands (commands/*.md), hooks (hooks/*, settings.json).
data/source/       Uploaded input workbooks, renamed to clean names. READ-ONLY.
data/pdf/          Uploaded PDF snapshots (History & Forecast etc.). READ-ONLY.
docs/decisions/    Decision logs — one dated file per decision. See README there.
tasks/open|done/   One file per tracked task. See tasks/README.md.
output/            Deliverables: reconciled workbooks + their HTML reports.
scripts/           Reusable Python scripts (openpyxl). One script per recurring task.
resources/         Reference material, gitignored (large), auto-restored by hook.
CLAUDE.md          This file — the anatomy of each recurring task.
```

### Where every file goes

| File type | Home | Notes |
|---|---|---|
| Uploaded Excel sources | `data/source/` | Rename to clean names; never edit/delete (hook-enforced) |
| Uploaded PDFs | `data/pdf/` | Same rules as sources |
| Reconciled/output workbooks | `output/` | `<original name>_<what>-reconciled.xlsx`, audit sheet inside |
| HTML/visual reports | `output/` | Next to the workbook they describe (`<stem>_report.html`) |
| Python scripts | `scripts/` | One per recurring task, argparse-style like `reconcile_ls8.py` |
| Decision records | `docs/decisions/` | `YYYY-MM-DD-slug.md` from the template |
| Task records | `tasks/open/` → `tasks/done/` | `YYYY-MM-DD-slug.md` from the template |
| Reference material | `resources/` | Gitignored, read-only |
| Temp/scratch files | session scratchpad | Never the repo — the Stop hook rejects stray files |

A new file type gets a home added to this table (log the decision) — nothing
goes in the repo root except this file.

### resources/claude-cookbooks

A read-only clone of https://github.com/anthropics/claude-cookbooks (official
Anthropic examples: Claude API patterns, tool use, skills, agent SDK, etc.).
It is **not committed** (356 MB, gitignored); the SessionStart hook shallow-
clones it in the background if the folder is missing, so give it a moment on a
fresh session. Use it as reference when building or improving automation in
this workspace. Never edit files inside it, and never commit it.

Conventions: commit both the source snapshot and the reconciled output for each
run, so every reconciliation is reproducible from the repo alone. Use clear
names (`<original name>_LS8-reconciled.xlsx`).

## Task anatomy: "Reconcile half year according to LS8"

### The two files

1. **LS8 workbook** (`LS8_ Market Segment 2025 &YTD 2026.xlsx` or similar) —
   the **authoritative source**. LS8 = **lyf Sukhumvit 8 Bangkok**.
   - Sheet `LS8 Segment (2026)`: current-year months in cols D..O, segments in
     rows 9–38 (RNs / ADR / Revenue triplets), actuals block rows 50–53
     (`Actual Occupied Rooms w/o COMP, HU, SYS`, `Actual OCC (%)`,
     `Actual ADR`, `Actual Revenue`).
   - Sheet `LS8 Segment 2025`: prior year, same shape shifted up two rows
     (segments rows 7–36, actuals rows 48–51).
2. **Segment Half-year workbook** (`Segment_Half year_version 1.xlsx`) — the
   **target**. LS8 data lives on the **`LYF` tab** (plus `Summary` rows for LYF).
   Other tabs (SR9, AES, SP, Compset, Arrival, Summary-1) are other hotels /
   other sources — do not touch them for this task.

### Segment name mapping (LS8 → half-year LYF tab)

| LS8 sheet row label | LYF tab label |
|---|---|
| Corporate SS / COR SS | Corporate Short Stay |
| LS | Corporate Long Stay |
| Online | Online Business (Dynamic Rate) |
| ASR | ASR |
| Wholesale | Wholesale (Static Rate) |
| Group Corporate + Group Leisure + Group Series (summed) | Corporate Group |
| Employee Travel | Employee Travel |
| COM, HU, SYS | *excluded* — LYF totals are "w/o COMP, HU, SYS" |

### What is hardcoded input vs formula in the half-year LYF tab

- 2026 block: RN overview `C7:N7`, RN by segment `C14:N20`, revenue by segment
  `C24:N30` are hardcoded. Occ/ADR overview rows 8–9 pull from
  `Summary!C51:N52` (2026) which pull from **external workbooks** — see warning.
- 2025 block: RN overview `C51:N51`, RN by segment `C58:N64`, revenue by
  segment `C68:N74` are hardcoded. Occ/ADR rows 52–53 pull from
  `Summary!C60:N61` (hardcoded there). Revpar rows are `=Occ*ADR` formulas.
- Totals, mix %, MoM, and the check row (`C66:N66`, must all be TRUE:
  segment total = overview RN) are formulas — fix inputs only, let them recalc.
- Nationality blocks (rows 33–44 / 77–88) come from the monthly nationality
  workbook, **not** LS8 — out of scope.
- `Summary` "Revenue (THB)" rows for LYF (2026 row 57, 2025 row 63) are
  consistently ~3–4% above LS8 room revenue in every month of both years —
  different basis (not room-revenue-only). **Do not reconcile them to LS8.**

### How to run

```bash
pip install openpyxl   # if missing
python3 scripts/reconcile_ls8.py \
    --ls8 data/source/<LS8 file>.xlsx \
    --halfyear data/source/<half-year file>.xlsx \
    --out output/<half-year file>_LS8-reconciled.xlsx
```

The script asserts the layout (labels in col B) before writing, overwrites only
hardcoded input cells that differ (never formulas), and appends an audit sheet
`Recon LS8 (<year>)` listing every change old → new plus notes on what was
deliberately left alone. Add `--dry-run` to print differences without writing
anything (used by `/check-ls8`). Current version reconciles the **2025 block**
(per Mae's instruction July 2026); extend `main()` with the 2026 mapping (rows
in the module constants) when asked to reconcile the current year.

After reconciling, build the non-tech audit presentation:

```bash
python3 scripts/audit_report.py output/<reconciled file>.xlsx
```

It reads the audit sheet and writes `<stem>_report.html` next to the workbook
(bilingual EN/TH, light+dark theme). Deliver it with SendUserFile
(`display: render`) and republish the artifact (same file path keeps the same
URL within a session; pass the previous artifact URL from new sessions).

### Verification checklist (do all of these before delivering)

1. Script output: review the change list; every diff should trace to LS8.
2. Python re-check: segment RN sums per month == LS8 "Actual Occupied Rooms
   w/o COMP, HU, SYS" == overview RN row; revenue sums per month == LS8 Total
   Revenue; year totals match (2025: RN 59,216 / revenue 82,620,532.57).
3. Confirm the out-of-scope blocks are byte-identical to the original.
4. **External links warning**: this workbook links to other workbooks
   (`Pick Up Pace_SR9`, `result FY26`, etc. — 221 formula cells). Do **not**
   run LibreOffice `recalc.py` on the deliverable: it resolves those links,
   fails, and destroys them (#NAME?). Deliver the openpyxl-saved file directly;
   Excel recalculates everything (incl. check row 66) on open. If you need a
   LibreOffice recalc for QA, do it on a throwaway copy after freezing the
   at-risk cells (`external_links_at_risk()` in the xlsx skill's recalc.py)
   with cached values from the original. Note: in the July 2026 session
   LibreOffice failed to finish recalculating this workbook even after 50
   minutes — do NOT attempt it; the Python re-check (step 2) is the
   verification method for this workbook. Excel recalculates the simple
   SUM/mix/check formulas normally on open.

### Known open items (July 2026 run)

- 2026 H1 block has its own small diffs vs LS8, out of scope per Mae's choice:
  RN Feb Online 3409→3410, Mar Online 4095→4094, Feb Wholesale 1021→1020;
  revenue Online Jan/Feb/Mar + ASR Jan low by ~93,027 THB total
  (H1 38,106,441 vs LS8 38,199,467). Listed in the audit sheet notes.

## General rules for this repo

- **Git branch: always work on `main` only.** Do all work, commits, and pushes
  directly on the `main` branch. Never create or switch to another branch
  unless Mae explicitly asks. (If the session harness forces a different
  designated branch, follow the harness — but never invent branches yourself.)
- **Always commit before starting a task.** Real work gets a task file in
  `tasks/open/` first (see `tasks/README.md`); close it via `/task-done`.
- **Log decisions.** Anything Mae decides that future sessions must respect
  goes in `docs/decisions/` (`/log-decision`). Check there before re-asking or
  re-litigating a settled question.
- Ask Mae before touching tabs other than the one being reconciled.
- Every automated edit to a workbook must leave an audit trail (audit sheet
  and/or report committed to the repo).
- Keep fonts/formats of edited cells as found; write plain values, not styles.
- Commit messages: `Reconcile <scope> vs <source> (<month year>)` for
  reconciliations; `task: start|done <slug>`, `docs: log decision <slug>`,
  plain descriptive messages otherwise.

## Enforcement hooks (`.claude/hooks/`)

These rules are not advisory — hooks enforce them and Claude cannot skip them:

| Hook | Event | What it does |
|---|---|---|
| `protect_paths.py` | PreToolUse (Write/Edit) | Blocks edits inside `data/source/`, `data/pdf/`, `resources/` |
| `guard_bash.py` | PreToolUse (Bash) | Blocks branch creation, force-push, rm/mv/redirect into `data/source/` |
| `checkpoint.py` | UserPromptSubmit | Auto-commits pending changes before each new task |
| `finish_guard.py` | Stop | Refuses to end the turn with uncommitted or unpushed work |
| `session_status.sh` | SessionStart | Injects branch, open tasks, recent decisions into context |

If a hook blocks you, it is enforcing a rule from this file: don't work around
it (no bypass flags, no copies of protected files back into protected paths) —
do the compliant thing instead. Hook changes are decisions: log them.
