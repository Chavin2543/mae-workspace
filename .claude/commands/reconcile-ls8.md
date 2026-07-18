---
description: Reconcile the Segment Half-year workbook against the LS8 workbook (fix + audit + report)
---

Perform the full "reconcile half year according to LS8" task. Follow the
procedure in CLAUDE.md ("Task anatomy" section) exactly. The user may be
non-technical: keep every message short and plain, no jargon, and never show
raw tracebacks (summarize problems in one sentence and say what you'll do).

Arguments (optional): $ARGUMENTS may name which year block to reconcile
(e.g. "2025", "2026", or "both"). If not given, ask with AskUserQuestion —
offer: current year block / prior year block / both.

Steps:
1. Find the two input files. Prefer files the user uploaded this session
   (newest LS8-like file and newest Segment Half-year file). If either is
   missing, fall back to the newest matching file in `data/source/`. If still
   ambiguous, ask which file to use.
2. Copy fresh uploads into `data/source/` with clean names (no upload-hash
   prefixes).
3. Run `python3 scripts/reconcile_ls8.py` with the right flags (see script
   docstring; use `--dry-run` first if you want to sanity-check volume of
   changes). Read CLAUDE.md for the external-links warning: NEVER run
   LibreOffice recalc on the deliverable.
4. Verify with the checklist in CLAUDE.md (Python re-check of totals vs LS8;
   confirm out-of-scope blocks unchanged).
5. Build the visual audit report: `python3 scripts/audit_report.py <reconciled.xlsx>`
   which writes an HTML file next to the workbook.
6. Deliver to the user: send the reconciled workbook (attach) and the HTML
   report (render). Give a 3-5 line plain-language summary: how many numbers
   changed, the biggest corrections, and anything deliberately left alone.
7. Commit everything (sources, output, report) and push to the designated
   branch. Commit message per CLAUDE.md convention.
