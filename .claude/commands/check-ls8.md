---
description: Check the half-year workbook against LS8 without changing anything (report only)
---

Run the LS8 reconciliation in CHECK-ONLY mode — the user wants to see the
differences but nothing may be modified. Audience is non-technical: plain
short language, no jargon, no tracebacks.

Arguments (optional): $ARGUMENTS may name the year block ("2025", "2026",
"both"). Default: ask with AskUserQuestion.

Steps:
1. Locate input files exactly as described in `.claude/commands/reconcile-ls8.md`
   step 1 (uploads first, then `data/source/`).
2. Run `python3 scripts/reconcile_ls8.py --dry-run ...` — this prints every
   difference but writes no workbook.
3. Present the result as a short, friendly table grouped by area
   (Room Nights / Revenue / Occupancy-ADR), showing month, old value,
   LS8 value, and difference. If everything matches, celebrate briefly and
   say no changes are needed.
4. Offer to run `/reconcile-ls8` to actually fix the file if differences
   were found. Do not commit anything.
