# Work log

Short notes on what each session did, newest first — so any machine/session
can see the state of the workspace from git alone. One entry per finished
task: date, what was done, files touched. Keep entries to 2–4 lines, plain
language.

---

## 2026-07-19 — Deck financials now read the monthly P&L blocks only (Mae's rule)
Per Mae: the "YTD performance" summary at the top of the result sheets is
never used anymore. All deck YTD figures = Jan–Jun sums of the monthly
blocks; GOP margin budget now "—" (no revenue budget exists to compute it).
Verified all 5 slides vs independent sheet sums. Decision logged
(docs/decisions/2026-07-19-fin-monthly-blocks-only.md); deck + scripts updated.

## 2026-07-19 — Deck: fixed table overlay; revenue budget removed (not in file)
Mae was right: result sheets have budget rows only for OPEX/GOP/margin/EBIT —
the revenue "budget" was a sheet formula (GOP−OPEX), so the deck now shows —
for revenue budgets, like NPAT. Monthly P&L table tightened so it no longer
overlaps the footnote (all 5 financial slides). All lines recheck clean vs
result FY25/FY26. Files: output deck + scripts/management_deck/build_deck.js.

## 2026-07-19 — Deck: monthly P&L tables now compare ’26 vs ’25 on every line
Per Mae: added OPEX ’25, GOP ’25, EBIT ’25, NPAT ’25 rows to the monthly P&L
table on all 5 financial slides, paired under their ’26 rows (’25 in grey).
Deck rebuilt; extract_deck_data.py output path is now an argument.
Files: output/Portfolio_Performance_Report_H1-2026.pptx, scripts/management_deck/.

## 2026-07-19 — Project README
Added README.md: what the workspace is, Mae's commands, directory map, the
reliability rules (one branch, read-only originals, audit trails, hooks), and
script usage for technical readers. Root-file exception list updated in
CLAUDE.md.

## 2026-07-19 — Filing system, decision logs, task workflow, enforcement hooks
Merged the workspace-organization system into main: filing table in CLAUDE.md
(data/pdf, docs/decisions, tasks/open|done), decision logs seeded with July
rulings, task records, and hooks that enforce the rules (protect read-only
paths, block new branches/force-push, auto-commit before tasks, block finishing
unsynced). Stop hooks consolidated: finish_guard.py replaces git_sync_check.py;
git_session_sync.py now wired into SessionStart. New commands: /status,
/new-task, /task-done, /log-decision.

## 2026-07-19 — Branch cleanup + merged workspace-organization system
Mae set main as default and deleted the old branches. Merged the other
session's filing/task/decision system (tasks/, docs/decisions/, /status,
/new-task, /task-done, /log-decision, enforcement hooks) into main,
keeping /sync and all existing rules. This session now works on main
directly. One branch may remain while that other session is active.

## 2026-07-19 — All branches reconciled into main + "main only" rule
Merged claude/main-branch-only-rule-c5lwl2 (main-only rule, claude-cookbooks
reference library hook) into main; verified every other branch is fully
contained in main. Deleted merged disposable branches. New standing rule in
CLAUDE.md: always work on main ("always").

## 2026-07-19 — New skill: read-excel (read the WHOLE file, every tab)
Mae's lesson from the missed column-R data: created `.claude/skills/
read-excel/` (mandatory full-file scan before any Excel work or any
"data missing" claim) + `scripts/excel_map.py` (maps every sheet's true
size/data regions; --find searches every cell). Rule added to CLAUDE.md.

## 2026-07-19 — P&L tables: added "vs H1 25" column
Every financial slide's P&L table now has a colored vs-H1-2025 column
(% change; pts for margin; THB change for NPAT), per Mae.

## 2026-07-19 — Found monthly P&L + NPAT in result sheets; financials complete
Mae was right: monthly P&L (incl. NPAT) exists in the col-R blocks of both
result sheets (earlier scan stopped one column short). Financial slides now
show real H1 2025 for every line (no more n/a), an NPAT row and tile
(portfolio H1 2026 NPAT +7.9M vs -12.6M LY), and a monthly P&L table
(Revenue/OPEX/GOP/EBIT/NPAT) per slide.

## 2026-07-19 — Financial slides: H1 2025 comparison + label renames
P&L tables now compare against H1 2025 (revenue — the only 2025 line
recorded by half year; other lines n/a with FY 2025 kept in the footnote).
Renamed "MF Proj." to "MF budget" and "vs Proj." to "vs MF budget" (Mae).

## 2026-07-19 — Financial slides: 2025 reference + monthly tables; STR labels
Each P&L slide now shows an FY 2025 actual column (full year — the result
sheets have no H1 2025 P&L split) and a monthly revenue table in THB M for
2025 and 2026 with YoY row. Data labels added to the Bangkok/Pattaya STR
charts (the only charts without number tables); other charts stay clean.

## 2026-07-19 — Deck refreshed + new Financial Performance section
Rebuilt the management deck from the results-checked workbook (corrected
Jan occupancy flows through). New Section 4: P&L for portfolio + each
property — Revenue/OPEX/GOP(EBITDA)/JV/EBIT, Actual vs Ascott BP vs MF
Projection, plus monthly revenue 2026-vs-2025 charts (AES/SP June derived
from result totals). NPAT not in result sheets — noted on slides.
Seg & nationality renumbered to Section 5; deck now 42 slides.

## 2026-07-19 — Standing rule: result sheets are the official record
Agreed with Mae: never change the result FY25/FY26 numbers (frozen at
accounting close; ADR incl. breakfast except LYF). Booking-system RN
exports drift daily — use only for segment/nationality mix. Rule added
to CLAUDE.md.

## 2026-07-19 — Aligned workbook to result sheets (Mae's basis ruling)
Mae confirmed 2025 ADR already includes breakfast, so the 2026 incl-BF
numbers were already correct — no basis change. Fixed the 9 remaining
mismatches: Jan-2026 occupancy roots (AES tab, Summary LYF+SP — Mae's
explicit OK) and SR9 tab 2025 ADR row. SR9 Jan-2026 occ left as formula
(RN-derived, 87.71% vs result 87.02% — noted in audit). Jul-Dec 2025
gaps left per Mae. Audit sheet "Recon result check (Jul 2026)".
File: output/Segment_Half_year_ALLreconciled_results-checked.xlsx.

## 2026-07-19 — Cross-check vs result FY25 / result FY26 sheets
Mae uploaded a new master with "result FY25/FY26" sheets as the truth for
Occ/ADR/RevPAR. Checked all 4 properties, both years, vs Summary + property
tabs. Key finding: 2026 ADR/RevPAR in the workbook = ADR incl. breakfast
(matches the "ADR BF" row), while 2025 = excl. breakfast — mixed basis for
SR9/AES/SP (LYF ok). Details reported to Mae; no cells changed yet.
File saved: data/source/Segment_Half_year_ALLreconciled_with_results.xlsx.

## 2026-07-18 — Management PowerPoint report (H1 2026)
Built `output/Portfolio_Performance_Report_H1-2026.pptx` (27 slides, native
editable charts) from the reconciled workbook: arrivals (MOTS/Chinese/India/
Middle East/long-haul), Bangkok+Pattaya STR, portfolio & per-property
Occ/ADR/RevPAR vs budget (monthly+YTD), segmentation & nationality per
property. Rebuild scripts in `scripts/management_deck/` (extract_deck_data.py
→ deck_data.json → build_deck.js).

## 2026-07-18 — Deleted unverified AES Jul-Dec 2025 numbers
Mae's instruction: the AES tab's Jul-Dec 2025 block (overview + RN/revenue
by segment, 108 cells) had no source file to verify against, so it was
blanked in the ALL-reconciled deliverable. Old values are preserved in new
audit sheet "Recon AES clear (2025 Jul-Dec)"; formulas/Jan-Jun untouched.
New reusable `scripts/clear_cells.py`; report rebuilt.

## 2026-07-18 — Mae's rule: never touch the Summary tab (all properties)
Reverted every Summary-sheet edit and rebuilt both deliverables from the
original workbook; Summary is now byte-identical to Mae's upload. Both
reconcile scripts skip Summary by default (`--include-summary` kept as an
explicit override). Rule added to CLAUDE.md. Changes now: LS8 84 cells,
SR9/AES/SP 121 cells (was 108/148 with Summary).

## 2026-07-18 — Reconciled SR9 / AES / SP tabs (H1) from property sources
New sources in `data/source/`: SR9 Market Segment (2025+YTD2026), AES & SP
2025 Market Mix "as Jun25" (H1 only). New `scripts/reconcile_segments.py`
(surgical patch, same engine as LS8): SR9 2025 H1 + 2026 H1, AES 2025 H1,
SP 2025 H1 — 148 cells fixed on top of the LS8 file → cumulative deliverable
`output/Segment_Half_year_version_1_ALL-reconciled.xlsx` + combined report.
Mae's call: SP January untouched (keeps "Corporate Group with Banque" 672 RN
treatment absent from the source). AES/SP 2026: no source provided yet.
`audit_report.py` generalized to render all "Recon …" audit sheets.

## 2026-07-18 — Simple HTML summary of 2026 diffs vs LS8
Recomputed the 2026 H1 (Jan–Jun) differences fresh from the files: 8 cells
differ — RN Online Feb/Mar, Wholesale Feb, overview Mar (±1 each) and revenue
Online Jan/Feb/Mar + ASR Jan low by ฿93,026.53 total. No workbook changes
(2026 stays out of scope per Mae). File: `output/2026_differences_summary.html`.

## 2026-07-18 — Centralized git: created `main` as the single branch
All past work (both old `claude/...` branches) now lives on one central
branch, `main`. New rules in CLAUDE.md: sessions pull `main` at start
(new hook `.claude/hooks/git_session_sync.py`) and land finished work on
`main` at end. Added `/sync` command so Mae can save/share with one word.
Files: CLAUDE.md, WORKLOG.md, `.claude/hooks/git_session_sync.py`,
`.claude/commands/sync.md`, `.claude/commands/guide.md`.

## 2026-07-18 — Set up work log + auto-sync rule
Added this WORKLOG.md, a Stop hook (`.claude/hooks/git_sync_check.py`) that
reminds Claude to commit+push after every task, and the matching rule in
CLAUDE.md. Requested by Mae: always push updates and tell other machines
what's happening via git.

## 2026-07-18 — Fixed corrupt reconciled workbook (desktop Excel)
The reconciled file from the openpyxl round-trip lost chart/comment parts and
desktop Excel reported it as damaged. Rewrote `scripts/reconcile_ls8.py` to
patch cell values surgically into a byte-for-byte copy of the original zip
(same 108 corrections + audit sheet as raw XML). Verified: only intended
parts differ from the source; 2025 totals RN 59,216 / THB 82,620,532.57.
Files: `scripts/reconcile_ls8.py`, `output/Segment_Half_year_version_1_LS8-reconciled.xlsx`.

## 2026-07 (earlier) — First LS8 half-year reconciliation + workspace setup
Reconciled the LYF 2025 block of `Segment_Half_year_version_1.xlsx` against
LS8, built the bilingual HTML audit report, and set up the command suite
(`/guide`, `/reconcile-ls8`, `/check-ls8`, `/audit-report`) and session hook.
