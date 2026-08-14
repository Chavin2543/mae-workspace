# Work log

Short notes on what each session did, newest first — so any machine/session
can see the state of the workspace from git alone. One entry per finished
task: date, what was done, files touched. Keep entries to 2–4 lines, plain
language.

---

## 2026-08-14 — LS8 payment approval tables (4 payments)
- Read 4 uploaded payment documents for lyf Sukhumvit 8 / AMH Sukhumvit 8 Co., Ltd:
  2 Ananda PVs (PV-009 remittance to owner 1.5M THB, PV-010 Ascott fees
  684,677.08 THB) + 2 SCB international transfer applications (Adria Scan USD 584,
  Day Use HK HKD 5,434.78). All amounts cross-checked against their invoices.
- Filed the PDFs in `data/pdf/` (PV_*.pdf); the approval tables live in chat only.

## 2026-08-14 — AES payment approval table (5 PVs)
- Read 5 Ananda PVs for Ascott Embassy Sathorn / AMH Sathorn Co., Ltd.:
  PV-072 Expedia 684,449.48, PV-073 Booking.com 1,611,752.05, PV-075 AIMT
  professional fees 807,240.13, PV-076 AIMT central fee 997,021.75, PV-084
  DK Wow F&B 918,755.50 — total 5,019,218.91 THB, all cross-checked to invoices.
- PDFs filed in `data/pdf/`; the approval table lives in chat only.

## 2026-08-14 — SR9 payment approval table (6 PVs + audit confirmation)
- Read the "PV 08.26 over 500k due 15.08.26" bundle for Somerset Rama 9 / AMH
  Ratchada Co., Ltd.: PV-028 remittance to owner 18,700,000, PV-029 AIMT fees
  2,951,556.90, PV-030 DK Wow 1,228,618.00, PV-031 Asset World Wex laundry
  681,784.22, PV-032 Expedia 778,203.80, PV-033 Booking.com 683,966.56 —
  total 25,024,129.48 THB, matching Mae's cheque summary exactly.
- Also a BJC audit balance confirmation (94,352.60 THB as at 30 Jun 2026) —
  stamp/sign only, not a payment. PDFs filed in `data/pdf/`.

## 2026-08-14 — SP payment approval table (5 PVs)
- Read 5 PVs for Somerset Pattaya / AMH Pattaya Company Limited: PV-005 AIMT
  central + professional fees 1,800,324.66, PV-012 PEA electricity 1,447,431.37,
  PV-013 DK Wow F&B 1,161,117.00, PV-014 Booking.com 1,329,971.58, PV-015
  Laundry Pattaya 531,243.44 — total 6,270,088.05 THB, all cross-checked.
- Second batch: PV-016 Expedia compensation 516,472.75, PV-017 returning excess
  fund to owner 19,500,000.00, plus Travelscape LLC USD 1,500 international
  transfer and cover Memo SP26-08-001 (memo confirms PV-005 number is 2608).
  SP grand total 26,286,560.80 THB + 1,500.00 USD.
- PDFs filed in `data/pdf/`; the approval table lives in chat only.

## 2026-08-03 — Pattaya STR summary (H1 2025 vs H1 2026), like the Bangkok one
- New sources: STR Pattaya market + Upscale&UpperMid YTD Jun 2026 (both carry
  This Year 2026 AND Last Year 2025 monthly Jan-Jun).
- New scripts str_pattaya_summary.py / str_pattaya_report.py; outputs
  STR_Pattaya_by_market_2025-2026.xlsx + _report.html. All H1 figures
  calculated from raw months per the standing YTD decision.

## 2026-08-03 — STR summary: Bangkok overall now a full series (2024-2026)
- Mae uploaded Bangkok whole-market monthly data 2024-2025; filed as
  data/source/STR Bangkok market monthly 2024-2025.xls (no 2023 data exists).
- Bangkok overall joined everywhere as a gray dashed reference line: monthly
  tabs + charts, H1 tables (H1 2024/2025/2026, 2023 shown as em-dash), YTD
  block %chg now filled, report tiles/charts/tables. All calculated from raw
  monthly rows per the standing decision.

## 2026-08-03 — STR summary: added "26 vs 24" change column
- Overview H1 tables now have five %-change columns (24v23, 25v24, 26v25,
  26v24, 26v23); report change section shows "vs H1 2024" and "vs H1 2023"
  arrow columns per metric.

## 2026-08-03 — STR summary: all YTD figures now calculated from raw monthly data
- Mae's decision (docs/decisions/2026-08-03-str-ytd-calculate-ourselves.md):
  she does not trust STR's pre-computed YTD 2026 row, so every H1/YTD figure -
  every year including 2026 - is calculated from the monthly rows
  (Occ/RevPar day-weighted, ADR room-night-weighted). STR YTD rows unused.
- Effect: Luxury H1-26 now 66.8/7,082/4,729 (was STR's 67.2/6,845/4,600);
  Bangkok overall 70.5/3,952/2,787, no vs-2025 arrows (no 2025 monthly data).

## 2026-08-03 — STR summary: comparison switched to H1-only (Mae's fix)
- Mae: 2026 was missing from the Overview change view, and she compares H1
  (Jan-Jun) of each year only. Change view + bottom tables now use H1 2023,
  H1 2024, H1 2025, H1 2026 everywhere (2026 fully included).
- H1 2023-2025 are day-weighted from monthly data; H1 2026 uses STR's own YTD
  row -- for Luxury the day-weighted calc differs (ADR 7,082 vs official 6,845;
  supply changed mid-year), so the official row wins. Noted in both outputs.

## 2026-08-03 — STR summary extended with 2026 YTD (Jan-Jun) + Bangkok overall
- Mae uploaded five YTD 2026 STR files: the four classes + a new whole-market
  "Bangkok" file. Filed as data/source/STR Bangkok * YTD Jun 2026.xls.
- Monthly series now runs Jan 2023 - Jun 2026 (42 months) in both outputs.
- Overview tab + report tiles: "YTD 2026 vs same period 2025" comparison with
  STR's own %Chg, incl. Bangkok overall as reference. Yearly table gets a YTD
  column. Outputs renamed to STR_Bangkok_by_segment_2023-2026.* (old
  ..._2023-2025.* removed).

## 2026-08-03 — Payment approval tables (AMH Sathorn + AMH Ratchada)
Ran payment-approval-table on 2 vouchers: AMH Sathorn KP 2600000032 (land &
building tax 2026, 3,262,250.00 THB to T C Sathorn Condo) and AMH Ratchada
KP 2600000040 (IAR/BI/PV insurance premiums Jun 26–Jun 27, 1,384,670.58 THB
to Rabbit Care Broker). Tables delivered in chat; PDFs filed in data/pdf/.

## 2026-08-03 — Deck v11 reshaped with Mae (ADR story + slimmer slides)
Adopted Mae's v11 as master, then per her comments: performance pages show
actual/MF Projection/actual-2025 side by side; new ADR-by-segment slides
(portfolio H1 + per-property monthly 26 vs 25); nationality pairs merged;
segmentation tables show 26 vs 25 with change; P&L tables swap 2024 rows for
Budget rows; property sections reordered (performance -> nationality ->
segments -> ADR -> P&L). 48 slides. Data checks: Chinese demand UP +14.8%,
mix effect negligible — ADR drop = in-segment price cuts, deepest in Online.

## 2026-08-03 — Deck v11 adopted; performance pages get MF Projection + 2025 rows
Mae's v11 pptx (51 slides) is the new master. Performance tables (Portfolio,
SR9, AES, LYF, SP) rebuilt: each metric now shows actual 2026 / MF Projection /
actual 2025 side by side, monthly + weighted YTD, from result FY26/FY25.
Restructure plan + ADR narrative proposed to Mae (task still open).

## 2026-08-03 — STR summary: added "Change 2023-2025" view
- Mae re-uploaded the same four STR files (md5-identical) asking to see the
  change over 2023-2025 with line graphs.
- Overview tab now: yearly averages + STR's own YoY %Chg (24v23, 25v24) +
  computed 25v23, colored, with a yearly line chart per metric.
- HTML report: new "Change 2023 → 2025" section — three yearly mini line
  charts with hover + up/down arrows per market class. Same output filenames.

## 2026-08-03 — STR by-segment summary + graphs (Bangkok market classes)
- Combined the four uploaded Bangkok STR "Monthly Performance Data" reports
  (Luxury, Upper Upscale, Upscale, Upper Midscale) into one summary.
- New scripts: scripts/str_segment_summary.py (reads the 4 .xls, builds an Excel
  workbook — Overview + Occupancy/ADR/RevPAR sheets, each with a native line
  chart) and scripts/str_report.py (bilingual HTML report, 3 interactive line
  charts, light+dark).
- Deliverables: output/STR_Bangkok_by_segment_2023-2025.xlsx and
  ..._report.html. Sources filed in data/source/. Data verified vs the source
  Total rows.

## 2026-08-03 — Mae's Shared workbook adopted as master (H1 column re-added)
Mae uploaded her "Shared" copy as the new latest file. Diff showed it identical
to the previous master except the Compset H1 column was missing (edited from a
pre-H1 copy) — adopted her file, carried the H1 column (223 cells) back on,
restored caches, verified values. Snapshot in data/source/, master in output/.

## 2026-07-30 — Compset sheet: H1 (Jan-Jun) column with live formulas
Per Mae: column N of the Compset sheet now carries an "H1 (Jan-Jun)" column —
180 formulas across every block (Bangkok, Rachada, Sathorn, Nana, SP, Pattaya).
Occ/RevPar are day-weighted, ADR is room-night-weighted (same math as the deck
slides); partial years stay blank. Caches restored + patched; verified SP comp
ADR H1 = 4,214 matching slide 18. File: the results-checked master workbook.

## 2026-07-24 — New skill: payment vouchers → approval-email tables
Per Mae: she uploads ~4 payment vouchers (different companies) and gets one
copy-pasteable table per company (No. / Detail / Payment Voucher; Description,
Pay to, Amount) for her approval email. New `payment-approval-table` skill;
task anatomy added to CLAUDE.md. Format confirmed against the AMH Ratchada
VAT voucher sample.

## 2026-07-23 — New workflow: property email answers → Excel notes
Per Mae: when a property answers her monthly FS questions by email, the answer
is marked as an Excel note on that account's row in that month's column.
New `answer-note` skill + `scripts/mark_answer_note.py` (append-only, tested);
decision logged, CLAUDE.md documents the full questions→answers loop.

## 2026-07-22 — Arrivals YTD extended to H1 (Jan-Jun) in workbook + deck
Per Mae: every Jan-May YTD sum now covers Jan-Jun — 24 formulas widened,
9 hardcoded totals became real =SUM formulas, labels renamed to H1; caches
carry correct H1 values. Deck arrivals section now shows Jan-Jun with
"June preliminary" flags (MOTS YTD -4.9% vs LY). ME confirmed filled on
both sheets (52,702); only AOT + city blocks await other sources.

## 2026-07-22 — June arrivals completed on both sheets (Arrival too)
Per Mae: filled June 2026 into every remaining block — Arrival sheet (MOTS
total, China, India, America, EU, EU+USA formula) and the Summary-arrival
Middle East block. AOT + Bangkok/Pattaya city blocks left (different
sources, noted in audit). Caches restored after save; all reads verified.

## 2026-07-22 — Workbook cache repair + deck reads result sheets directly
The June-arrivals save had blanked ~8.6k stored formula results (openpyxl
drops external-link/shared caches). Rebuilt from git + new
scripts/restore_formula_caches.py; rule in CLAUDE.md + decision logged.
Deck Occ/ADR/RevPAR now come straight from result FY26/FY25 (ADR BF basis,
MF Projection budget) — only visible change: SR9 Jan occ 87.71%→87.02%
(official). Verified all months vs result sheets.

## 2026-07-22 — June 2026 arrivals filled (Total/China/India/ME/long-haul)
Mae uploaded the MOTS June nationality file (filed in data/source/). Filled
June into Summary-arrival (total, China, India, Europe, America, LH formula)
and Arrival (Middle East) — 7 cells, audit sheet "Recon arrivals Jun 2026".
"Jun (-4)" = 4 days lost to a system error: cells keep reported figures;
audit sheet carries the pro-rated full-June estimates (total ~2.12M,
range 2.12-2.27M). ME Jan-Mar revisions noted, not applied.

## 2026-07-20 — Deck: two-comparison tile lines split into per-sign colors
Mae caught that "x vs budget · y vs LY" lines used one color for both parts.
tile() now takes multi-color runs (subRuns helper); each delta is green/red
by its own sign on the exec-summary and all performance-slide tiles.
58 delta segments verified. Files: output deck + build_deck.js.

## 2026-07-20 — Deck: sign-based tile colors + EBIT summary box
Per Mae: every key-data box's bottom +/- line is now green when better, red
when worse (arrivals tiles were partly hardcoded); financial slides' summary
box shows EBIT instead of NPAT (NPAT stays in the tables). Verified 40
signed deltas across all slides. Files: output deck + build_deck.js.

## 2026-07-20 — Deck restructured into property chapters (management comment)
Management wants the story per property: deck now runs Arrivals · Market ·
Portfolio (perf+P&L) · then one chapter per hotel (SR9, AES, LYF, SP), each
with performance, P&L, segmentation and nationality slides together.
44 slides; content unchanged, order+sections reworked in build_deck.js.

## 2026-07-20 — New: Q2 2026 one-slide summary (MF Asia format)
Per Mae: built output/Thailand_SA_Q2-2026_summary.pptx — overview bullets +
blue table (Occ, ADR, short-stay revenue share, nationality YTD) for the four
hotels, Q2 = Apr–Jun, all figures from the reconciled workbook via
deck_data. New reusable script scripts/quarter_summary.py.

## 2026-07-20 — Deck: vs H1 24 column, MOTS COVID/peak marks, April note
Per Mae: financial tables got a "vs H1 24" delta column; the MOTS arrivals
table grays out 2020/2021 with a (COVID) label and highlights the 2019 peak
total 39,916,251 in yellow; portfolio slide note now attributes April
softness to AES + LYF rather than the market (Bangkok compset was flat).
Files: output deck + scripts/management_deck/build_deck.js.

## 2026-07-19 — Deck: 2024 added to STR, Section 3 and Section 4
Mae uploaded result FY24 (filed: data/source/result FY24.xlsx). STR charts
got a 2024 series; Section 3 slides a 2024 RevPAR line + vs-24 row; financial
slides ’24 monthly P&L rows + H1 2024 column (monthly blocks only, ATB
excluded from portfolio, leap-year Feb handled). All values verified vs the
FY24 sheet. Files: output deck + scripts/management_deck/.

## 2026-07-19 — Deck: data labels on Section 5 charts
Per Mae: added value labels to the segment stacked bars and the nationality
H1 comparison bars on all four property slides (zeros hidden to keep them
readable). Files: output deck + scripts/management_deck/build_deck.js.

## 2026-07-19 — Deck: removed the Ascott Budget (BP) column
Per Mae: the financial P&L tables compare to the MF budget only. Deleted the
Budget (BP) column on all 5 financial slides; subtitles/footnote updated.
Files: output deck + scripts/management_deck/build_deck.js.

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
