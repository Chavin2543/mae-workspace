# Deck financial slides: fix table overlay + revenue budget recheck

- **Started:** 2026-07-19
- **Requested by:** Mae
- **Status:** done (2026-07-19)

## Goal
1. Monthly P&L table no longer overlaps the footnote on any financial slide.
2. Budget columns show only what the result FY25/FY26 sheets actually record —
   Mae believes there is no revenue budget in the file (verify, fix display).
3. Recheck OPEX / GOP / GOP margin / JV / EBIT / NPAT against the result
   sheets only.

## Inputs
- `output/Segment_Half_year_ALLreconciled_results-checked.xlsx` (result FY25/FY26)
- `scripts/management_deck/build_deck.js`

## Plan / checklist
- [x] Verify which budget rows exist in the result sheets (excel_map + formula read)
- [x] Numeric cross-check every P&L line, all properties, deck vs sheets
- [x] Shift monthly table up / tighten rows so it clears the footnote
- [x] Show "—" for budgets that are not in the file (revenue)
- [x] Rebuild, verify geometry, deliver, WORKLOG, commit & push (branch + main)

## Outcome
**Mae was right about the revenue budget.** The result FY25/FY26 monthly
blocks carry Ascott BP + MF Projection rows ONLY for OPEX, GOP, % GOP margin
and EBIT. There are no budget rows for Total Revenue or NPAT. The YTD block's
"Revenue" BP/MF cells are sheet formulas `= GOP budget − OPEX budget` (and JV
budget `= EBIT budget − GOP budget`). Deck now shows "—" for revenue budgets
(like NPAT); JV kept (derived from two real budget rows) and disclosed in the
footnote.

**Numeric recheck (all pass):** for every property + portfolio, deck values ==
result FY26 cells; YTD actuals == Σ Jan–Jun monthly actuals (Jul–Dec empty);
GOP margin == GOP/Revenue; JV == EBIT − GOP; portfolio NPAT H1 +7.9M.
Financial section uses result FY25/FY26 sheets only ✓.

**Overlay fixed on all 5 financial slides** (20–24): monthly table y 4.62→4.55,
rows 0.20→0.19 with small cell margins, font 8 — table bottom 6.64, footnote
at 6.95. Verified in slide XML.
File: `output/Portfolio_Performance_Report_H1-2026.pptx` (rebuilt).
