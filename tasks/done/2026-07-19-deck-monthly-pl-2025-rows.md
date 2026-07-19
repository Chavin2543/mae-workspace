# Deck: add OPEX/GOP/EBIT/NPAT ’25 rows to monthly P&L tables

- **Started:** 2026-07-19
- **Requested by:** Mae
- **Status:** done (2026-07-19)

## Goal
On every financial-section slide of the management deck, the monthly P&L
table should show 2025 next to 2026 for every line (Revenue, OPEX, GOP,
EBIT, NPAT) so the months can be compared year over year.

## Inputs
- `output/Segment_Half_year_ALLreconciled_results-checked.xlsx` (result FY25/FY26 sheets)
- `scripts/management_deck/` (extract_deck_data.py → deck_data.json → build_deck.js)

## Plan / checklist
- [x] Add ’25 rows (OPEX, GOP, EBIT, NPAT) paired with their ’26 rows in build_deck.js
- [x] Adjust layout so the taller table still fits the slide
- [x] Rebuild the deck and verify the financial slides (XML-level: 11-row table,
      geometry fits above the footnote; values cross-checked vs deck_data.json —
      LibreOffice cannot render pptx in this container, so no screenshot)
- [x] Deliver to Mae, WORKLOG, commit & push (branch + main)

## Outcome
All 5 financial slides (Portfolio, SR9, AES, LYF, SP — slides 20–24) now show
the monthly P&L table with each line paired ’26/’25: Revenue, OPEX, GOP, EBIT,
NPAT (10 data rows). ’25 rows are muted grey; zebra bands group each pair.
Chart shrunk slightly (h 3.0→2.55) and row height 0.25→0.20 so it fits.
Spot-check: portfolio NPAT ’25 H1 = −12.6M ✓ (matches result FY25).
Also fixed `extract_deck_data.py` to take the output path as an argument
(was hardcoded to an old session's scratchpad).
File: `output/Portfolio_Performance_Report_H1-2026.pptx` (rebuilt).
