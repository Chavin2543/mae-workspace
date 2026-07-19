# Deck: add OPEX/GOP/EBIT/NPAT ’25 rows to monthly P&L tables

- **Started:** 2026-07-19
- **Requested by:** Mae
- **Status:** open

## Goal
On every financial-section slide of the management deck, the monthly P&L
table should show 2025 next to 2026 for every line (Revenue, OPEX, GOP,
EBIT, NPAT) so the months can be compared year over year.

## Inputs
- `output/Segment_Half_year_ALLreconciled_results-checked.xlsx` (result FY25/FY26 sheets)
- `scripts/management_deck/` (extract_deck_data.py → deck_data.json → build_deck.js)

## Plan / checklist
- [ ] Add ’25 rows (OPEX, GOP, EBIT, NPAT) paired with their ’26 rows in build_deck.js
- [ ] Adjust layout so the taller table still fits the slide
- [ ] Rebuild the deck and visually check a financial slide
- [ ] Deliver to Mae, WORKLOG, commit & push (branch + main)

## Outcome
(filled in when done)
