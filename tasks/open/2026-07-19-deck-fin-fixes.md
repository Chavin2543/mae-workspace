# Deck financial slides: fix table overlay + revenue budget recheck

- **Started:** 2026-07-19
- **Requested by:** Mae
- **Status:** open

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
- [ ] Verify which budget rows exist in the result sheets (excel_map + formula read)
- [ ] Numeric cross-check every P&L line, all properties, deck vs sheets
- [ ] Shift monthly table up / tighten rows so it clears the footnote
- [ ] Show "—" for budgets that are not in the file (revenue)
- [ ] Rebuild, verify geometry, deliver, WORKLOG, commit & push (branch + main)

## Outcome
(filled in when done)
