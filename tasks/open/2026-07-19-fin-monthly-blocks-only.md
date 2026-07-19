# Deck financials: use monthly P&L blocks only (drop YTD summary block)

- **Started:** 2026-07-19
- **Requested by:** Mae
- **Status:** open

## Goal
The financial section must not read the small "YTD performance" summary at the
top of the result sheets at all. Every YTD figure = Jan–Jun sum of the monthly
P&L blocks; budgets shown only where real monthly budget rows exist
(OPEX, GOP, EBIT).

## Inputs
- `output/Segment_Half_year_ALLreconciled_results-checked.xlsx` (result FY25/FY26)
- `scripts/management_deck/` (extract + build)

## Plan / checklist
- [ ] Extend the monthly-block scan to capture Ascott BP / MF Projection rows
- [ ] Compute all YTD act/bp/proj as Jan–Jun sums (no YTD-block reads)
- [ ] GOP margin budget → "—" (needs a revenue budget, which the file lacks)
- [ ] Rebuild, verify values vs sheet sums, deliver
- [ ] Decision log + WORKLOG + commit & push (branch + main)

## Outcome
(filled in when done)
