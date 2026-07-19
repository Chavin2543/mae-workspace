# Deck financials: use monthly P&L blocks only (drop YTD summary block)

- **Started:** 2026-07-19
- **Requested by:** Mae
- **Status:** done (2026-07-19)

## Goal
The financial section must not read the small "YTD performance" summary at the
top of the result sheets at all. Every YTD figure = Jan–Jun sum of the monthly
P&L blocks; budgets shown only where real monthly budget rows exist
(OPEX, GOP, EBIT).

## Inputs
- `output/Segment_Half_year_ALLreconciled_results-checked.xlsx` (result FY25/FY26)
- `scripts/management_deck/` (extract + build)

## Plan / checklist
- [x] Extend the monthly-block scan to capture Ascott BP / MF Projection rows
- [x] Compute all YTD act/bp/proj as Jan–Jun sums (no YTD-block reads)
- [x] GOP margin budget → "—" (needs a revenue budget, which the file lacks)
- [x] Rebuild, verify values vs sheet sums, deliver
- [x] Decision log + WORKLOG + commit & push (branch + main)

## Outcome
`extract_deck_data.py` no longer reads the "YTD performance" summary block at
all — the monthly-block scan now also captures the Ascott BP / MF Projection
rows (block-boundary guard added so the portfolio block cannot swallow SR9's
budget rows), and all YTD figures are Jan–Jun sums. Portfolio OPEX budgets =
sum of the four properties (FY26 portfolio block has no OPEX rows). Deck right
table verified on all 5 slides against independent sheet sums: budgets shown
for OPEX/GOP/EBIT (+JV = EBIT−GOP), "—" for Revenue, GOP margin, NPAT.
Values match the previous deck (the old YTD block summed the same months) —
the change is provenance, plus GOP margin budget now honestly "—".
Decision logged: `docs/decisions/2026-07-19-fin-monthly-blocks-only.md`.
File: `output/Portfolio_Performance_Report_H1-2026.pptx` (rebuilt).
