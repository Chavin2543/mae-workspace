# Financial figures: monthly P&L blocks only, never the YTD summary block

- **Date:** 2026-07-19
- **Decided by:** Mae
- **Status:** active

## Context
The result FY25/FY26 sheets have two areas: the monthly P&L blocks (label
col R, months T..AE) and a small "YTD performance" summary at the top. The
deck's financial section originally read the top summary block, whose
Revenue and JV "budget" cells are derived formulas (Revenue = GOP − OPEX
budgets; JV = EBIT − GOP budgets), not real budget lines.

## Decision
Never use the "YTD performance" summary block at the top of the result
sheets. All financial figures (deck or otherwise) come from the monthly
P&L blocks only; YTD = Jan–Jun sums. Budgets exist only as Ascott BP /
MF Projection rows for OPEX, GOP, % GOP and EBIT — show "—" for any line
without a real budget row (Revenue, NPAT; GOP margin YTD budget would need
a revenue budget, so it is "—" too). JV expense = EBIT − GOP (disclosed in
the slide footnote).

## Consequences
- `scripts/management_deck/extract_deck_data.py` computes `data.fin` from
  the monthly blocks; do not reintroduce reads of the top block.
- Never invent or derive a revenue budget. If Mae supplies a real revenue
  budget source later, add it as a new input, not from the YTD block.
