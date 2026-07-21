# Deck: per-delta colors in tile sub-lines

- **Started:** 2026-07-20
- **Requested by:** Mae
- **Status:** done (2026-07-20)

## Goal
Tile bottom lines with two comparisons (vs budget / vs LY) must color each
delta by its own sign — green better, red worse.

## Plan / checklist
- [x] tile() accepts multi-run subs; split all two-delta tiles
- [x] Rebuild, verify every delta segment, deliver, commit & push

## Outcome
tile() now accepts multi-run sub-lines; new subRuns() helper colors each
comparison separately. Split: exec-summary portfolio tiles (occ/ADR/RevPAR),
exec Chinese-arrivals tile, and the OCC/ADR/RevPAR YTD tiles on all 5
performance slides — "x vs budget" and "y vs LY" now green/red
independently ("+1.5% vs LY" shows green even when vs-budget is red).
Verified: 58 signed delta segments across all 44 slides, zero mismatches.
File: output/Portfolio_Performance_Report_H1-2026.pptx (rebuilt).
