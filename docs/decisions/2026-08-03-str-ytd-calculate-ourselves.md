# STR YTD figures: always calculate from raw monthly data, never use STR's YTD rows

- **Date:** 2026-08-03
- **Decided by:** Mae
- **Status:** active

## Context
The Bangkok STR by-segment summary (scripts/str_segment_summary.py +
str_report.py) compares H1 (Jan–Jun) of 2023–2026. STR's reports include a
pre-computed "YTD 2026" row that cannot be reproduced from the monthly rows in
the same file (Bangkok overall: STR YTD Occ 71.69 / ADR 3,758.68 vs
month-calculated 70.52 / 3,952.34; Luxury similar; the other three classes
match). Claude explained the gap as STR aggregating raw room counts with
changing supply; Mae reviewed it and does not trust STR's YTD 2026 row.

## Decision
All YTD/H1 figures in the STR summary — **every year, including 2026** — are
calculated by us from the raw monthly rows: Occ and RevPAR day-weighted,
ADR room-night-weighted (revenue ÷ occupied). STR's own YTD rows and their
% Chg are **not used** anywhere in the deliverables.

## Consequences
- Every year in the comparison uses one identical method (internally
  consistent; Occ × ADR = RevPAR holds everywhere).
- Our 2026 numbers deliberately differ from STR's published YTD row for
  Luxury and Bangkok overall (~1–5%); the outputs carry a note saying the
  figures are calculated from monthly data.
- Bangkok overall has no 2025 monthly data in hand, so it shows no
  vs-2025 change; if Mae uploads Bangkok-overall monthly files for earlier
  years, its change columns can be filled the same way.
- Applies to future STR summary updates unless Mae changes her mind.
