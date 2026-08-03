# STR by-segment summary + graphs (Bangkok market classes)

- **Started:** 2026-08-03
- **Requested by:** Mae
- **Status:** done (2026-08-03)

## Goal
Combine the four Bangkok STR "Monthly Performance Data" reports (Luxury,
Upper Upscale, Upscale, Upper Midscale) into one summary file comparing the
market classes by Occupancy, ADR and RevPAR, with easy-to-read graphs.

## Inputs
- data/source/STR Bangkok Luxury monthly 2023-2025.xls
- data/source/STR Bangkok Upper Upscale monthly 2023-2025.xls
- data/source/STR Bangkok Upscale monthly 2023-2025.xls
- data/source/STR Bangkok Upper Midscale monthly 2023-2025.xls

(Each: one sheet "Standard Monthly", Occ/ADR/RevPAR This-Year + %Chg,
Jan 2023–Dec 2025 monthly plus yearly totals.)

## Plan / checklist
- [x] Copy the four uploads into data/source/ with clean names
- [x] Extract Occ/ADR/RevPAR per segment per month into one table
- [x] Build one Excel summary workbook: data tabs + native line charts
- [x] Build an HTML visual report (line charts per metric, 4 segments)
- [x] Verify extracted numbers against the source files
- [x] WORKLOG, commit, push

## Outcome
Delivered:
- `output/STR_Bangkok_by_segment_2023-2025.xlsx` — Overview tab (yearly averages)
  + Occupancy / ADR / RevPAR tabs, each with a native line chart.
- `output/STR_Bangkok_by_segment_2023-2025_report.html` — bilingual EN/TH visual
  report, three interactive line charts, light+dark theme (also published as an
  Artifact).
Scripts: `scripts/str_segment_summary.py`, `scripts/str_report.py`.
Data checked against the STR source "Total" rows (e.g. Luxury 2025 avg
Occ 67.0% / ADR 6,998 / RevPAR 4,691). All 4 files = one sheet "Standard
Monthly", 36 monthly rows Jan 2023–Dec 2025. Figures are STR "This Year"
actuals; the %Chg columns were left out of the summary (kept it to the levels).
