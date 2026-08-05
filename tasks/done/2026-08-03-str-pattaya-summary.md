# STR Pattaya summary + graphs (like the Bangkok one)

- **Started:** 2026-08-03
- **Requested by:** Mae
- **Status:** done (2026-08-03)

## Goal
Same deliverable as the Bangkok STR summary, for Pattaya: one Excel file +
visual HTML report with line graphs, comparing H1 2025 vs H1 2026 for
Pattaya Area overall and Pattaya Upscale & Upper Mid.

## Inputs
- data/source/STR Pattaya market YTD Jun 2026.xls (submarket, This Year + Last Year)
- data/source/STR Pattaya Upscale UpperMid YTD Jun 2026.xls (submarket class)
Both: monthly Jan–Jun 2026 + Last Year (Jan–Jun 2025) + YTD row.

## Plan / checklist
- [x] Script scripts/str_pattaya_summary.py (Excel) + str_pattaya_report.py (HTML)
- [x] H1 figures calculated from raw monthly rows (standing decision — STR YTD rows unused)
- [x] Monthly Jan–Jun charts: 2025 dashed vs 2026 solid, both series
- [x] Verify vs sources; screenshot; WORKLOG, commit, push (branch + main)

## Outcome
Delivered output/STR_Pattaya_by_market_2025-2026.xlsx (Overview + 3 monthly
tabs, native charts) and ..._report.html (bilingual, interactive). Calculated
H1s cross-checked against STR YTD rows (within 0.2pp/6 THB, method noted).
Story: Pattaya H1 2026 is WEAKER than H1 2025 — overall Occ 64.5 (-2.4%),
RevPAR -2.3%; the Upscale & Upper Mid class fell harder (Occ -3.1%,
RevPAR -5.0%), with May-Jun the weak months (Occ ~49-52%).
