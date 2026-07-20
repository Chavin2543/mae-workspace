# Q2 2026 performance summary slide (MF Asia format)

- **Started:** 2026-07-20
- **Requested by:** Mae
- **Status:** done (2026-07-20)

## Goal
One-slide Q2 2026 summary like the MF (Asia) Thailand/Serviced Apartments
format: overview bullets + table (Occ, ADR, short-stay revenue share,
nationality YTD) for SR9, AES, LYF, SP — from the workbook in git.

## Inputs
- output/Segment_Half_year_ALLreconciled_results-checked.xlsx (+ deck_data extract)

## Plan / checklist
- [x] Compute Q2 (Apr-Jun) occ/ADR, short-stay revenue share, nationality YTD
- [x] Build slide script (scripts/quarter_summary.py), generate pptx
- [x] Verify numbers, deliver, WORKLOG, commit & push (branch + main)

## Outcome
Built `output/Thailand_SA_Q2-2026_summary.pptx` — one slide in the MF (Asia)
Thailand/Serviced Apartments format: Thailand header + flag, overview bullets
(arrivals −2.3% YTD, Middle East −25%, Chinese +18%, Q2 portfolio OCC 67.5%
vs 72.5% budget / ADR −3.0% vs budget, April = AES & Lyf), and the blue table:
SR9 76.0% / THB 2,781 / short stay 86% / Chinese 49.7% Taiwan 9.0% Thai 5.8%;
AES 60.7% / 3,315 / 79% / Chinese 37.1% Japan 7.2% Finland 5.8%;
LYF 64.7% / 1,220 / 99% / Chinese 21.5% SEA 27% Taiwan 5.5%;
SP 65.6% / 3,471 / 96% / Chinese 24.1% India 8.7% Thai 8.4%.
Q2 = Apr–Jun; occ/ADR day-weighted from the result-sheet monthlies; short
stay = 1 − long-stay share of segment revenue (Q2); nationality = share of
total YTD room nights, SEA grouped (ID/SG/MY) when it outranks countries.
New reusable script: `scripts/quarter_summary.py` (--months, --label).
