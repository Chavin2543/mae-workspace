# STR summary: add "Change 2023-2025" view (yearly lines + % change)

- **Started:** 2026-08-03
- **Requested by:** Mae
- **Status:** done (2026-08-03)

## Goal
Mae re-uploaded the same four STR files asking to "see change in 2023-2025",
recommending line graphs. Extend the existing summary (same one file) with a
year-over-year change view: yearly-average line graphs per metric and % change
columns/arrows, in both the Excel workbook and the HTML report.

## Inputs
Same four files as tasks/done/2026-08-03-str-segment-summary.md (byte-identical
re-uploads, verified by md5).

## Plan / checklist
- [x] Extract STR's own %Chg for the yearly Total rows
- [x] Excel Overview tab: add 24v23 / 25v24 / 25v23 % columns + yearly line chart per metric
- [x] HTML report: add "Change 2023→2025" section (yearly mini line charts + arrows)
- [x] Verify %chg vs STR published figures; rebuild, screenshot
- [x] WORKLOG, commit, push (branch + main)

## Outcome
Same two output files, extended in place:
- Excel `Overview` tab: yearly averages + "24 vs 23" / "25 vs 24" (STR's own
  %Chg, verified: Luxury Occ +6.9 / −8.7) / "25 vs 23" (computed), green/red,
  plus a yearly line chart per metric (3 native charts added).
- HTML report: new "Change 2023 → 2025" card — yearly mini line charts for
  Occ/ADR/RevPAR with hover tooltips and ▲/▼ % arrows per market class.
Key story: 2024 was the peak for every class; in 2025 Luxury and Upper Upscale
fell below 2023 occupancy (−2.4% / −3.5%) while Upscale (+3.9%) and Upper
Midscale (+5.8%) ended above; ADR rose across all classes, so RevPAR is up
2023→2025 everywhere except Upper Upscale (−0.5%).
