# Deck: remove the Ascott "Budget (BP)" column from financial slides

- **Started:** 2026-07-19
- **Requested by:** Mae
- **Status:** done (2026-07-19)

## Goal
The right-hand P&L table compares to the MF budget only — delete the
Ascott Budget (BP) column on all 5 financial slides.

## Plan / checklist
- [x] Remove the column + widen remaining columns in build_deck.js
- [x] Update slide subtitles/footnote that mention Ascott BP
- [x] Rebuild, verify, deliver, WORKLOG, commit & push (branch + main)

## Outcome
Right-hand P&L table on all 5 financial slides is now 6 columns:
Actual · MF budget · vs MF budget · H1 2025 · vs H1 25 (Ascott Budget (BP)
column deleted; remaining columns widened). Subtitles and footnote no longer
mention Ascott BP. Verified in slide XML on all 5 slides. Extraction still
captures the Ascott BP rows in deck_data.json (unused) in case Mae wants
them back. File: `output/Portfolio_Performance_Report_H1-2026.pptx`.
