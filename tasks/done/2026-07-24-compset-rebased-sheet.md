# Add re-based STR (compset) data into the Excel workbook

- **Started:** 2026-07-24
- **Requested by:** Mae
- **Status:** done (2026-07-24)

## Goal
New sheet in the workbook holding the re-based CoStar STR data for Bangkok
and Pattaya (Jan-Jun 2026 + derived 2025 + YTD), so the Excel carries the
same numbers as the deck. Old Compset tab untouched (fixed rows feed the
extract; old basis preserved).

## Plan / checklist
- [x] New sheet "Compset re-based (Jun 2026)" with both markets
- [x] Cache restore + read-back verify
- [x] Commit & push, deliver

## Outcome
New sheet "Compset re-based (Jun 2026)" added to the workbook: Bangkok and
Pattaya blocks with Occ/ADR/RevPAR — 2026 Jan-Jun, 2025 derived from the
reports' % change, monthly % chg rows, H1 YTD + % vs LY. Sources noted on
the sheet; old Compset tab untouched (old basis + fixed rows preserved).
Caches restored (8,589); read-back verified. The uploaded pptx copy had no
embedded comments; second Pattaya xls was identical to the filed one.
