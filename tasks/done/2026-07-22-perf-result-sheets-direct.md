# Deck perf slides: read result FY26/FY25 directly (Mae priority rule)

- **Started:** 2026-07-22
- **Requested by:** Mae
- **Status:** done (2026-07-22)

## Goal
Monthly Occ/ADR/RevPAR (2026 act, MF budget, 2025) in the deck must come
straight from result FY26/FY25 sheets, not the Summary tab. Verified diffs:
AES Jan ADR 3,781->3,595; SP Jan ADR 4,064->3,835 (result sheets official).

## Plan / checklist
- [x] extract_deck_data.py: overwrite perf occ/adr/revpar + bg + ly from result blocks
- [x] Update appendix source note; rebuild; verify vs result sheets
- [x] Deliver, WORKLOG, commit & push

## Outcome
Two things done. (1) Repaired the workbook: the June-arrivals openpyxl
save had blanked stored results of external-link/shared formulas; rebuilt
from the pre-June git version, re-applied June + audit sheets, and restored
8,576 cached values with the new scripts/restore_formula_caches.py.
Rule added to CLAUDE.md + decision 2026-07-22-preserve-formula-caches.
(2) Deck Section 3 now reads result FY26/FY25 directly: Occ = Actual Result
rows, ADR = "ADR BF" rows (incl breakfast, Mae's basis; LYF plain ADR),
RevPAR = Occ x ADR, budget = MF Projection rows. Verified equal to the
result sheets for all properties/months. Only visible change: SR9 Jan
occupancy now shows the official 87.02% (was 87.71% RN-derived) and its
RevPAR 2,584 (was 2,605). "Recon result check (Jul 2026)" sheet untouched.
Files: workbook + deck rebuilt.
