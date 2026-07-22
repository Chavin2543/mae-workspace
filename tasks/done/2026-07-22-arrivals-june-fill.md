# Fill June 2026 arrivals into the workbook (MOTS file)

- **Started:** 2026-07-22
- **Requested by:** Mae
- **Status:** done (2026-07-22)

## Goal
Fill June 2026 arrivals: Total, Chinese, India, Middle East, Long-haul
(EU + America) into Summary-arrival / Arrival tabs of the checked workbook.

## Inputs
- data/source/MOTS arrivals by nationality Jun 2026.xlsx (uploaded; June marked "Jun (-4)" = preliminary)

## Plan / checklist
- [x] Verify series match Jan-May (done: total/China/India/EU exact; USA block = America region; ME has small upward revisions Jan-Mar)
- [x] Write June cells + audit sheet
- [x] Read back, commit & push, deliver

## Outcome
June 2026 filled in the checked workbook (7 cells, all previously empty):
Total 1,841,551 (H16) · China 283,297 (H29) · India 163,378 (H42) ·
America 90,980 (H56) · Europe 289,156 (H64) · LH total =H64+H56 (H71) ·
Middle East 52,702 (Arrival W131). Audit sheet "Recon arrivals Jun 2026"
records everything + 3 caveats: June is preliminary ("Jun (-4)"); the USA
block is the America-region series; MOTS revised ME Jan-Mar upward (left
unchanged). China Jan-Jun sum re-checked = source YTD 2,601,609 exactly.
Source filed: data/source/MOTS arrivals by nationality Jun 2026.xlsx.
