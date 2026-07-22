# Fill June arrivals across all blocks of Summary-arrival + Arrival

- **Started:** 2026-07-22
- **Requested by:** Mae
- **Status:** done (2026-07-22)

## Goal
Every 2026 arrivals row on both sheets gets its June number from the MOTS
June file (series verified by Jan-May match), not just the 5 deck series.

## Plan / checklist
- [x] Scan both sheets for 2026 rows with Jan-May filled, June empty
- [x] Match each series to the MOTS source; fill June where verified
- [x] Append to audit sheet; restore formula caches; verify
- [x] Commit & push, deliver

## Outcome
June 2026 now on BOTH sheets. Arrival: MOTS total H56 (1,841,551), China
H131 (283,297), India W123 (163,378), America AL131 (90,980), EU AL139
(289,156), EU+USA AL147 (=AL139+AL131). Summary-arrival: Middle East X42
(52,702) — completing the earlier 7 cells. Audit sheet extended (second
pass table + not-filled list: AOT block and Bangkok/Pattaya city blocks
use different sources; the Jan-Mar MOTS table is quarterly by design).
Caches restored (8,578) per the new rule; read-back all ✓.

Third pass (same day): all Jan-May YTD sums extended to H1 (Jan-Jun) —
24 SUM formulas widened one column, 9 hardcoded totals (India, ME, LH)
converted to real =SUM(Jan:Jun) formulas, 23 labels renamed to H1; correct
H1 values written into the formula caches (34 + 3 LH chain cells). Deck
updated to match: arrivals arrays 6 months, all YTD tiles/labels now
Jan-Jun with "June preliminary" flags; MOTS YTD -4.9% vs LY (June missing
4 days drags it). STR stays Jan-May (its data really ends in May).
