# Adopt Mae workbook v3; rebuild deck; revisit ADR slide

- **Started:** 2026-07-24
- **Requested by:** Mae
- **Status:** done (2026-07-24) — superseded by Mae's rollback request

## Goal
Her uploaded v3 (Excel-recalculated; June in old Compset tab; June revenue
actuals AES/SP; LYF/SP segment refresh) becomes the master workbook.
Rebuild deck from it and reopen the ADR-challenge discussion.

## Plan / checklist
- [x] Copy upload to output (direct copy - keeps Excel caches; no openpyxl save)
- [x] Re-extract + rebuild deck; verify ADR slide figures
- [x] Commit & push; discuss slide 13 with Mae

## Outcome
V3 was adopted and the deck rebuilt from it, but Mae then reported the
file as bugged and asked to roll back. Workbook restored to the pre-compset
version (dd040ca): June arrivals + H1 sums + both audit sheets intact; old
Compset tab June empty; no re-based sheet. Mae's v3 (June revenue actuals
AES/SP, LYF/SP segment refresh, her Compset June entries) remains saved in
git at 51bdfa9 and can be re-applied selectively once the bug is known.
Deck rebuilt from the restored workbook. ADR-slide discussion continues.
