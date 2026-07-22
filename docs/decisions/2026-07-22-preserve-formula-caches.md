# Preserve cached formula values on every workbook save

- **Date:** 2026-07-22
- **Decided by:** Claude (confirmed pattern; standing fix for Mae's workbook)
- **Status:** active

## Context
Filling the June 2026 arrivals with a plain openpyxl save silently blanked
the stored results of ~8,600 formula cells (external-link and shared
formulas get an empty `<v/>`). Excel recalculates on open, but every
`data_only=True` read — the deck extraction, verification checks — broke.

## Decision
Any openpyxl save of a delivered workbook must be followed by
`scripts/restore_formula_caches.py <last-good.xlsx> <saved.xlsx>` (donor =
previous version from git), then a read-back check of a known formula cell.
The deck extraction reads Occ/ADR/RevPAR straight from the result sheets
(ADR BF basis for SR9/AES/SP; LYF plain ADR; FY25/FY24 plain ADR already
include breakfast).

## Consequences
- Never trust a freshly saved workbook without the restore + read-back.
- Stale caches on recalculated cells are expected — note them in the audit.
- The pre-save version must exist in git first (commit before editing).
