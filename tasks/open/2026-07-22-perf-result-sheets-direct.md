# Deck perf slides: read result FY26/FY25 directly (Mae priority rule)

- **Started:** 2026-07-22
- **Requested by:** Mae
- **Status:** open

## Goal
Monthly Occ/ADR/RevPAR (2026 act, MF budget, 2025) in the deck must come
straight from result FY26/FY25 sheets, not the Summary tab. Verified diffs:
AES Jan ADR 3,781->3,595; SP Jan ADR 4,064->3,835 (result sheets official).

## Plan / checklist
- [ ] extract_deck_data.py: overwrite perf occ/adr/revpar + bg + ly from result blocks
- [ ] Update appendix source note; rebuild; verify vs result sheets
- [ ] Deliver, WORKLOG, commit & push

## Outcome
(filled in when done)
