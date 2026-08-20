# Licence tracker checklist comes from the team's handbook (2510_SA list)

- **Date:** 2026-08-20
- **Decided by:** Mae
- **Status:** active (supersedes the "wait for files" step of 2026-08-20-licence-tracker-shape.md)

## Context
Mae uploaded `2510_SA_List_of_Handover_Docs_number_checked_2.xlsx` and said it
is "the handbook for every hotel checklist" — not one hotel's file.

## Decision
The licence tracker's seeded checklist is rebuilt from that handbook's
renewable items (licences, permits, certificates, taxes, insurance, HR
registrations, operating contracts), the same list for all four hotels.
- **Renew by** defaults come from the handbook's own "Kept by" column:
  kept by ATB (property team) → Property; kept by Mitsui/Ananda → Owner.
- One-time archival documents (BOD minutes, SPA, CFA package, construction
  contracts, audited FS) are records, not renewals — excluded from the tracker.
- Known counterparties from the PV work stay pre-filled (laundry per hotel,
  DK Wow F&B at SR9/AES/SP); OTA rows (Booking/Expedia/Agoda) and work permits
  kept even though the handbook doesn't list them.

## Consequences
- `scripts/build_license_tracker.py` constants now mirror the handbook; if the
  team updates the handbook, update the constants from it, not from guesses.
- Still pending from Mae: real reference numbers, issue/expiry dates and fees
  per hotel; and a once-over of the Renew by defaults.
