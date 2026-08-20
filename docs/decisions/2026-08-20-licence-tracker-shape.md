# Licence & contract tracker: keep the full seed list, slim it from Mae's files

- **Date:** 2026-08-20
- **Decided by:** Mae
- **Status:** active

## Context
`output/License_Contract_Tracker_4properties.xlsx` opens with 222 seeded
checklist rows across the four hotels and 20 columns, all grey ("No date")
until real documents are loaded. Claude asked whether that is too much to
read, and offered to slim the seed list or add a simplified one-page view
before any files arrived.

## Decision
Leave the tracker as built. Do **not** pre-emptively cut the seed list or add
another summary view. Mae uploads her real licence and contract documents
first; the file is then slimmed **from her documents**, not from guesses.

## Consequences
- When the uploads arrive: delete every seeded row that is not a real document
  for that hotel, fill in the real numbers/dates, and hide the columns Mae does
  not type into. The register should end up matching her actual documents.
- Do not re-ask whether to simplify the layout before the files are in hand.
- The Owner / Property defaults in `RESPONSIBILITY` (scripts/build_license_tracker.py)
  are conventional practice, not read from the agreements — Mae still has to
  review that column once. Flag it again when loading her files.
- Rebuilding with `scripts/build_license_tracker.py` starts from empty rows, so
  once Mae has filled anything in, never overwrite her copy with a rebuild.
