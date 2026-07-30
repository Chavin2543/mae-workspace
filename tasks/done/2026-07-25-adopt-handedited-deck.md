# Adopt Mae hand-edited pptx as deck master

- **Started:** 2026-07-25
- **Requested by:** Mae
- **Status:** done (2026-07-25)

## Goal
Her 50-slide edit becomes the working master. Future deck changes are
surgical pptx edits - build_deck.js must NOT be re-run over it.

## Plan / checklist
- [x] Copy upload to output/
- [x] Guard build_deck.js with a warning header; decision log
- [x] Commit & push

## Outcome
Adopted (both uploads identical; used the newer). 50 slides: 4 new (BKK +
Pattaya "Tourist Arrivals + Spending power" titles with blank content
slides at 10-13) and "Us vs our compsets" moved before the ADR slide.
build_deck.js guarded with a do-not-rerun header; decision logged
(2026-07-25-deck-hand-edited-master).
