# Deck: tile delta colors + EBIT tile

- **Started:** 2026-07-20
- **Requested by:** Mae
- **Status:** done (2026-07-20)

## Goal
Every key-data box: bottom +/- line green when better, red when worse, on
all slides. Financial summary box shows EBIT instead of NPAT.

## Plan / checklist
- [x] Audit all tile calls; make colors sign-based
- [x] finSlide: NPAT tile -> EBIT tile
- [x] Rebuild, verify, deliver, WORKLOG, commit & push

## Outcome
All key-data boxes now color the bottom +/- line by sign: green = better,
red = worse (previously some arrivals tiles were hardcoded GOOD/MUT/BAD).
Changed: exec-summary arrivals tiles, MOTS, Chinese, India, Middle East and
long-haul tiles — now computed from the actual delta. Neutral reference
tiles ("of pre-COVID level") stay grey. Financial summary box now shows
EBIT H1 2026 (vs H1 2025, THB delta) instead of NPAT on all 5 P&L slides —
NPAT stays in both tables. Verified: 40 signed delta texts across all 44
slides, zero color mismatches; 5 EBIT tiles, 0 NPAT tiles.
File: output/Portfolio_Performance_Report_H1-2026.pptx (rebuilt).
