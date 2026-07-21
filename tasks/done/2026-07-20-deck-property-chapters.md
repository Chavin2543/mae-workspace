# Deck restructure: per-property chapters (management feedback)

- **Started:** 2026-07-20
- **Requested by:** Mae (management comment)
- **Status:** done (2026-07-20)

## Goal
Merge the Performance, Financial and Segmentation/Nationality sections so
the deck presents property by property: Portfolio chapter first, then one
chapter per hotel (performance -> P&L -> segmentation -> nationality).

## Plan / checklist
- [x] Parameterize slide builders (section eyebrow), wrap loops into functions
- [x] New composition: Sec 3 Portfolio; Sec 4-7 = SR9, AES, LYF, SP chapters
- [x] Update title-slide agenda + dividers
- [x] Rebuild, verify slide order, deliver, WORKLOG, commit & push

## Outcome
Deck restructured from topic sections to property chapters (44 slides):
1 Tourist arrivals · 2 Market STR (unchanged) · 3 Portfolio (performance +
P&L) · 4 SR9 · 5 AES · 6 LYF · 7 SP — each property chapter = divider →
performance vs budget → P&L → segmentation (2 slides) → nationality
(2 slides) → Appendix last. Slide builders parameterized with a section
eyebrow and wrapped as functions (portfolioPerfSlide/propPerfSlide/
portfolioFinSlide/propFinSlide/segSlides/natSlide) with a composition block
at the end of build_deck.js. Slide content itself unchanged. Order verified
from presentation.xml. Title-slide agenda updated.
File: output/Portfolio_Performance_Report_H1-2026.pptx (rebuilt).
