# Adopt Mae's Shared workbook as master + re-add Compset H1 column

- **Started:** 2026-08-03
- **Status:** done 2026-08-03

## Outcome

Mae uploaded `..._resultschecked_Shared.xlsx` as the new latest file. Full diff
against the previous master showed it identical everywhere except the Compset
sheet, which was missing the H1 (Jan-Jun) column (she had edited from a
pre-H1 copy). Adopted her file as the master
(`output/Segment_Half_year_ALLreconciled_results-checked.xlsx`), snapshot in
`data/source/Segment_Half_year_ALLreconciled_results-checked_Shared_2026-08-03.xlsx`,
then carried the full column N (223 cells: headers + 180 weighted formulas,
formats, widths) back onto it. Caches restored from her upload (8,834) and the
143 computable H1 cells patched. Verified: N55=76.6, N144=68.8, N206=4,213.7;
zero value differences vs her upload outside Compset.
