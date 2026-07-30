# Compset sheet: H1 column with weighted formulas

- **Started:** 2026-07-27
- **Status:** done 2026-07-30

## Outcome

Added a live "H1 (Jan-Jun)" column (col N) to the Compset sheet of
`output/Segment_Half_year_ALLreconciled_results-checked.xlsx` — 180 formulas
across all blocks (Bangkok, Rachada, Sathorn, Nana, SP, Pattaya market),
replacing Mae's earlier plain =AVERAGE() cells:

- Occupancy: day-weighted `SUMPRODUCT(B:G,{31,28,31,30,31,30})/181`
- ADR: room-night-weighted `SUMPRODUCT(adr,occ,days)/SUMPRODUCT(occ,days)`
  (each ADR row paired with its own year's occupancy row)
- RevPar: day-weighted, same as occupancy
- Blank result when the six months are not all present (partial years)

Formulas recalculate automatically in Excel. After the openpyxl save, restored
8,834 pre-existing formula caches from the git donor and patched cached values
into the 143 computable new H1 cells; verified N206 (SP comp ADR) = 4,213.7
matching slide 18's ฿4,214, and all non-Compset cells byte-identical to the
pre-edit file.
