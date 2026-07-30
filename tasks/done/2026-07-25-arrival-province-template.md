# Fill Arrival province 2019-2026 template (BKK, Phuket, Pattaya)

- **Started:** 2026-07-25
- **Requested by:** Mae
- **Status:** done (2026-07-25)

## Goal
Fill Mae's "Arrival province 2019-2026" template with domestic/foreign
arrivals by month: Bangkok, Phuket + new Chonburi(Pattaya) block, from the
MOTS provincial file (2025+2026 Jan-Jun; 2019/2024 not in source).

## Inputs
- data/source/MOTS provincial arrivals Jan-Jun 2026.xlsx
- Mae template upload ac3a800a

## Plan / checklist
- [x] Extract Thai/foreign per province per month (2569P + 2568R cols)
- [x] Fill template + add Pattaya block; deliver to output/
- [x] Commit & push

## Outcome
Filled output/Arrival province 2019-2026.xlsx: Bangkok + Phuket blocks
(2025 + 2026, Jan-Jun, Domestics = คนไทย / International = ต่างชาติ) and a
new Pattaya (Chonburi) block added in the same format (rows 46-63, with the
same SUM formulas). 2019 and 2024 left empty — not in the source file; the
source note is written at the bottom of the sheet. Jan-Mar 2026 are final
(R), Apr-Jun preliminary (p). Read-back verified against the source.
Source filed: data/source/MOTS provincial arrivals Jan-Jun 2026.xlsx.
