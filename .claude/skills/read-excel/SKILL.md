---
name: read-excel
description: MANDATORY whenever opening, inspecting, or extracting data from any Excel workbook (.xlsx) in this workspace — including saying that some data "does not exist" in a file. Read the WHOLE file first - every tab, every row, every column - using scripts/excel_map.py. Triggers on any task that reads a workbook, checks data availability, reconciles, or builds a report from Excel.
---

# Read Excel completely — every tab, full width, full height

## Why this skill exists (Mae's rule, July 2026)

A real mistake happened in this workspace: Claude scanned the `result FY25`
sheet with a reading window that stopped at **column Q**. The monthly P&L data
(Revenue, OPEX, GOP, EBIT, NPAT) started at **column R** — one column further.
Claude told Mae the data "does not exist". Mae knew it did, and she was right.
Never let a partial scan produce a false "the data is not there".

## The rule

1. **Before reading any workbook in detail, map it completely:**

   ```bash
   python3 scripts/excel_map.py <file.xlsx>
   ```

   This prints every sheet, its TRUE size (`max_row x max_column`), and which
   regions contain data. Read the output. If a sheet is 48 columns wide, your
   scans must cover all 48 columns — not the first 15.

2. **Never conclude "the data does not exist" from a partial scan.** Before
   saying anything is missing, search the whole file:

   ```bash
   python3 scripts/excel_map.py <file.xlsx> --find "NPAT"
   ```

   `--find` checks every cell of every sheet. Only after this returns nothing
   may you say the data is not in the file — and even then, say "I could not
   find it" and show what you searched.

3. **When dumping cell ranges by hand, always dump to `ws.max_column` and
   `ws.max_row`**, never a hard-coded guess like `range(1, 18)`. Wide sheets
   in this workspace routinely keep separate data blocks side by side
   (occupancy blocks on the left, P&L blocks starting around column R,
   nationality counts around column AG).

4. **Check every tab**, including ones whose names look irrelevant. In this
   workspace, key data has lived in: `result FY25` / `result FY26` (official
   Occ/ADR/RevPAR + monthly P&L incl. NPAT), `Summary`, `Summary-arrival`
   (arrivals + top-10 nationality), `Compset` (STR, incl. Pattaya at row ~299),
   `Arrival`, and the property tabs `SR9` / `AES` / `LYF` / `SP` (hidden
   RN-by-nationality columns AG..AS on both year blocks).

5. **Formula cells keep stale cached values** after a surgical patch until
   Excel reopens the file — when extracting, prefer the root hardcoded cell
   (`data_only=False` to identify it) or recompute derived values.
