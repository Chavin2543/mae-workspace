#!/usr/bin/env python3
"""Reconcile the lyf Sukhumvit 8 (LS8) block of the Segment Half-year workbook
against the LS8 Market Segment source workbook.

LS8 = "lyf Sukhumvit 8 Bangkok" = the `LYF` tab of the half-year workbook.
The LS8 workbook is the authoritative source; this script overwrites the
hardcoded input cells in the half-year workbook so they match LS8, and adds
an audit sheet listing every change (old -> new).

Scope (current run): 2025 comparison block only.
  - LYF!C51:N51            Room Nights (monthly overview, w/o COMP/HU/SYS)
  - LYF!C58:N64            Room Nights by Segment (7 segments)
  - LYF!C68:N74            Revenue by Segment (7 segments)
  - Summary!C60:N61        LYF 2025 Occupancy % / ADR  (feeds LYF!C52:N53;
                           Revpar row 62 is =Occ*ADR formulas and recalculates)
Not touched (documented in the audit sheet):
  - Summary!C63:N63 "Revenue (THB)" - consistently ~3-4% above LS8 room
    revenue for both 2025 and 2026 -> different basis (not room revenue only),
    so it cannot be reconciled to LS8.
  - The 2026 H1 block (out of scope for this run; known diffs are listed in
    the audit sheet as informational notes).

Usage:
  python3 scripts/reconcile_ls8.py \
      --ls8 data/source/LS8_Market_Segment_2025_YTD_2026.xlsx \
      --halfyear data/source/Segment_Half_year_version_1.xlsx \
      --out output/Segment_Half_year_version_1_LS8-reconciled.xlsx
Then recalculate with the xlsx skill's recalc script (LibreOffice).
"""
import argparse
import datetime as dt

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

# --- LS8 source sheet layout ("LS8 Segment 2025"), months in cols D..O -------
LS8_SHEET_2025 = "LS8 Segment 2025"
LS8_COL0 = 4  # column D = Jan

# segment label in LS8 col B -> (RNs row, Revenue row)
LS8_SEGMENT_ROWS = {
    "Corporate SS":    (7, 9),
    "LS":              (10, 12),
    "Online":          (13, 15),
    "ASR":             (16, 18),
    "Wholesale":       (19, 21),
    "Group Corporate": (22, 24),
    "Group Leisure":   (25, 27),
    "Group Series":    (28, 30),
    "Employee Travel": (31, 33),
}
LS8_ACTUAL_ROOMS_ROW = 48   # "Actual Occupied Rooms  w/o COMP, HU, SYS"
LS8_ACTUAL_OCC_ROW = 49     # "Actual OCC (%)"
LS8_ACTUAL_ADR_ROW = 50     # "Actual ADR"

# --- Half-year workbook layout (2025 block), months in cols C..N -------------
HY_COL0 = 3  # column C = Jan

# LYF row -> (label in LYF col B, list of LS8 segment labels summed)
# "Corporate Group" in LYF aggregates the three LS8 group segments.
LYF_RN_ROWS = {
    58: ("Corporate Short Stay", ["Corporate SS"]),
    59: ("Corporate Long Stay", ["LS"]),
    60: ("Online Business (Dynamic Rate)", ["Online"]),
    61: ("ASR", ["ASR"]),
    62: ("Wholesale (Static Rate)", ["Wholesale"]),
    63: ("Corporate Group", ["Group Corporate", "Group Leisure", "Group Series"]),
    64: ("Employee Travel", ["Employee Travel"]),
}
LYF_REV_ROWS = {r + 10: v for r, v in LYF_RN_ROWS.items()}  # 68..74, same labels
LYF_OVERVIEW_RN_ROW = 51
SUMMARY_OCC_ROW, SUMMARY_ADR_ROW = 60, 61  # row 62 Revpar = formulas, untouched

AUDIT_SHEET = "Recon LS8 (2025)"


def month_values(ws, row, col0=LS8_COL0):
    return [ws.cell(row, col0 + i).value or 0 for i in range(12)]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ls8", required=True)
    ap.add_argument("--halfyear", required=True)
    ap.add_argument("--out", required=True)
    args = ap.parse_args()

    src = openpyxl.load_workbook(args.ls8, data_only=True)
    ls8 = src[LS8_SHEET_2025]

    # sanity-check the LS8 layout before trusting hardcoded row numbers
    assert ls8["B4"].value == 2025, f"LS8 sheet year != 2025: {ls8['B4'].value}"
    for label, (rn_row, rev_row) in LS8_SEGMENT_ROWS.items():
        assert ls8.cell(rn_row, 2).value == label, \
            f"LS8 B{rn_row} expected {label!r}, got {ls8.cell(rn_row, 2).value!r}"
        assert ls8.cell(rev_row, 3).value == "Revenue"
    assert "Actual Occupied Rooms" in str(ls8.cell(LS8_ACTUAL_ROOMS_ROW, 2).value)
    assert "Actual OCC" in str(ls8.cell(LS8_ACTUAL_OCC_ROW, 2).value)
    assert "Actual ADR" in str(ls8.cell(LS8_ACTUAL_ADR_ROW, 2).value)

    # pull authoritative monthly vectors from LS8
    seg_rn = {k: month_values(ls8, v[0]) for k, v in LS8_SEGMENT_ROWS.items()}
    seg_rev = {k: month_values(ls8, v[1]) for k, v in LS8_SEGMENT_ROWS.items()}
    actual_rooms = month_values(ls8, LS8_ACTUAL_ROOMS_ROW)
    actual_occ = month_values(ls8, LS8_ACTUAL_OCC_ROW)
    actual_adr = month_values(ls8, LS8_ACTUAL_ADR_ROW)

    # target workbook: keep formulas -> data_only=False
    wb = openpyxl.load_workbook(args.halfyear, data_only=False)
    lyf, summary = wb["LYF"], wb["Summary"]
    assert summary["A49"].value == "LYF" and summary["B50"].value == 2026
    assert summary["B59"].value == 2025, "Summary LYF 2025 header moved"

    changes = []  # (sheet, cell, item, month, old, new)

    def set_cell(ws, row, col, new, item, month):
        cell = ws.cell(row, col)
        old = cell.value
        if isinstance(old, str) and old.startswith("="):
            return  # never overwrite an existing formula
        if old is None:
            old = 0
        if isinstance(new, float):
            new = round(new, 10)
        if (isinstance(old, (int, float)) and isinstance(new, (int, float))
                and abs(float(old) - float(new)) < 1e-6):
            return
        cell.value = new
        changes.append((ws.title, f"{get_column_letter(col)}{row}",
                        item, month, old, new))

    # 1) RN by Segment + 2) Revenue by Segment
    for rows_map, data, kind, nd in ((LYF_RN_ROWS, seg_rn, "RN", 0),
                                     (LYF_REV_ROWS, seg_rev, "Revenue", 2)):
        for lyf_row, (label, src_labels) in rows_map.items():
            got = lyf.cell(lyf_row, 2).value
            assert got == label, f"LYF B{lyf_row} expected {label!r}, got {got!r}"
            vec = [round(sum(data[s][i] for s in src_labels), nd)
                   for i in range(12)]
            for i, v in enumerate(vec):
                v = int(v) if nd == 0 else v
                set_cell(lyf, lyf_row, HY_COL0 + i, v,
                         f"{label} {kind}", MONTHS[i])

    # 3) Monthly overview room nights (w/o COMP/HU/SYS)
    assert lyf.cell(LYF_OVERVIEW_RN_ROW, 2).value == "Room Nights"
    for i, v in enumerate(actual_rooms):
        set_cell(lyf, LYF_OVERVIEW_RN_ROW, HY_COL0 + i, int(v),
                 "Room Nights (overview)", MONTHS[i])

    # 4) Summary sheet: LYF 2025 Occ / ADR / Revpar (feeds LYF!C52:N54)
    for i in range(12):
        set_cell(summary, SUMMARY_OCC_ROW, HY_COL0 + i, actual_occ[i],
                 "Occupancy % (Summary LYF 2025)", MONTHS[i])
        set_cell(summary, SUMMARY_ADR_ROW, HY_COL0 + i, actual_adr[i],
                 "ADR (Summary LYF 2025)", MONTHS[i])

    # 5) Audit sheet
    if AUDIT_SHEET in wb.sheetnames:
        del wb[AUDIT_SHEET]
    au = wb.create_sheet(AUDIT_SHEET)
    bold = Font(bold=True, name="Arial")
    plain = Font(name="Arial")
    au["A1"] = "Reconciliation of LYF (lyf Sukhumvit 8 / LS8) 2025 block vs LS8 Market Segment workbook"
    au["A1"].font = bold
    au["A2"] = f"Run date: {dt.date.today().isoformat()}   Source: {args.ls8}   Target: {args.halfyear}"
    au["A2"].font = plain
    au["A3"] = "Authority: LS8 workbook, sheet 'LS8 Segment 2025'. All cells below were overwritten to match LS8."
    au["A3"].font = plain
    hdr = ["Sheet", "Cell", "Item", "Month", "Old value", "New value (LS8)", "Diff"]
    for j, h in enumerate(hdr, start=1):
        c = au.cell(5, j, h)
        c.font = bold
    r = 6
    for sheet, cellref, item, month, old, new in changes:
        au.cell(r, 1, sheet).font = plain
        au.cell(r, 2, cellref).font = plain
        au.cell(r, 3, item).font = plain
        au.cell(r, 4, month).font = plain
        au.cell(r, 5, old).font = plain
        au.cell(r, 6, new).font = plain
        d = au.cell(r, 7)
        d.value = (f"={get_column_letter(6)}{r}-{get_column_letter(5)}{r}"
                   if isinstance(old, (int, float)) and isinstance(new, (int, float))
                   else "")
        d.font = plain
        r += 1
    r += 1
    notes = [
        "Notes (reviewed, intentionally NOT changed):",
        "1. Summary!C63:N63 'Revenue (THB)' (LYF 2025) is ~3-4% above LS8 room revenue in every month of both "
        "2025 and 2026, so it is on a different basis (not room revenue only) and was left as-is.",
        "2. LYF nationality blocks (rows 33-44 / 77-88) come from the monthly nationality source workbook, not LS8.",
        "3. 2026 H1 block was out of scope for this run. Known diffs vs LS8 'LS8 Segment (2026)': "
        "RN Feb Online 3409 vs 3410, Mar Online 4095 vs 4094, Feb Wholesale 1021 vs 1020; "
        "Revenue Online Jan/Feb/Mar and ASR Jan are below LS8 by ~93,027 THB in total (H1 revenue 38,106,441 vs LS8 38,199,467).",
        "4. Derived cells (totals, mix %, Revpar on LYF, MoM, check row 66) recalculate from the corrected inputs.",
    ]
    for n in notes:
        au.cell(r, 1, n).font = plain if not n.startswith("Notes") else bold
        r += 1
    au.column_dimensions["A"].width = 12
    au.column_dimensions["C"].width = 34
    for col in "EFG":
        au.column_dimensions[col].width = 16

    wb.save(args.out)
    print(f"Wrote {args.out} with {len(changes)} corrected cells.")
    for ch in changes:
        print("  ", ch)


if __name__ == "__main__":
    main()
