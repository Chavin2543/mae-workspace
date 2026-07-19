# LS8 reconciliation: scope and verification method

- **Date:** 2026-07-18
- **Decided by:** Mae (scope), Claude confirmed (method)
- **Status:** active

## Context
First run of "Reconcile half year according to LS8" (July 2026). The half-year
workbook had diffs in both the 2025 and 2026 blocks, and the workbook links to
external workbooks (221 formula cells).

## Decision
1. Reconcile the **2025 block only**; the 2026 H1 diffs are out of scope until
   Mae asks (they are listed in the audit sheet notes).
2. `Summary` "Revenue (THB)" rows for LYF are a different basis (~3–4% above
   LS8 room revenue) — **never reconcile them to LS8**.
3. **Never run LibreOffice recalc on this workbook** — it destroys external
   links (#NAME?) and did not finish even after 50 minutes. Verification is
   the Python re-check (sums vs LS8 actuals), and Excel recalculates on open.

## Consequences
Future runs fix only hardcoded input cells in the agreed block, leave all
formulas and other tabs alone, and verify with Python — details in CLAUDE.md
"Task anatomy".
