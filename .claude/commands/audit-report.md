---
description: Rebuild and show the visual audit report from the latest reconciled workbook
---

Rebuild the non-technical audit presentation from the most recent reconciled
workbook and show it to the user. Keep all communication short and plain.

Steps:
1. Find the newest `*_LS8-reconciled.xlsx` in `output/`. If none exists, tell
   the user kindly that no reconciliation has been run yet and point them to
   `/reconcile-ls8`.
2. Run `python3 scripts/audit_report.py <that file>` — it reads the
   "Recon LS8" audit tab and writes an HTML report next to the workbook.
3. Send the HTML to the user with display "render" so it opens in the side
   panel. One-line caption describing the run date and change count.
4. If the user asks for a shareable link, publish the HTML as an Artifact
   (keep the same artifact URL across republs when updating).
