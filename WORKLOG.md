# Work log

Short notes on what each session did, newest first — so any machine/session
can see the state of the workspace from git alone. One entry per finished
task: date, what was done, files touched. Keep entries to 2–4 lines, plain
language.

---

## 2026-07-18 — Reconciled SR9 / AES / SP tabs (H1) from property sources
New sources in `data/source/`: SR9 Market Segment (2025+YTD2026), AES & SP
2025 Market Mix "as Jun25" (H1 only). New `scripts/reconcile_segments.py`
(surgical patch, same engine as LS8): SR9 2025 H1 + 2026 H1, AES 2025 H1,
SP 2025 H1 — 148 cells fixed on top of the LS8 file → cumulative deliverable
`output/Segment_Half_year_version_1_ALL-reconciled.xlsx` + combined report.
Mae's call: SP January untouched (keeps "Corporate Group with Banque" 672 RN
treatment absent from the source). AES/SP 2026: no source provided yet.
`audit_report.py` generalized to render all "Recon …" audit sheets.

## 2026-07-18 — Simple HTML summary of 2026 diffs vs LS8
Recomputed the 2026 H1 (Jan–Jun) differences fresh from the files: 8 cells
differ — RN Online Feb/Mar, Wholesale Feb, overview Mar (±1 each) and revenue
Online Jan/Feb/Mar + ASR Jan low by ฿93,026.53 total. No workbook changes
(2026 stays out of scope per Mae). File: `output/2026_differences_summary.html`.

## 2026-07-18 — Set up work log + auto-sync rule
Added this WORKLOG.md, a Stop hook (`.claude/hooks/git_sync_check.py`) that
reminds Claude to commit+push after every task, and the matching rule in
CLAUDE.md. Requested by Mae: always push updates and tell other machines
what's happening via git.

## 2026-07-18 — Fixed corrupt reconciled workbook (desktop Excel)
The reconciled file from the openpyxl round-trip lost chart/comment parts and
desktop Excel reported it as damaged. Rewrote `scripts/reconcile_ls8.py` to
patch cell values surgically into a byte-for-byte copy of the original zip
(same 108 corrections + audit sheet as raw XML). Verified: only intended
parts differ from the source; 2025 totals RN 59,216 / THB 82,620,532.57.
Files: `scripts/reconcile_ls8.py`, `output/Segment_Half_year_version_1_LS8-reconciled.xlsx`.

## 2026-07 (earlier) — First LS8 half-year reconciliation + workspace setup
Reconciled the LYF 2025 block of `Segment_Half_year_version_1.xlsx` against
LS8, built the bilingual HTML audit report, and set up the command suite
(`/guide`, `/reconcile-ls8`, `/check-ls8`, `/audit-report`) and session hook.
