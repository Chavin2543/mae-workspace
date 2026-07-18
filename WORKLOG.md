# Work log

Short notes on what each session did, newest first — so any machine/session
can see the state of the workspace from git alone. One entry per finished
task: date, what was done, files touched. Keep entries to 2–4 lines, plain
language.

---

## 2026-07-18 — Centralized git: created `main` as the single branch
All past work (both old `claude/...` branches) now lives on one central
branch, `main`. New rules in CLAUDE.md: sessions pull `main` at start
(new hook `.claude/hooks/git_session_sync.py`) and land finished work on
`main` at end. Added `/sync` command so Mae can save/share with one word.
Files: CLAUDE.md, WORKLOG.md, `.claude/hooks/git_session_sync.py`,
`.claude/commands/sync.md`, `.claude/commands/guide.md`.

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
