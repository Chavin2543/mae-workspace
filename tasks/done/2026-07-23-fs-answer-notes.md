# New workflow: mark property email answers as Excel notes

- **Started:** 2026-07-23
- **Requested by:** Mae
- **Status:** done (2026-07-23)

## Goal
Every month Mae reviews each property's financial statement and sends questions
about odd items. When a property answers back by email, Claude marks the answer
as an Excel note on that account's row in that month's column. Set this up as a
repeatable skill + helper script.

## Inputs
No workbook this time — this task builds the tooling. Monthly FS workbooks
(e.g. MMR files) are uploaded per month.

## Plan / checklist
- [x] Skill `.claude/skills/answer-note/SKILL.md` — the workflow
- [x] Script `scripts/mark_answer_note.py` — add/append an Excel note to a cell
- [x] Decision log: where answers live and how they're formatted
- [x] CLAUDE.md: short section on the monthly FS review answer step
- [x] WORKLOG + commit + push, land on main

## Outcome
Delivered the answer-marking workflow:
- `.claude/skills/answer-note/SKILL.md` — email reply → Excel note on the
  account row + month column (append-only, ask on ambiguous matches).
- `scripts/mark_answer_note.py` — tested: appends below existing notes,
  verifies read-back, refuses data/source/, warns about formula caches.
- Decision: `docs/decisions/2026-07-23-fs-answer-notes.md`.
- CLAUDE.md: new "Monthly FS review (questions → answers)" section.
Left open: Gmail connector not connected — Mae pastes replies for now.
