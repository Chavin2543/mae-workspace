# New workflow: mark property email answers as Excel notes

- **Started:** 2026-07-23
- **Requested by:** Mae
- **Status:** open

## Goal
Every month Mae reviews each property's financial statement and sends questions
about odd items. When a property answers back by email, Claude marks the answer
as an Excel note on that account's row in that month's column. Set this up as a
repeatable skill + helper script.

## Inputs
No workbook this time — this task builds the tooling. Monthly FS workbooks
(e.g. MMR files) are uploaded per month.

## Plan / checklist
- [ ] Skill `.claude/skills/answer-note/SKILL.md` — the workflow
- [ ] Script `scripts/mark_answer_note.py` — add/append an Excel note to a cell
- [ ] Decision log: where answers live and how they're formatted
- [ ] CLAUDE.md: short section on the monthly FS review answer step
- [ ] WORKLOG + commit + push, land on main

## Outcome
(fill when done)
