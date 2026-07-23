# Property email answers become Excel notes on the reviewed cell

- **Date:** 2026-07-23
- **Decided by:** Mae
- **Status:** active

## Context
New monthly workflow: Mae reviews every property's financial statement, sends
questions about odd items, and the properties answer by email. She wants the
answers kept with the numbers, not lost in the inbox.

## Decision
When a property answers, Claude marks each answer as an **Excel note** (legacy
comment) on that account's row in the month column the question was about —
using the `answer-note` skill and `scripts/mark_answer_note.py`. Notes are
labelled `Ans <PROPERTY> (email <date>)`, are **appended** below any existing
note, and are written on the review copy in `output/` (never `data/source/`).

## Consequences
- The workbook itself carries the full question → answer trail per month.
- Existing notes (Mae's shorthand, preparer notes) are never overwritten.
- Ambiguous matches are asked, not guessed; unanswered questions are reported
  so Mae can chase the property.
- Gmail is not connected yet, so answers are pasted/uploaded; if Mae connects
  Gmail in claude.ai settings, Claude can read the replies directly.
