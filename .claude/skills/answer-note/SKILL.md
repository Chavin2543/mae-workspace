---
name: answer-note
description: Mark a property's email answer as an Excel note on the matching account row and month column in the monthly financial-statement review workbook. Use when the user says a property "answered", "replied to my question", pastes or uploads an email reply, or asks to "mark the answer on the cell / in that month".
---

## When to use
Mae's monthly FS review loop is: flag big month-over-month gaps → Mae keys
shorthand notes → Claude writes variance questions (skills
`mmr-variance-review` / `note-to-question`) → Mae emails the questions to the
property. **This skill is the last step:** when the property answers, record
each answer as an Excel **note** on that account's row, in the column of the
month the question was about.

Trigger phrases: "SR9 answered", "they replied", "mark this answer",
"บันทึกคำตอบ", or Mae pastes/uploads an email reply.

## Inputs to identify first
1. **The answers** — pasted email text, an uploaded email/screenshot, or (if
   the Gmail connector is connected) the email itself. Note which property
   sent it and which review month it belongs to.
2. **The workbook** — the review copy in `output/` for that property + month
   (e.g. `output/MMR_SL_SR9_Jun26-review.xlsx`). If only the original upload
   exists in `data/source/`, copy it to `output/` first — `data/source/` is
   read-only. If the answers arrive next month inside a newer workbook
   upload, mark the notes in the newer file on the old month's column.
3. **The question list** — the questions column / flagged rows from the
   review, to match each answer to its account row.
4. **Row + column per answer** — account row (match code + name, cols A/B on
   Mapping sheets) and the month column the question was about. Read the
   workbook with `scripts/excel_map.py` first (read-excel skill is mandatory).

## Steps
1. Split the email into per-question answers. Match each to its account row.
   **If a match is ambiguous or an answer has no matching question, ask Mae —
   never guess a row.**
2. For each matched answer, write the note:
   ```bash
   python3 scripts/mark_answer_note.py <workbook> --sheet <sheet> --cell <col+row> \
       --label "Ans <PROPERTY> (email <d Mon yyyy>)" --text "<answer, trimmed>"
   ```
   The script appends below any existing note (Mae's shorthand and preparer
   notes are never lost) and verifies the note reads back.
3. Trim answers to the substance (drop greetings/signatures); keep the
   property's own wording, Thai or English as sent.
4. If the workbook has formulas and is a deliverable, restore formula caches
   per CLAUDE.md (`scripts/restore_formula_caches.py`).
5. Report to Mae: a short table — account → month → one-line answer — plus
   anything unanswered or unmatched.
6. Commit + push (WORKLOG entry), land on `main`.

## Rules
- One property per list; never merge properties.
- Answers referencing an older month go on that older month's cell.
- Never edit `data/source/`; never overwrite existing notes (append only).
- Unanswered questions stay open — list them so Mae can chase.

## Done when
Every answer in the email is a note on the right cell (verified by re-read),
unmatched/unanswered items are reported to Mae, and the workbook is committed.
