---
description: Record a decision so future sessions follow it without re-asking
argument-hint: the decision to record
---

Record this decision: $ARGUMENTS (if empty, ask Mae what was decided).

1. Check `docs/decisions/` for an existing entry on the same topic — if one
   exists, either update it or create the new file and mark the old one
   `superseded by <new file>`.
2. Copy `docs/decisions/TEMPLATE.md` to
   `docs/decisions/YYYY-MM-DD-<short-slug>.md` (today's date) and fill in
   Context, Decision, and Consequences in plain language.
3. If the decision changes standing rules (scope, workflow, file locations),
   also update the matching section of CLAUDE.md so the two never disagree.
4. Commit (`docs: log decision <slug>`) and push, then confirm to Mae in one
   plain sentence.
