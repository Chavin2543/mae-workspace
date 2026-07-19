---
description: Start a tracked task (commits pending work first, creates the task file)
argument-hint: short description of the task
---

Start a new tracked task for: $ARGUMENTS (ask Mae for a one-line goal if empty).

Follow `tasks/README.md` exactly:

1. Commit any pending work first with a proper message (the checkpoint hook is
   a safety net, not the standard).
2. Copy `tasks/TEMPLATE.md` to `tasks/open/YYYY-MM-DD-<short-slug>.md` (today's
   date) and fill in Goal, Inputs, and a realistic checklist.
3. Commit it: `task: start <slug>` and push.
4. Then begin the actual work, ticking checklist items as they finish.

If `tasks/open/` already has a task, ask Mae whether to resume it or close it
before opening a new one.
