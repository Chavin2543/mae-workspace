---
description: Close the current open task (fill outcome, move to done, commit & push)
---

Close the task in `tasks/open/` (if several, ask Mae which one).

1. Verify the task's checklist is genuinely complete — run any verification
   steps that were skipped. If something can't be completed, record it under
   Outcome as left open instead of silently dropping it.
2. Fill in **Outcome**: deliverable paths in `output/`, key numbers, links to
   any `docs/decisions/` files created, anything left open.
3. Set Status to `done (YYYY-MM-DD)` and `git mv` the file to `tasks/done/`.
4. Commit (`task: done <slug>`) and push.
5. Tell Mae in one or two plain sentences what was delivered.
