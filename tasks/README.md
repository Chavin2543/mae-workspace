# Agent tasks

One markdown file per piece of work, so any session can see what is in flight
and what was delivered.

## Lifecycle

1. **Start** (`/new-task` or by hand): commit any pending work first — the
   UserPromptSubmit hook auto-checkpoints if you forget — then copy
   `TEMPLATE.md` to `tasks/open/YYYY-MM-DD-short-slug.md`, fill in the goal and
   plan, and commit (`task: start <slug>`) before doing the actual work.
2. **Work**: update the checklist as steps finish. Commits during the task use
   the normal message conventions.
3. **Finish** (`/task-done`): fill in the outcome (links to outputs, audit
   sheets, decisions), move the file to `tasks/done/`, commit and push. The
   Stop hook will not let a session end with uncommitted work.

Open tasks are announced at session start (SessionStart hook). Resume or
explicitly close them before starting new work.
