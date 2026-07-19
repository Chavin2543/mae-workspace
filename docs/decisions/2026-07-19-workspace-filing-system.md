# Workspace filing system, task workflow, and enforcement hooks

- **Date:** 2026-07-19
- **Decided by:** Workspace owner
- **Status:** active

## Context
The workspace needed a standard place for every file type, a durable record of
decisions and tasks, and guarantees (not just written rules) that sessions
follow the workflow.

## Decision
1. **Filing:** every file type has one home — see the "Where every file goes"
   table in CLAUDE.md. Uploaded originals in `data/` are read-only.
2. **Decisions:** logged in `docs/decisions/`, one dated file each.
3. **Tasks:** one file per task in `tasks/open/` → moved to `tasks/done/` when
   finished; **always commit before starting a task** (auto-enforced).
4. **Enforcement:** rules are enforced by hooks in `.claude/hooks/` (blocked
   actions, auto-checkpoint commits, no finishing with uncommitted/unpushed
   work) — Claude cannot skip them.

## Consequences
New file types get a home added to the CLAUDE.md table (and this log) rather
than being dropped in the repo root. Hook changes are themselves decisions —
log them.
