# Git: work on the main branch only

- **Date:** 2026-07-19
- **Decided by:** Workspace owner
- **Status:** active

## Context
Recurring Excel work doesn't need feature branches; extra branches confuse the
history of what was delivered.

## Decision
All work, commits, and pushes go directly to `main`. Never create or switch to
another branch unless Mae explicitly asks. If the session harness forces a
designated branch, follow the harness, but never invent branches.

## Consequences
Enforced by the `guard_bash.py` hook, which blocks branch-creation commands
(`git checkout -b`, `git switch -c`, `git branch <name>`) and force-pushes.
