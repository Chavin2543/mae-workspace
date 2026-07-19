# Reference library: claude-cookbooks in resources/, not committed

- **Date:** 2026-07-19
- **Decided by:** Workspace owner
- **Status:** active

## Context
The workspace owner wants https://github.com/anthropics/claude-cookbooks
available in the workspace as reference for building automation.

## Decision
Shallow-clone it to `resources/claude-cookbooks/` (356 MB). `resources/` is
gitignored; the SessionStart hook re-clones it in the background whenever it is
missing. It is read-only (enforced by the `protect_paths.py` hook).

## Consequences
Never commit `resources/`; never edit files inside it. On a fresh session,
give the background clone a moment before reading from it.
