#!/usr/bin/env python3
"""SessionStart hook: keep every machine on the latest central `main`.

Fetches origin/main and fast-forwards the current branch to it when that is
safe (clean worktree, purely behind). Otherwise it tells Claude what to do.
Silent when already up to date; never fails the session (offline is fine).
"""
import json
import os
import subprocess


def git(repo, *args, timeout=20):
    try:
        p = subprocess.run(
            ["git", "-C", repo, *args],
            capture_output=True, text=True, timeout=timeout,
        )
        return p.returncode, p.stdout.strip()
    except Exception:
        return 1, ""


def say(msg):
    print(json.dumps({"systemMessage": msg}))


def main():
    repo = os.environ.get("CLAUDE_PROJECT_DIR") or os.getcwd()
    if git(repo, "rev-parse", "--git-dir")[0] != 0:
        return

    if git(repo, "fetch", "origin", "main", timeout=30)[0] != 0:
        # Offline, or no main branch yet — only warn if main is known to exist
        if git(repo, "rev-parse", "--verify", "origin/main")[0] == 0:
            say("Could not reach GitHub to check for updates — working with "
                "the local copy. Claude: run `git fetch origin main` before "
                "pushing.")
        return
    if git(repo, "rev-parse", "--verify", "origin/main")[0] != 0:
        return  # no central branch yet

    _, behind = git(repo, "rev-list", "--count", "HEAD..origin/main")
    if behind in ("", "0"):
        return  # already contains everything on main

    _, dirty = git(repo, "status", "--porcelain")
    if not dirty and git(repo, "merge", "--ff-only", "origin/main")[0] == 0:
        say(f"Workspace updated with the latest work from main "
            f"({behind} new commit(s)). Read WORKLOG.md to see what changed.")
        return

    say(f"There is newer work on origin/main ({behind} commit(s)) that this "
        "copy does not have. Claude: before starting any task, bring it in "
        "(merge or rebase onto origin/main) so work is never based on stale "
        "files. Read WORKLOG.md after syncing.")


if __name__ == "__main__":
    main()
