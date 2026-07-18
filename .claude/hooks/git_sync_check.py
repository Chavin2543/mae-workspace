#!/usr/bin/env python3
"""Stop hook: enforce Mae's rule — after every task, commit and push, and
record what was done in WORKLOG.md so other machines/sessions can follow.

Blocks the first attempt to stop while the repo has uncommitted changes or
unpushed commits, reminding Claude to sync. Never blocks twice in a row
(stop_hook_active), and never blocks mid-task work-in-progress on purpose —
the reason text tells Claude to continue only if the task is actually done.
"""
import json
import os
import subprocess
import sys


def git(repo, *args):
    try:
        return subprocess.run(
            ["git", "-C", repo, *args],
            capture_output=True, text=True, timeout=10,
        ).stdout.strip()
    except Exception:
        return ""


def main():
    try:
        data = json.load(sys.stdin)
    except Exception:
        return
    if data.get("stop_hook_active"):
        return  # already nudged once this turn — allow stopping
    repo = os.environ.get("CLAUDE_PROJECT_DIR") or os.getcwd()
    if not git(repo, "rev-parse", "--git-dir"):
        return

    problems = []
    dirty = git(repo, "status", "--porcelain")
    if dirty:
        problems.append(f"uncommitted changes:\n{dirty}")
    upstream = git(repo, "rev-parse", "--abbrev-ref", "@{u}")
    if upstream:
        ahead = git(repo, "log", "--oneline", "@{u}..HEAD")
        if ahead:
            problems.append(f"commits not pushed to {upstream}:\n{ahead}")

    if problems:
        print(json.dumps({
            "decision": "block",
            "reason": (
                "Mae's standing rule: after every finished task, add a short "
                "entry to WORKLOG.md (date, what was done, files touched), "
                "commit, and push — so other machines can see what happened "
                "via git. The repo is not synced:\n\n"
                + "\n\n".join(problems)
                + "\n\nIf the task is finished: update WORKLOG.md, commit and "
                "push now. If you are genuinely mid-task (work in progress "
                "across turns), you may stop without pushing."
            ),
        }))


if __name__ == "__main__":
    main()
