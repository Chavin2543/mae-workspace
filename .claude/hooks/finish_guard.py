#!/usr/bin/env python3
"""Stop hook: refuse to end the turn while work is uncommitted or unpushed.

Repo rule: every run must be committed and pushed so it is reproducible from
the repo alone. Exit 2 sends Claude back to finish the job; stop_hook_active
prevents infinite loops.
"""
import json
import subprocess
import sys

try:
    data = json.load(sys.stdin)
except Exception:
    data = {}

if data.get("stop_hook_active"):
    sys.exit(0)


def run(*args):
    return subprocess.run(args, capture_output=True, text=True)


dirty = run("git", "status", "--porcelain").stdout.strip()
if dirty:
    n = len(dirty.splitlines())
    sys.stderr.write(
        f"Workspace hook: {n} uncommitted change(s) remain:\n{dirty}\n"
        "Repo rule: commit all work with a descriptive message and push "
        "(git push -u origin <current branch>) before finishing. If some "
        "files are scratch/temp, delete them (they belong in the scratchpad, "
        "not the repo), then finish.\n"
    )
    sys.exit(2)

ahead = run("git", "rev-list", "--count", "@{u}..HEAD")
if ahead.returncode == 0 and ahead.stdout.strip() not in ("", "0"):
    sys.stderr.write(
        f"Workspace hook: {ahead.stdout.strip()} commit(s) not pushed yet. "
        "Run git push -u origin <current branch> before finishing.\n"
    )
    sys.exit(2)
sys.exit(0)
