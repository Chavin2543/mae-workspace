#!/usr/bin/env python3
"""UserPromptSubmit hook: auto-commit any pending work before a new task starts.

Guarantees the repo rule "always commit before starting a task" even if the
previous turn left changes behind. Stdout is added to Claude's context.
Never blocks the prompt (always exits 0).
"""
import subprocess
import sys


def run(*args):
    return subprocess.run(args, capture_output=True, text=True)


dirty = run("git", "status", "--porcelain").stdout.strip()
if not dirty:
    sys.exit(0)

n = len(dirty.splitlines())
run("git", "add", "-A")
c = run("git", "commit", "-m",
        f"checkpoint: auto-commit {n} pending file(s) before new task [hook]")
if c.returncode == 0:
    sha = run("git", "rev-parse", "--short", "HEAD").stdout.strip()
    print(f"[workspace hook] Auto-committed checkpoint {sha} "
          f"({n} pending file(s)) so this task starts from a committed state. "
          "If that checkpoint holds unfinished work, fold it into your next "
          "proper commit message.")
sys.exit(0)
