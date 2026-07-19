#!/usr/bin/env python3
"""PreToolUse hook (Bash): block shell commands that break workspace rules.

Blocked: creating/switching to new git branches (main-only rule), force-push,
deleting or overwriting uploaded originals in data/source/.
Exit 2 blocks the command; stderr is shown to Claude.
"""
import json
import re
import sys

try:
    data = json.load(sys.stdin)
except Exception:
    sys.exit(0)

cmd = (data.get("tool_input", {}) or {}).get("command", "") or ""

RULES = [
    (r"git\s+checkout\s+(-\w*b\w*|--orphan)\b",
     "creates a new branch — repo rule: work on the current branch only"),
    (r"git\s+switch\s+(-c|-C|--create|--orphan)\b",
     "creates a new branch — repo rule: work on the current branch only"),
    (r"git\s+branch\s+(?![-\s]|$)",
     "creates a new branch — repo rule: work on the current branch only"),
    (r"git\s+push\b[^;&|]*(\s--force\b|\s-f\b|\s--force-with-lease\b)",
     "force-push is not allowed in this workspace"),
    (r"(^|[;&|]\s*)rm\s[^;&|]*data/source",
     "deletes uploaded originals in data/source/ — these are never removed"),
    (r"(^|[;&|]\s*)mv\s[^;&|]*data/source/",
     "moves/overwrites uploaded originals in data/source/ — copy (cp) instead"),
    (r">\s*\"?\'?[^\s;&|]*data/source/",
     "writes into data/source/ — uploaded originals are read-only"),
]

for pattern, why in RULES:
    if re.search(pattern, cmd):
        sys.stderr.write(
            f"BLOCKED by workspace hook: this command {why}. "
            "Adjust the command instead of retrying it verbatim.\n"
        )
        sys.exit(2)
sys.exit(0)
