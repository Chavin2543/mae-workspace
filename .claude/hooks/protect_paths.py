#!/usr/bin/env python3
"""PreToolUse hook (Write|Edit|NotebookEdit): block edits to read-only areas.

data/source/  = uploaded originals, never edited in place (repo rule).
resources/    = reference library (claude-cookbooks etc.), read-only.
Exit 2 blocks the tool call; stderr is shown to Claude.
"""
import json
import os
import sys

try:
    data = json.load(sys.stdin)
except Exception:
    sys.exit(0)

tool_input = data.get("tool_input", {}) or {}
fp = tool_input.get("file_path") or tool_input.get("notebook_path") or ""
if not fp:
    sys.exit(0)

root = os.environ.get("CLAUDE_PROJECT_DIR") or os.getcwd()
try:
    rel = os.path.relpath(os.path.realpath(fp), os.path.realpath(root))
except ValueError:
    sys.exit(0)
rel = rel.replace(os.sep, "/")

PROTECTED = {
    "data/source/": "uploaded original workbooks are never edited in place — "
                    "copy the file to output/ (or scratchpad) and edit the copy",
    "data/pdf/": "uploaded original PDFs are never edited in place — "
                 "copy the file elsewhere and work on the copy",
    "resources/": "the reference library is read-only — quote or copy what you "
                  "need instead of editing it",
}
for prefix, why in PROTECTED.items():
    if rel.startswith(prefix):
        sys.stderr.write(
            f"BLOCKED by workspace hook: '{rel}' is protected ({why}).\n"
        )
        sys.exit(2)
sys.exit(0)
