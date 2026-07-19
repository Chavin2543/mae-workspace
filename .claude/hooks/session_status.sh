#!/bin/bash
# SessionStart hook: inject current workspace state into Claude's context.
cd "${CLAUDE_PROJECT_DIR:-.}" 2>/dev/null || exit 0
echo "Workspace state at session start:"
echo "- git branch: $(git branch --show-current 2>/dev/null || echo '?')"
open_tasks=$(ls -1 tasks/open 2>/dev/null | grep -v '^\.gitkeep$')
if [ -n "$open_tasks" ]; then
  echo "- OPEN TASKS (resume or close these first):"
  echo "$open_tasks" | sed 's/^/    - tasks\/open\//'
else
  echo "- no open tasks"
fi
recent=$(ls -1 docs/decisions 2>/dev/null | grep -E '^[0-9]{4}-' | sort | tail -3)
if [ -n "$recent" ]; then
  echo "- recent decisions (docs/decisions/):"
  echo "$recent" | sed 's/^/    - /'
fi
exit 0
