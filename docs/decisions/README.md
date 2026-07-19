# Decision logs

One file per decision, named `YYYY-MM-DD-short-slug.md`. Newest date = current
truth; a later decision can supersede an earlier one (say so explicitly in the
new file).

**When to write one:** any time Mae (or the workspace owner) makes a choice
that future sessions must respect — scope choices ("don't touch the 2026
block"), method choices ("never LibreOffice-recalc this workbook"), workflow
rules ("main branch only"). Use `/log-decision` or copy `TEMPLATE.md`.

**Why:** sessions are ephemeral; anything not written down here is forgotten.
Claude reads the recent decisions at session start (SessionStart hook) and must
check this folder before re-litigating anything that was already decided.
