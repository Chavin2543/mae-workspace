# Recheck the team's handover-documents list (2510_SA)

- **Started:** 2026-08-20
- **Requested by:** Mae
- **Status:** done (2026-08-20)

## Goal
Recheck `2510_SA_List_of_Handover_Docs_number_checked_2.xlsx` (2 tabs:
Handover Docs, All Doc Lists) — numbering, and consistency between the tabs.

## Inputs
- data/source/SA_List_of_Handover_Docs_number_checked_2.xlsx

## Plan / checklist
- [x] Read both tabs in full (excel_map + full row dump)
- [x] Check the running numbers in All Doc Lists
- [x] Cross-check every Handover item against All Doc Lists (dates, sets, memos)
- [x] Check Mitsui-kept items that are absent from the Handover tab
- [x] Write findings into an annotated copy in output/, deliver, commit

## Outcome
17 findings written to `output/SA_List_of_Handover_Docs_rechecked_2026-08-20.xlsx`
(new tab "Recheck 2026-08-20"; the team's two tabs untouched — verified
value-identical). Headlines: numbering broken twice (#6 never used; #44-47 used
twice, so the real count is 204 not 201); the CFA package still dated 2017 on
the Handover tab though All Doc Lists corrected it to 2021 as a known typo; JVA
listed for handover despite an "exclude from handover" memo; 2H-2021 EIA report
listed twice; several date/sets disagreements between the tabs. Source filed at
`data/source/SA_List_of_Handover_Docs_number_checked_2.xlsx`.
