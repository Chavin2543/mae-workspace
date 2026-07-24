# Create "payment-approval-table" skill

- **Started:** 2026-07-24
- **Requested by:** Mae
- **Status:** done (2026-07-24)

## Goal
A reusable skill: Mae uploads the month's payment vouchers (4 PVs, different
companies), Claude reads them and returns copy-pasteable tables (one per
company) for her approval email — No. / Detail / Payment Voucher, with
Description, Pay to, Amount in the Detail column.

## Inputs
- Sample PV shared in chat 2026-07-24: AMH Ratchada Co., Ltd. (Ananda
  Development voucher), VAT P.P.30 June 2026, pay to Revenue Department,
  1,297,598.79 THB.

## Plan / checklist
- [ ] Write `.claude/skills/payment-approval-table/SKILL.md`
- [ ] Document the recurring task in CLAUDE.md
- [ ] Show Mae the sample output for the voucher she shared, confirm format
- [ ] WORKLOG entry, commit, push, land on main

## Outcome
Skill created at `.claude/skills/payment-approval-table/SKILL.md`; recurring
task documented in CLAUDE.md. Sample table for the AMH Ratchada voucher shown
to Mae in chat for format confirmation — refine the skill if she wants
changes when the first real 4-PV batch arrives.
