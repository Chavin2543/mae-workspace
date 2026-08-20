# License & contract tracking sheet — 4 properties (SR9, AES, LYF, SP)

- **Started:** 2026-08-20
- **Requested by:** Mae / workspace owner
- **Status:** open

## Goal
One workbook that tracks every licence, permit, tax registration, insurance and
supplier/OTA contract across the four Mitsui Thailand properties, with expiry
dates, days-left, colour warnings and a dashboard showing what expires in the
next 30/60/90 days.

## Inputs
- Mae's chosen setup (2026-08-20): she will upload her existing licence and
  contract files; alerts = colour coding + summary dashboard tab.
- Counterparties already known from the payment-voucher work (WORKLOG Aug 2026):
  AIMT, Booking.com BV, Expedia/Travelscape, DK Wow Venture, I Klean Laundry,
  Asset World Wex, Laundry Pattaya, PEA, DayUse HK, Adria Scan.
- Uploads to come: filed in `data/source/` (Excel) / `data/pdf/` (PDF).

## Plan / checklist
- [x] Agree structure with Mae (upload-first, dashboard + colours)
- [x] Build `scripts/build_license_tracker.py` (register + dashboard + guide)
- [x] Generate the empty tracker with a seeded checklist of standard Thai hotel
      licences + known contracts per property
- [x] Split every row by renewal responsibility: Owner vs Property
      (Mae's request 2026-08-20) — new `Renew by` column, Owner block
      sorted above Property in each hotel, colour-coded, dashboard split
      into an Owner table and a Property table + a NOT ASSIGNED safety net
- [x] Fixed frozen pane (item name stayed visible when scrolling to dates)
- [x] Asked Mae whether to simplify before uploads; she chose to wait for
      her files — see docs/decisions/2026-08-20-licence-tracker-shape.md
- [x] Dashboard: added a by-document-type view under each of Owner and
      Property, with a check cell tying it to the by-hotel view
- [x] Mae's team handbook received (2510_SA list) — confirmed as the standard
      checklist for every hotel; tracker seed rebuilt from it (278 rows),
      Renew-by derived from the handbook's Kept-by column
- [ ] WAITING ON MAE: reference numbers, issue/expiry dates, fees per hotel
- [ ] Mae reviews the Owner/Property defaults and corrects any
- [ ] Mae uploads her licence/contract documents
- [ ] Load real reference numbers, issue/expiry dates, fees into the register
- [ ] Verify: every uploaded document appears as a row; dates parse; dashboard
      counts equal register counts
- [ ] Deliver the filled workbook + WORKLOG entry, commit & push

## Outcome
(filled in when done)
