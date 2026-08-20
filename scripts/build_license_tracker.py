#!/usr/bin/env python3
"""Build the licence & contract tracking workbook for the four Mitsui Thailand
properties (SR9, AES, LYF, SP).

Structure of the workbook it writes:

  Guide / วิธีใช้   plain EN/TH instructions for Mae
  Dashboard        Owner block + Property block; each block counted twice,
                   by hotel and by document type, then a combined total
                   and the "expiring next" list
  Register         the one master table — every licence/permit/tax/insurance/
                   contract for all four properties, one row each, each
                   tagged Renew by = Owner or Property
  Lists            dropdown source lists (hidden-ish reference tab)

Everything downstream (days left, status, action-by date, the dashboard) is an
Excel formula on top of the Register, so Mae only ever types into the Register
and the rest updates itself when she opens the file.

Usage:
    python3 scripts/build_license_tracker.py \
        --out output/License_Contract_Tracker_4properties.xlsx
    python3 scripts/build_license_tracker.py --out <file> --blank   # no seed rows
"""

import argparse
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# --------------------------------------------------------------------------
# Reference data
# --------------------------------------------------------------------------

PROPERTIES = [
    # code, full name, legal entity
    ("SR9", "Somerset Rama 9 Bangkok", "AMH Ratchada Co., Ltd."),
    ("AES", "Ascott Embassy Sathorn Bangkok", "AMH Sathorn Co., Ltd."),
    ("LYF", "lyf Sukhumvit 8 Bangkok", "AMH Sukhumvit 8 Co., Ltd."),
    ("SP", "Somerset Pattaya", "AMH Pattaya Co., Ltd."),
]

CATEGORIES = [
    "Licence",
    "Permit",
    "Certificate",
    "Tax / Registration",
    "Insurance",
    "Contract - Management",
    "Contract - OTA",
    "Contract - Service",
    "Contract - Utility",
    "Contract - Lease / Finance",
]

RENEW_BY = ["Owner", "Property"]

# Who renews each item. Derived from the handbook's own "Kept by" column
# (2510_SA list): documents kept by the property team (ATB) -> Property;
# documents kept by the owner side (Mitsui / Ananda) -> Owner. Items the
# handbook does not carry keep a conventional default. Mae can flip any row
# with the dropdown.
RESPONSIBILITY = {
    # --- Owner: kept on the owner side in the handbook --------------------
    "Construction permit (Yor.Por.4)": "Owner",
    "Occupation permit (Aor.5)": "Owner",
    "Land & building tax": "Owner",
    "Signage tax": "Owner",
    "Industrial all risks insurance": "Owner",
    "Business interruption insurance": "Owner",
    "Political violence insurance": "Owner",
    "Serviced residence management agreement (Ascott)": "Owner",
    "Management service / central fee agreement": "Owner",
    "Land / building lease": "Owner",
    "F&B outlet operator (MOU + lease + service)": "Owner",
    # --- Property: kept by the property team (ATB) in the handbook --------
    "Hotel operating licence": "Property",
    "Hotel manager licence": "Property",
    "Food establishment licence": "Property",
    "Alcohol sale licence (per outlet)": "Property",
    "Gym licence": "Property",
    "Sauna licence": "Property",
    "Swimming pool licence": "Property",
    "Hazardous substance licence": "Property",
    "Generator licence": "Property",
    "LPG storage licence": "Property",
    "Annual building inspection (Ror.1)": "Property",
    "EIA monitoring report": "Property",
    "Public liability insurance": "Property",
    "Fidelity insurance": "Property",
    "Motor insurance (per vehicle)": "Property",
    "Employee group insurance": "Property",
    "Provident fund registration": "Property",
    "Social security registration": "Property",
    "Workmen compensation registration (KorTor.39)": "Property",
    "Welfare committee appointment": "Property",
    "OHS committee appointment": "Property",
    "Safety officer (Jor Por) appointments": "Property",
    "Workplace environment assessment": "Property",
    "Work permits / visas — foreign staff": "Property",
    "Booking.com accommodation agreement": "Property",
    "Expedia lodging agreement": "Property",
    "Agoda accommodation agreement": "Property",
    "Breakfast service contract": "Property",
    "Limousine service contract": "Property",
    "Car rental contract": "Property",
    "Public area music contract": "Property",
    "Music design contract": "Property",
    "Newspaper / magazine subscription": "Property",
    "Document storage service": "Property",
    "Photocopier rental": "Property",
    "Fire alarm maintenance": "Property",
    "Sound system maintenance": "Property",
    "Electrical equipment maintenance": "Property",
    "Elevator maintenance (per lift)": "Property",
    "Elevator annual audit": "Property",
    "Fitness equipment maintenance": "Property",
    "PABX maintenance": "Property",
    "Air, water & wastewater checking": "Property",
    "Annual energy conservation audit": "Property",
    "Annual building inspection contract": "Property",
    "Hygiene services": "Property",
    "Cleaning service": "Property",
    "Pest control": "Property",
    "Landscape service": "Property",
    "Laundry service": "Property",
    "Lobby aroma scent service": "Property",
    "Leased line (back office)": "Property",
    "Guest internet": "Property",
    "PMS maintenance": "Property",
    "Accounting software maintenance": "Property",
    "Security service": "Property",
    "CCTV system": "Property",
    "Car parking system": "Property",
    "DayUse day-booking agreement": "Property",
    "Adria Scan document services": "Property",
}


STATUSES = ["Expired", "Urgent", "Due soon", "OK", "No date"]

RENEW_CYCLES = ["Annual", "Semi-annual", "Every 2 years", "Every 3 years", "Every 5 years",
                "One-off", "Rolling / evergreen", "Monthly", "Other"]

# Renewable licences / permits / certificates / taxes / insurance — taken from
# the team's handover handbook ("2510_SA List of Handover Docs", All Doc Lists
# tab), which Mae confirmed is the standard checklist for every hotel.
# One-time archival documents (BOD minutes, SPA, CFA package, construction
# contracts, audited FS) are records to keep, not renewals — not seeded here.
# (category, item EN, item TH, authority, typical renewal cycle)
COMMON_LICENCES = [
    ("Licence", "Hotel operating licence", "ใบอนุญาตประกอบธุรกิจโรงแรม",
     "District Office / DOPA", "Every 5 years"),
    ("Licence", "Hotel manager licence", "ใบอนุญาตผู้จัดการโรงแรม",
     "District Office / DOPA", "Every 5 years"),
    ("Licence", "Food establishment licence", "ใบอนุญาตสถานที่จำหน่ายอาหาร",
     "Local authority", "Annual"),
    ("Licence", "Alcohol sale licence (per outlet)", "ใบอนุญาตขายสุรา (ต่อร้าน)",
     "Excise Department", "Annual"),
    ("Licence", "Gym licence", "ใบอนุญาตกิจการฟิตเนส", "Local authority",
     "Annual"),
    ("Licence", "Sauna licence", "ใบอนุญาตกิจการซาวน่า", "Local authority",
     "Annual"),
    ("Licence", "Swimming pool licence", "ใบอนุญาตกิจการสระว่ายน้ำ",
     "Local authority", "Annual"),
    ("Licence", "Hazardous substance licence", "ใบอนุญาตวัตถุอันตราย",
     "Local authority / FDA", "Other"),
    ("Licence", "Generator licence", "ใบอนุญาตเครื่องกำเนิดไฟฟ้า",
     "Dept. of Energy Business", "Annual"),
    ("Licence", "LPG storage licence", "ใบอนุญาตเก็บก๊าซ LPG",
     "Dept. of Energy Business", "Other"),
    ("Permit", "Construction permit (Yor.Por.4)", "ใบอนุญาตก่อสร้าง (ยผ.4)",
     "Local authority", "One-off"),
    ("Permit", "Occupation permit (Aor.5)", "ใบรับรองการใช้อาคาร (อ.5)",
     "Local authority", "One-off"),
    ("Certificate", "Annual building inspection (Ror.1)",
     "รายงานตรวจสอบอาคารประจำปี (ร.1)", "Licensed inspector", "Annual"),
    ("Certificate", "EIA monitoring report", "รายงานติดตามมาตรการ EIA",
     "ONEP / consultant", "Semi-annual"),
    ("Tax / Registration", "Land & building tax", "ภาษีที่ดินและสิ่งปลูกสร้าง",
     "Local authority", "Annual"),
    ("Tax / Registration", "Signage tax", "ภาษีป้าย", "Local authority",
     "Annual"),
    ("Insurance", "Industrial all risks insurance",
     "ประกันภัยความเสี่ยงภัยทรัพย์สิน (IAR)", "Insurer / broker", "Annual"),
    ("Insurance", "Business interruption insurance",
     "ประกันภัยธุรกิจหยุดชะงัก", "Insurer / broker", "Annual"),
    ("Insurance", "Political violence insurance",
     "ประกันภัยความรุนแรงทางการเมือง", "Insurer / broker", "Annual"),
    ("Insurance", "Public liability insurance",
     "ประกันภัยความรับผิดต่อบุคคลภายนอก", "Insurer / broker", "Annual"),
    ("Insurance", "Fidelity insurance", "ประกันภัยความซื่อสัตย์พนักงาน",
     "Insurer / broker", "Annual"),
    ("Insurance", "Motor insurance (per vehicle)", "ประกันภัยรถยนต์ (ต่อคัน)",
     "Insurer / broker", "Annual"),
    ("Insurance", "Employee group insurance", "ประกันภัยกลุ่มพนักงาน",
     "Insurer / broker", "Annual"),
    ("Tax / Registration", "Provident fund registration",
     "ทะเบียนกองทุนสำรองเลี้ยงชีพ", "MOF / fund manager", "One-off"),
    ("Tax / Registration", "Social security registration",
     "ทะเบียนนายจ้างประกันสังคม", "Social Security Office", "One-off"),
    ("Tax / Registration", "Workmen compensation registration (KorTor.39)",
     "ทะเบียนกองทุนเงินทดแทน (กท.39)", "Social Security Office", "One-off"),
    ("Certificate", "Welfare committee appointment",
     "แต่งตั้งคณะกรรมการสวัสดิการ", "Internal / DLPW", "Every 2 years"),
    ("Certificate", "OHS committee appointment",
     "แต่งตั้งคณะกรรมการความปลอดภัยฯ (คปอ.)", "Internal / DLPW",
     "Every 2 years"),
    ("Certificate", "Safety officer (Jor Por) appointments",
     "แต่งตั้งเจ้าหน้าที่ความปลอดภัย (จป.)", "Internal / DLPW", "Other"),
    ("Certificate", "Workplace environment assessment",
     "ตรวจวัดสภาพแวดล้อมในการทำงาน", "Accredited laboratory", "Annual"),
    ("Permit", "Work permits / visas — foreign staff",
     "ใบอนุญาตทำงาน/วีซ่าพนักงานต่างชาติ", "DOE / Immigration", "Annual"),
]

# Contracts — same handbook source. Owner-level agreements first, then the
# operating service contracts the property team runs day to day.
COMMON_CONTRACTS = [
    ("Contract - Management", "Serviced residence management agreement (Ascott)",
     "สัญญาบริหารเซอร์วิสเรสซิเดนซ์ (Ascott)",
     "Ascott International Management (Thailand) Ltd.", "Rolling / evergreen"),
    ("Contract - Management", "Management service / central fee agreement",
     "สัญญาค่าบริการส่วนกลาง",
     "Ascott International Management (Thailand) Ltd.", "Rolling / evergreen"),
    ("Contract - Lease / Finance", "Land / building lease",
     "สัญญาเช่าที่ดิน/อาคาร", "Landlord", "Other"),
    ("Contract - Service", "F&B outlet operator (MOU + lease + service)",
     "สัญญาผู้ประกอบการร้านอาหาร", "", "Other"),
    ("Contract - OTA", "Booking.com accommodation agreement",
     "สัญญากับ Booking.com", "Booking.com B.V.", "Rolling / evergreen"),
    ("Contract - OTA", "Expedia lodging agreement", "สัญญากับ Expedia",
     "Travelscape LLC / Expedia", "Rolling / evergreen"),
    ("Contract - OTA", "Agoda accommodation agreement", "สัญญากับ Agoda",
     "Agoda Company Pte. Ltd.", "Rolling / evergreen"),
    ("Contract - Service", "Breakfast service contract", "สัญญาบริการอาหารเช้า",
     "", "Annual"),
    ("Contract - Service", "Limousine service contract", "สัญญาบริการรถลีมูซีน",
     "", "Annual"),
    ("Contract - Service", "Car rental contract", "สัญญาเช่ารถยนต์", "",
     "Other"),
    ("Contract - Service", "Public area music contract",
     "สัญญาเพลงพื้นที่ส่วนกลาง (ลิขสิทธิ์)", "", "Annual"),
    ("Contract - Service", "Music design contract", "สัญญาออกแบบเสียงเพลง", "",
     "Annual"),
    ("Contract - Service", "Newspaper / magazine subscription",
     "สัญญาหนังสือพิมพ์/นิตยสาร", "", "Annual"),
    ("Contract - Service", "Document storage service", "สัญญารับฝากเอกสาร", "",
     "Annual"),
    ("Contract - Service", "Photocopier rental", "สัญญาเช่าเครื่องถ่ายเอกสาร",
     "", "Other"),
    ("Contract - Service", "Fire alarm maintenance",
     "สัญญาบำรุงรักษาระบบแจ้งเหตุเพลิงไหม้", "", "Annual"),
    ("Contract - Service", "Sound system maintenance",
     "สัญญาบำรุงรักษาระบบเสียง", "", "Annual"),
    ("Contract - Service", "Electrical equipment maintenance",
     "สัญญาบำรุงรักษาอุปกรณ์ไฟฟ้า", "", "Annual"),
    ("Contract - Service", "Elevator maintenance (per lift)",
     "สัญญาบำรุงรักษาลิฟต์ (ต่อตัว)", "", "Annual"),
    ("Contract - Service", "Elevator annual audit", "สัญญาตรวจสอบลิฟต์ประจำปี",
     "", "Annual"),
    ("Contract - Service", "Fitness equipment maintenance",
     "สัญญาบำรุงรักษาอุปกรณ์ฟิตเนส", "", "Annual"),
    ("Contract - Service", "PABX maintenance", "สัญญาบำรุงรักษาระบบโทรศัพท์",
     "", "Annual"),
    ("Contract - Service", "Air, water & wastewater checking",
     "สัญญาตรวจวัดอากาศ น้ำ และน้ำทิ้ง", "", "Annual"),
    ("Contract - Service", "Annual energy conservation audit",
     "สัญญาตรวจสอบการอนุรักษ์พลังงานประจำปี", "", "Annual"),
    ("Contract - Service", "Annual building inspection contract",
     "สัญญาตรวจสอบอาคารประจำปี", "", "Annual"),
    ("Contract - Service", "Hygiene services", "สัญญาบริการสุขอนามัย", "",
     "Annual"),
    ("Contract - Service", "Cleaning service", "สัญญาบริการทำความสะอาด", "",
     "Annual"),
    ("Contract - Service", "Pest control", "สัญญากำจัดแมลง", "", "Annual"),
    ("Contract - Service", "Landscape service", "สัญญาดูแลภูมิทัศน์", "",
     "Annual"),
    ("Contract - Service", "Laundry service", "สัญญาบริการซักรีด", "", "Annual"),
    ("Contract - Service", "Lobby aroma scent service",
     "สัญญาบริการน้ำหอมล็อบบี้", "", "Annual"),
    ("Contract - Service", "Leased line (back office)",
     "สัญญาลีสไลน์สำนักงาน", "", "Annual"),
    ("Contract - Service", "Guest internet", "สัญญาอินเทอร์เน็ตแขก", "",
     "Annual"),
    ("Contract - Service", "PMS maintenance", "สัญญาบำรุงรักษาระบบ PMS", "",
     "Annual"),
    ("Contract - Service", "Accounting software maintenance",
     "สัญญาบำรุงรักษาซอฟต์แวร์บัญชี", "", "Annual"),
    ("Contract - Service", "Security service", "สัญญารักษาความปลอดภัย", "",
     "Annual"),
    ("Contract - Service", "CCTV system", "สัญญาระบบกล้องวงจรปิด", "", "Annual"),
    ("Contract - Service", "Car parking system", "สัญญาระบบที่จอดรถ", "",
     "Annual"),
]

# Counterparties already known per property (payment-voucher work, Aug 2026).
KNOWN_COUNTERPARTIES = {
    "SR9": {"Laundry service": "Asset World Wex Co., Ltd.",
            "F&B outlet operator (MOU + lease + service)": "DK Wow Venture Co., Ltd."},
    "AES": {"Laundry service": "I Klean Laundry Co., Ltd.",
            "F&B outlet operator (MOU + lease + service)": "DK Wow Venture Co., Ltd."},
    "SP": {"Laundry service": "Laundry Pattaya",
           "F&B outlet operator (MOU + lease + service)": "DK Wow Venture Co., Ltd."},
}

# Extra property-specific contract rows.
EXTRA_ROWS = {
    "LYF": [
        ("Contract - OTA", "DayUse day-booking agreement", "สัญญากับ DayUse",
         "DayUse Hong Kong Ltd.", "Rolling / evergreen"),
        ("Contract - Service", "Adria Scan document services",
         "สัญญาบริการ Adria Scan", "Adria Scan", "Annual"),
    ],
}

# --------------------------------------------------------------------------
# Column layout of the Register sheet
# --------------------------------------------------------------------------

COLUMNS = [
    ("No.", 6),
    ("Property", 11),
    ("Legal entity", 24),
    ("Category", 22),
    ("Renew by", 11),
    ("Item (EN)", 40),
    ("รายการ (TH)", 34),
    ("Authority / Counterparty", 34),
    ("Reference / Licence no.", 22),
    ("Issue date", 12),
    ("Expiry date", 12),
    ("Renewal cycle", 16),
    ("Notice (days)", 12),
    ("Action by", 12),
    ("Days left", 11),
    ("Status", 12),
    ("Responsible", 16),
    ("Fee / Value (THB)", 16),
    ("Auto-renew", 12),
    ("Document location", 26),
    ("Notes", 40),
    ("SortKey", 10),   # hidden helper: days-left made unique per row
]
HEADER_ROW = 4
FIRST_DATA_ROW = 5
N_ROWS = 400  # formulas run this far so Mae can add rows without extra work

C = {name: get_column_letter(i + 1) for i, (name, _) in enumerate(COLUMNS)}

# --------------------------------------------------------------------------
# Styling helpers
# --------------------------------------------------------------------------

NAVY = "1F3864"
HEAD_FILL = PatternFill("solid", fgColor=NAVY)
TITLE_FONT = Font(bold=True, size=16, color=NAVY)
SUB_FONT = Font(size=10, color="595959")
HEAD_FONT = Font(bold=True, color="FFFFFF", size=10)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FILL_EXPIRED = PatternFill("solid", fgColor="F8CBAD")   # red-ish
FILL_URGENT = PatternFill("solid", fgColor="FFC7CE")    # pink
FILL_SOON = PatternFill("solid", fgColor="FFE699")      # amber
FILL_OK = PatternFill("solid", fgColor="C6EFCE")        # green
FILL_NODATE = PatternFill("solid", fgColor="EDEDED")    # grey

FILL_OWNER = PatternFill("solid", fgColor="DDEBF7")     # light blue
FILL_PROPERTY = PatternFill("solid", fgColor="E4DFEC")  # light purple
FONT_OWNER = Font(color="1F3864", bold=True)
FONT_PROPERTY = Font(color="5B3A8E", bold=True)

FONT_EXPIRED = Font(color="833C00", bold=True)
FONT_URGENT = Font(color="9C0006", bold=True)
FONT_SOON = Font(color="7F6000")
FONT_OK = Font(color="375623")
FONT_NODATE = Font(color="808080")


def title_block(ws, title, subtitle, width_cols):
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A2"] = subtitle
    ws["A2"].font = SUB_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=width_cols)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=width_cols)
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[3].height = 6


# --------------------------------------------------------------------------
# Sheet builders
# --------------------------------------------------------------------------

def build_lists(wb):
    """Reference tab holding the dropdown source lists."""
    ws = wb.create_sheet("Lists")
    ws["A1"] = "Dropdown lists — used by the Register. Edit here to change the choices."
    ws["A1"].font = Font(bold=True, color=NAVY)

    blocks = [
        ("A", "Property", [p[0] for p in PROPERTIES]),
        ("C", "Category", CATEGORIES),
        ("E", "Renewal cycle", RENEW_CYCLES),
        ("G", "Auto-renew", ["Yes", "No"]),
        ("I", "Status (reference)", STATUSES),
        ("O", "Renew by", RENEW_BY),
    ]
    for col, head, values in blocks:
        ws[f"{col}3"] = head
        ws[f"{col}3"].font = Font(bold=True)
        ws[f"{col}3"].fill = HEAD_FILL
        ws[f"{col}3"].font = HEAD_FONT
        for i, v in enumerate(values):
            ws[f"{col}{4 + i}"] = v
        ws.column_dimensions[col].width = 22

    # Property reference (code -> full name -> legal entity)
    ws["K3"] = "Property"
    ws["L3"] = "Full name"
    ws["M3"] = "Legal entity"
    for cell in ("K3", "L3", "M3"):
        ws[cell].fill = HEAD_FILL
        ws[cell].font = HEAD_FONT
    for i, (code, full, entity) in enumerate(PROPERTIES):
        ws.cell(row=4 + i, column=11, value=code)
        ws.cell(row=4 + i, column=12, value=full)
        ws.cell(row=4 + i, column=13, value=entity)
    ws.column_dimensions["K"].width = 12
    ws.column_dimensions["L"].width = 32
    ws.column_dimensions["M"].width = 26
    return ws


def seed_rows(blank=False):
    """Build the seeded checklist rows.

    One block per property, and inside each property the Owner-renewed items
    come first, then the Property-renewed ones — so the two responsibilities
    read as two separate blocks even before Mae touches the filter.
    """
    if blank:
        return []

    missing = set()
    rows = []
    for code, _full, entity in PROPERTIES:
        items = list(COMMON_LICENCES) + list(COMMON_CONTRACTS) + \
            list(EXTRA_ROWS.get(code, []))
        known = KNOWN_COUNTERPARTIES.get(code, {})
        block = []
        for category, item_en, item_th, authority, cycle in items:
            who = RESPONSIBILITY.get(item_en)
            if who is None:
                missing.add(item_en)
            block.append({
                "Property": code,
                "Legal entity": entity,
                "Category": category,
                "Renew by": who or "",
                "Item (EN)": item_en,
                "รายการ (TH)": item_th,
                "Authority / Counterparty": known.get(item_en, authority),
                "Renewal cycle": cycle,
            })
        # Owner block first, then Property; inside each, keep category order
        order = {w: i for i, w in enumerate(RENEW_BY)}
        block.sort(key=lambda x: (order.get(x["Renew by"], 99),
                                  CATEGORIES.index(x["Category"])))
        rows.extend(block)

    if missing:
        raise SystemExit(
            "RESPONSIBILITY is missing an Owner/Property answer for:\n  - "
            + "\n  - ".join(sorted(missing)))
    return rows


def build_register(wb, rows):
    ws = wb.create_sheet("Register")
    ncols = len(COLUMNS)
    title_block(
        ws,
        "Licence & Contract Register — Mitsui Thailand portfolio",
        "Type into the white columns only. Days left / Action by / Status "
        "calculate themselves. ใส่ข้อมูลเฉพาะช่องสีขาว ช่องสูตรจะคำนวณเอง",
        ncols,
    )

    for i, (name, width) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=HEADER_ROW, column=i, value=name)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = BOX
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[HEADER_ROW].height = 30

    last_row = FIRST_DATA_ROW + N_ROWS - 1

    for offset in range(N_ROWS):
        r = FIRST_DATA_ROW + offset
        data = rows[offset] if offset < len(rows) else None

        # running number, only shown when the row has an item
        ws[f"{C['No.']}{r}"] = (
            f'=IF({C["Item (EN)"]}{r}="","",'
            f'COUNTA(${C["Item (EN)"]}${FIRST_DATA_ROW}:{C["Item (EN)"]}{r}))'
        )

        if data:
            for key, value in data.items():
                if value:
                    ws[f"{C[key]}{r}"] = value

        # notice period default: 60 days before expiry
        if data:
            ws[f"{C['Notice (days)']}{r}"] = 60

        # Action by = expiry - notice days
        ws[f"{C['Action by']}{r}"] = (
            f'=IF({C["Expiry date"]}{r}="","",'
            f'{C["Expiry date"]}{r}-IF({C["Notice (days)"]}{r}="",0,'
            f'{C["Notice (days)"]}{r}))'
        )
        # Days left = expiry - today
        ws[f"{C['Days left']}{r}"] = (
            f'=IF({C["Expiry date"]}{r}="","",{C["Expiry date"]}{r}-TODAY())'
        )
        # Status
        ws[f"{C['Status']}{r}"] = (
            f'=IF({C["Item (EN)"]}{r}="","",'
            f'IF({C["Expiry date"]}{r}="","No date",'
            f'IF({C["Days left"]}{r}<0,"Expired",'
            f'IF({C["Days left"]}{r}<=30,"Urgent",'
            f'IF({C["Days left"]}{r}<=90,"Due soon","OK")))))'
        )
        # hidden sort key: days left nudged by the row number so two items
        # expiring on the same day still get two distinct ranks on the Dashboard
        ws[f"{C['SortKey']}{r}"] = (
            f'=IF({C["Item (EN)"]}{r}="","",'
            f'IF({C["Expiry date"]}{r}="","",'
            f'{C["Days left"]}{r}+ROW()/100000))'
        )

        for i in range(1, ncols + 1):
            c = ws.cell(row=r, column=i)
            c.border = BOX
            c.alignment = Alignment(vertical="top", wrap_text=(i in (5, 6, 7, 20)))
        for key in ("Issue date", "Expiry date", "Action by"):
            ws[f"{C[key]}{r}"].number_format = "DD-MMM-YYYY"
            ws[f"{C[key]}{r}"].alignment = Alignment(horizontal="center")
        ws[f"{C['Days left']}{r}"].number_format = "#,##0"
        ws[f"{C['Days left']}{r}"].alignment = Alignment(horizontal="center")
        ws[f"{C['Fee / Value (THB)']}{r}"].number_format = "#,##0.00"
        ws[f"{C['No.']}{r}"].alignment = Alignment(horizontal="center")
        ws[f"{C['Status']}{r}"].alignment = Alignment(horizontal="center")

    # --- dropdowns -------------------------------------------------------
    validations = [
        ("Property", f"Lists!$A$4:$A${3 + len(PROPERTIES)}"),
        ("Category", f"Lists!$C$4:$C${3 + len(CATEGORIES)}"),
        ("Renew by", f"Lists!$O$4:$O${3 + len(RENEW_BY)}"),
        ("Renewal cycle", f"Lists!$E$4:$E${3 + len(RENEW_CYCLES)}"),
        ("Auto-renew", "Lists!$G$4:$G$5"),
    ]
    for key, formula in validations:
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        dv.error = "Pick a value from the list / เลือกจากรายการ"
        ws.add_data_validation(dv)
        dv.add(f"{C[key]}{FIRST_DATA_ROW}:{C[key]}{last_row}")

    # --- colour the whole row by status ---------------------------------
    visible_last = C["Notes"]   # helper SortKey column sits after this

    # Renew by gets its own colour, added FIRST so it outranks the status
    # colour on that one cell — Owner and Property stay readable in every row.
    who_range = f"{C['Renew by']}{FIRST_DATA_ROW}:{C['Renew by']}{last_row}"
    who_abs = f"${C['Renew by']}{FIRST_DATA_ROW}"
    for label, fill, font in (("Owner", FILL_OWNER, FONT_OWNER),
                              ("Property", FILL_PROPERTY, FONT_PROPERTY)):
        ws.conditional_formatting.add(
            who_range,
            FormulaRule(formula=[f'{who_abs}="{label}"'], fill=fill, font=font,
                        stopIfTrue=True),
        )

    body = f"A{FIRST_DATA_ROW}:{visible_last}{last_row}"
    status_abs = f"${C['Status']}{FIRST_DATA_ROW}"
    rules = [
        ("Expired", FILL_EXPIRED, FONT_EXPIRED),
        ("Urgent", FILL_URGENT, FONT_URGENT),
        ("Due soon", FILL_SOON, FONT_SOON),
        ("OK", FILL_OK, FONT_OK),
        ("No date", FILL_NODATE, FONT_NODATE),
    ]
    for label, fill, font in rules:
        ws.conditional_formatting.add(
            body,
            FormulaRule(formula=[f'{status_abs}="{label}"'], fill=fill,
                        font=font, stopIfTrue=False),
        )

    # Freeze through Item (EN) so a row still says what it is once you scroll
    # right to the dates. Derived from C, not hardcoded, so inserting another
    # column never silently strips the item name out of the frozen pane again.
    first_scrolling_col = get_column_letter(
        COLUMNS.index(("Item (EN)", 40)) + 2)
    ws.freeze_panes = f"{first_scrolling_col}{FIRST_DATA_ROW}"
    ws.auto_filter.ref = f"A{HEADER_ROW}:{visible_last}{last_row}"
    ws.column_dimensions[C["SortKey"]].hidden = True
    ws.sheet_view.zoomScale = 90
    return ws, last_row


def build_dashboard(wb, last_row):
    """Dashboard split two ways: who renews it, and what type of document.

    For each responsibility (Owner, Property) there are two tables reading the
    same rows — one **by hotel**, one **by document type** — so a red number
    can be traced to both "which hotel" and "what kind of document". The two
    tables of a pair must show the same Total; a check cell says so out loud.
    """
    ws = wb.create_sheet("Dashboard", 0)
    title_block(
        ws,
        "Licence & Contract Dashboard — 4 properties",
        "Split by who renews it (OWNER / PROPERTY), then by hotel and by "
        "document type. All of it reads the Register tab. "
        "แยกตามผู้รับผิดชอบ แล้วแยกตามโรงแรมและประเภทเอกสาร",
        9,
    )

    reg = lambda key: (f"Register!${C[key]}${FIRST_DATA_ROW}:"
                       f"${C[key]}${last_row}")
    reg_status, reg_item, reg_who = reg("Status"), reg("Item (EN)"), reg("Renew by")

    ws["A4"] = "As at"
    ws["A4"].font = Font(bold=True)
    ws["B4"] = "=TODAY()"
    ws["B4"].number_format = "DD-MMM-YYYY"

    STATUS_ORDER = ["Expired", "Urgent", "Due soon", "OK", "No date"]
    STATUS_HEAD = ["Expired", "Urgent (≤30 d)", "Due soon (≤90 d)", "OK",
                   "No date"]
    COUNT_FILLS = [(2, FILL_EXPIRED, FONT_EXPIRED), (3, FILL_URGENT, FONT_URGENT),
                   (4, FILL_SOON, FONT_SOON), (5, FILL_OK, FONT_OK)]

    HOTEL_ROWS = [(code, code, full) for code, full, _e in PROPERTIES]
    TYPE_ROWS = [(cat, cat, "") for cat in CATEGORIES]

    def matrix(top, corner, dim_key, rows_def, who, accent):
        """One count table. `who` filters Renew by; None means 'not filled in'."""
        for i, h in enumerate([corner] + STATUS_HEAD + ["Total"], start=1):
            c = ws.cell(row=top, column=i, value=h)
            c.fill = PatternFill("solid", fgColor=accent)
            c.font = HEAD_FONT
            c.alignment = Alignment(horizontal="center", vertical="center",
                                    wrap_text=True)
            c.border = BOX
        ws.row_dimensions[top].height = 30

        who_crit = f'{reg_who},"{who}"' if who else f'{reg_who},""'
        first = top + 1
        for i, (value, label, note) in enumerate(rows_def):
            r = first + i
            ws.cell(row=r, column=1, value=label)
            if note:
                ws.cell(row=r, column=9, value=note).font = Font(color="595959")
            dim_crit = f'{reg(dim_key)},"{value}"'
            for j, status in enumerate(STATUS_ORDER):
                ws.cell(row=r, column=2 + j,
                        value=(f'=COUNTIFS({dim_crit},{who_crit},'
                               f'{reg_status},"{status}")'))
            ws.cell(row=r, column=7,
                    value=f'=COUNTIFS({dim_crit},{who_crit},{reg_item},"?*")')

        sub = first + len(rows_def)
        ws.cell(row=sub, column=1, value="Subtotal").font = Font(bold=True)
        for col in range(2, 8):
            L = get_column_letter(col)
            c = ws.cell(row=sub, column=col,
                        value=f"=SUM({L}{first}:{L}{sub - 1})")
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="F2F2F2")

        for r in range(first, sub + 1):
            for col in range(1, 8):
                c = ws.cell(row=r, column=col)
                c.border = BOX
                if col > 1:
                    c.alignment = Alignment(horizontal="center")

        for col, fill, font in COUNT_FILLS:
            L = get_column_letter(col)
            ws.conditional_formatting.add(
                f"{L}{first}:{L}{sub}",
                FormulaRule(formula=[f"{L}{first}>0"], fill=fill, font=font))
        return sub

    def section(top, who, banner, accent, with_types=True):
        """Banner + the by-hotel table and (optionally) the by-type table."""
        ws.cell(row=top, column=1, value=banner).font = Font(
            bold=True, size=12, color=accent)
        hotel_sub = matrix(top + 1, "Hotel", "Property", HOTEL_ROWS, who, accent)
        if not with_types:
            return hotel_sub, hotel_sub + 2
        type_sub = matrix(hotel_sub + 2, "Document type", "Category",
                          TYPE_ROWS, who, accent)
        # the two tables count the same rows, so their totals must agree
        chk = ws.cell(row=type_sub, column=9,
                      value=(f'=IF(G{type_sub}=G{hotel_sub},'
                             f'"✓ same total as by hotel",'
                             f'"CHECK — does not match by hotel")'))
        chk.font = Font(color="595959", italic=True)
        return hotel_sub, type_sub + 2

    owner_sub, nxt = section(
        6, "Owner",
        "OWNER — renewed by the asset company (AMH …) / เจ้าของอาคาร", "1F3864")
    prop_sub, nxt = section(
        nxt, "Property",
        "PROPERTY — renewed by the hotel operation / ฝ่ายโรงแรม", "5B3A8E")
    none_sub, nxt = section(
        nxt, None,
        "NOT ASSIGNED — no Renew by chosen yet / ยังไม่ได้ระบุผู้รับผิดชอบ",
        "808080", with_types=False)

    # combined total, proving the blocks are the whole register
    gt = nxt
    ws.cell(row=gt, column=1, value="ALL").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=gt, column=1).fill = HEAD_FILL
    ws.cell(row=gt, column=1).border = BOX
    for col in range(2, 8):
        L = get_column_letter(col)
        c = ws.cell(row=gt, column=col,
                    value=f"={L}{owner_sub}+{L}{prop_sub}+{L}{none_sub}")
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center")
        c.border = BOX
    ws.cell(row=gt, column=9,
            value="Owner + Property + Not assigned = every row in the Register"
            ).font = SUB_FONT

    # --- soonest-expiring list, both responsibilities together -----------
    lst_head = gt + 3
    ws.cell(row=lst_head - 1, column=1,
            value="Expiring next — soonest first, both responsibilities"
            ).font = Font(bold=True, size=12, color=NAVY)

    list_cols = [
        ("Hotel", C["Property"], 1),
        ("Renew by", C["Renew by"], 2),
        ("Document type", C["Category"], 3),
        ("Item", C["Item (EN)"], 4),
        ("Counterparty", C["Authority / Counterparty"], 5),
        ("Expiry date", C["Expiry date"], 6),
        ("Days left", C["Days left"], 7),
        ("Status", C["Status"], 8),
        ("Responsible", C["Responsible"], 9),
    ]
    for head, _letter, col in list_cols:
        c = ws.cell(row=lst_head, column=col, value=head)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = Alignment(horizontal="center", wrap_text=True)
        c.border = BOX

    reg_sort = reg("SortKey")
    n_list = 30
    for k in range(1, n_list + 1):
        r = lst_head + k
        # rank comes from the row position, so no Rank column is needed
        key = f"SMALL({reg_sort},ROW()-{lst_head})"
        for _head, letter, col in list_cols:
            src = f"Register!${letter}${FIRST_DATA_ROW}:${letter}${last_row}"
            ws.cell(row=r, column=col,
                    value=f'=IFERROR(INDEX({src},MATCH({key},{reg_sort},0)),"")')
        ws.cell(row=r, column=6).number_format = "DD-MMM-YYYY"
        for _head, _letter, col in list_cols:
            c = ws.cell(row=r, column=col)
            c.border = BOX
            c.alignment = Alignment(horizontal="center" if col in (1, 2, 6, 7, 8)
                                    else "left", wrap_text=col in (3, 4, 5),
                                    vertical="center")

    # Renew by colour first so it wins on its own cell
    for label, fill, font in (("Owner", FILL_OWNER, FONT_OWNER),
                              ("Property", FILL_PROPERTY, FONT_PROPERTY)):
        ws.conditional_formatting.add(
            f"B{lst_head + 1}:B{lst_head + n_list}",
            FormulaRule(formula=[f'$B{lst_head + 1}="{label}"'], fill=fill,
                        font=font, stopIfTrue=True))
    for label, fill, font in (("Expired", FILL_EXPIRED, FONT_EXPIRED),
                              ("Urgent", FILL_URGENT, FONT_URGENT),
                              ("Due soon", FILL_SOON, FONT_SOON),
                              ("OK", FILL_OK, FONT_OK)):
        ws.conditional_formatting.add(
            f"A{lst_head + 1}:I{lst_head + n_list}",
            FormulaRule(formula=[f'$H{lst_head + 1}="{label}"'], fill=fill,
                        font=font))

    # column A carries the row labels of the matrices (longest text on the
    # sheet), so it sets the width; the rest serve both the tables and the list
    widths = [27, 13, 17, 26, 21, 14, 12, 12, 17]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # --- legend ----------------------------------------------------------
    r = lst_head + n_list + 3
    ws.cell(row=r, column=1, value="Who renews it / ใครต่ออายุ").font = Font(
        bold=True, color=NAVY)
    for label, meaning, fill, font in (
        ("Owner", "the asset company (AMH …): building, entity, taxes, asset "
                  "insurance, head agreements — เจ้าของอาคาร",
         FILL_OWNER, FONT_OWNER),
        ("Property", "the hotel operation: guest-facing licences, staff, "
                     "suppliers, OTAs — ฝ่ายโรงแรม",
         FILL_PROPERTY, FONT_PROPERTY),
    ):
        r += 1
        c = ws.cell(row=r, column=1, value=label)
        c.fill, c.font, c.border = fill, font, BOX
        c.alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=meaning).font = SUB_FONT
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)

    r += 2
    ws.cell(row=r, column=1, value="Colours / สีบอกสถานะ").font = Font(
        bold=True, color=NAVY)
    for label, meaning, fill, font in (
        ("Expired", "past the expiry date — เลยกำหนดแล้ว", FILL_EXPIRED,
         FONT_EXPIRED),
        ("Urgent", "expires within 30 days — ครบกำหนดใน 30 วัน", FILL_URGENT,
         FONT_URGENT),
        ("Due soon", "expires within 90 days — ครบกำหนดใน 90 วัน", FILL_SOON,
         FONT_SOON),
        ("OK", "more than 90 days left — ยังมีเวลา", FILL_OK, FONT_OK),
        ("No date", "expiry date not filled in yet — ยังไม่ได้ใส่วันหมดอายุ",
         FILL_NODATE, FONT_NODATE),
    ):
        r += 1
        c = ws.cell(row=r, column=1, value=label)
        c.fill, c.font, c.border = fill, font, BOX
        c.alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=meaning).font = SUB_FONT
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)

    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A5"
    return ws


def build_guide(wb):
    ws = wb.create_sheet("Guide", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 95

    ws["B1"] = "How to use this file / วิธีใช้ไฟล์นี้"
    ws["B1"].font = TITLE_FONT
    ws.row_dimensions[1].height = 26

    lines = [
        ("Three tabs", "Dashboard = the overview. Register = where you type. "
                       "Lists = the dropdown choices."),
        ("แท็บ 3 อัน", "Dashboard คือภาพรวม, Register คือที่กรอกข้อมูล, "
                       "Lists คือรายการตัวเลือก"),
        ("", ""),
        ("Step 1", "Go to the Register tab. Each row is one licence, permit, "
                   "tax, insurance policy or contract."),
        ("Step 2", "Fill in: Reference / Licence no., Issue date, Expiry date, "
                   "Responsible person, Fee, and where the document is kept."),
        ("", ""),
        ("Renew by", "The important split: OWNER or PROPERTY — who has to renew "
                     "it. Pick from the dropdown. Owner rows are blue, Property "
                     "rows are purple, and the Dashboard shows the two as two "
                     "separate tables."),
        ("Renew by = Owner", "The asset company (AMH … Co., Ltd.): the building "
                             "and its inspections, the company registrations and "
                             "taxes, asset insurance, the management agreement, "
                             "lease and bank facilities."),
        ("Renew by = Property", "The hotel operation: guest-facing licences "
                                "(food, alcohol, pool), staff permits and health "
                                "certificates, OTA agreements, and operating "
                                "suppliers such as laundry, security, cleaning."),
        ("ผู้ต่ออายุ", "Owner = เจ้าของอาคาร (บริษัท AMH) ดูแลตัวอาคาร ทะเบียนบริษัท "
                      "ภาษี ประกันทรัพย์สิน สัญญาหลัก / Property = ฝ่ายโรงแรม "
                      "ดูแลใบอนุญาตหน้างาน พนักงาน OTA และซัพพลายเออร์"),
        ("See one side only", "On the Register click the filter arrow on "
                              "'Renew by' and tick just Owner or just Property."),
        ("Step 3", "Leave the grey formula columns alone — Action by, Days left "
                   "and Status fill themselves in."),
        ("Step 4", "Open the Dashboard tab to see what is expiring."),
        ("", ""),
        ("The Dashboard", "Owner first, then Property. Each one is counted "
                          "twice — once BY HOTEL and once BY DOCUMENT TYPE "
                          "(licence, permit, certificate, tax, insurance, "
                          "contract) — so a red number tells you which "
                          "hotel and what kind of document."),
        ("แดชบอร์ด", "นับ 2 แบบ: แยกตามโรงแรม และแยกตามประเภทเอกสาร"),
        ("The ✓ mark", "Next to each by-type table: it confirms that table counts the same rows as the by-hotel table above it."),
        ("", ""),
        ("Dates", "Type dates as a real date (e.g. 31/12/2026). If Excel shows "
                  "a number instead, format the cell as Date."),
        ("Notice (days)", "How many days before expiry you want to start the "
                          "renewal. Default 60. 'Action by' = expiry minus this."),
        ("Rows", "Formulas already run to row 404. Just keep typing down — do "
                 "not insert rows above 404 and you never need to copy formulas."),
        ("Delete a row", "Clear the Item (EN) cell and the row drops out of the "
                         "count and the dashboard."),
        ("", ""),
        ("Colours", "Red = expired. Pink = expires within 30 days. Amber = "
                    "within 90 days. Green = fine. Grey = no expiry date yet."),
        ("สี", "แดง = หมดอายุแล้ว, ชมพู = ครบกำหนดใน 30 วัน, เหลือง = ใน 90 วัน, "
               "เขียว = ยังมีเวลา, เทา = ยังไม่ได้ใส่วันหมดอายุ"),
        ("", ""),
        ("Check this first", "The Owner / Property split I filled in is my "
                             "best guess from normal practice. Please read down "
                             "the Renew by column once and change anything that "
                             "is different in your agreements."),
        ("", ""),
        ("The seeded rows", "The checklist comes from your team's handbook "
                            "(the 2510_SA document list) — the renewable items, "
                            "seeded for all four hotels. No dates yet. Delete "
                            "anything a hotel does not have, add anything "
                            "missing. Archive-only papers (BOD minutes, SPA, "
                            "CFA, construction contracts) are not here — they "
                            "do not expire."),
        ("รายการที่ใส่ไว้ให้", "มาจากคู่มือของทีม (ลิสต์ 2510_SA) เฉพาะรายการที่ต้อง"
                              "ต่ออายุ ใส่ให้ครบทั้ง 4 โรงแรม ยังไม่มีวันที่ "
                              "ลบอันที่ไม่มี เพิ่มอันที่ขาด"),
        ("", ""),
        ("Rebuild", "scripts/build_license_tracker.py rebuilds this file's "
                    "structure. Rebuilding starts from empty rows, so keep "
                    "your filled-in copy — do not overwrite it."),
    ]
    r = 3
    for label, text in lines:
        if not label and not text:
            r += 1
            continue
        c = ws.cell(row=r, column=2, value=label)
        c.font = Font(bold=True, color=NAVY)
        c.alignment = Alignment(vertical="top")
        t = ws.cell(row=r, column=3, value=text)
        t.alignment = Alignment(vertical="top", wrap_text=True)
        ws.row_dimensions[r].height = 30 if len(text) > 95 else 16
        r += 1
    return ws


# --------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--out", required=True, help="output .xlsx path")
    ap.add_argument("--blank", action="store_true",
                    help="no seeded checklist rows, just the empty structure")
    args = ap.parse_args()

    wb = Workbook()
    wb.remove(wb.active)

    rows = seed_rows(blank=args.blank)
    build_lists(wb)
    _, last_row = build_register(wb, rows)
    build_dashboard(wb, last_row)
    build_guide(wb)

    # tab order: Guide, Dashboard, Register, Lists
    wb._sheets = [wb["Guide"], wb["Dashboard"], wb["Register"], wb["Lists"]]
    wb.active = 1

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    print(f"Wrote {out}")
    print(f"  seeded rows: {len(rows)}  (formulas run to row {last_row})")
    by_prop = {}
    for row in rows:
        by_prop[row["Property"]] = by_prop.get(row["Property"], 0) + 1
    for code, n in by_prop.items():
        print(f"    {code}: {n} rows")


if __name__ == "__main__":
    main()
