#!/usr/bin/env python3
"""Build the licence & contract tracking workbook for the four Mitsui Thailand
properties (SR9, AES, LYF, SP).

Structure of the workbook it writes:

  Guide / วิธีใช้   plain EN/TH instructions for Mae
  Dashboard        counts per property x status + "expiring next 90 days" list
  Register         the one master table — every licence/permit/tax/insurance/
                   contract for all four properties, one row each
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

STATUSES = ["Expired", "Urgent", "Due soon", "OK", "No date"]

RENEW_CYCLES = ["Annual", "Every 2 years", "Every 3 years", "Every 5 years",
                "One-off", "Rolling / evergreen", "Monthly", "Other"]

# Standard licences/permits every Thai hotel-and-serviced-residence carries.
# Seeded for all four properties as a checklist — Mae deletes what does not
# apply and fills in the real numbers and dates from her documents.
# (category, item EN, item TH, authority, typical renewal cycle)
COMMON_LICENCES = [
    ("Licence", "Hotel operating licence", "ใบอนุญาตประกอบธุรกิจโรงแรม",
     "Dept. of Provincial Administration / District Office", "Every 5 years"),
    ("Licence", "Hotel manager licence", "ใบอนุญาตผู้จัดการโรงแรม",
     "Dept. of Provincial Administration / District Office", "Every 5 years"),
    ("Permit", "Building construction permit (อ.1)", "ใบอนุญาตก่อสร้างอาคาร (อ.1)",
     "Local authority (BMA / City Municipality)", "One-off"),
    ("Certificate", "Certificate of building use (อ.6)", "ใบรับรองการใช้อาคาร (อ.6)",
     "Local authority (BMA / City Municipality)", "One-off"),
    ("Certificate", "Annual building inspection report (ร.1)",
     "รายงานการตรวจสอบอาคารประจำปี (ร.1)",
     "Dept. of Public Works / local authority", "Annual"),
    ("Certificate", "Fire protection system inspection",
     "การตรวจสอบระบบป้องกันอัคคีภัย", "Licensed inspector", "Annual"),
    ("Certificate", "Lift / elevator inspection certificate",
     "ใบรับรองการตรวจสอบลิฟต์", "Licensed engineer", "Annual"),
    ("Certificate", "Electrical system inspection",
     "การตรวจสอบระบบไฟฟ้า", "Licensed engineer", "Annual"),
    ("Certificate", "Generator / fuel storage permit",
     "ใบอนุญาตเครื่องกำเนิดไฟฟ้า / เก็บน้ำมันเชื้อเพลิง",
     "Dept. of Energy Business / local authority", "Annual"),
    ("Certificate", "Boiler / pressure vessel inspection",
     "การตรวจสอบหม้อไอน้ำ / ภาชนะรับแรงดัน", "Licensed engineer", "Annual"),
    ("Licence", "Food establishment licence",
     "ใบอนุญาตสถานที่จำหน่ายอาหาร/สะสมอาหาร",
     "Local authority (BMA / City Municipality)", "Annual"),
    ("Licence", "Alcohol sale licence", "ใบอนุญาตขายสุรา",
     "Excise Department", "Annual"),
    ("Licence", "Tobacco sale licence", "ใบอนุญาตขายยาสูบ",
     "Excise Department", "Annual"),
    ("Licence", "Swimming pool / public place licence",
     "ใบอนุญาตสถานที่สาธารณะ (สระว่ายน้ำ)",
     "Local authority (BMA / City Municipality)", "Annual"),
    ("Certificate", "Swimming pool water quality test",
     "ผลตรวจคุณภาพน้ำสระว่ายน้ำ", "Accredited laboratory", "Other"),
    ("Certificate", "Drinking / potable water quality test",
     "ผลตรวจคุณภาพน้ำดื่ม", "Accredited laboratory", "Annual"),
    ("Certificate", "Wastewater treatment monitoring report",
     "รายงานผลการตรวจวัดคุณภาพน้ำทิ้ง",
     "Dept. of Pollution Control / local authority", "Other"),
    ("Certificate", "EIA environmental monitoring report",
     "รายงานผลการปฏิบัติตามมาตรการ EIA",
     "Office of Natural Resources & Environmental Policy (ONEP)", "Other"),
    ("Permit", "Waste / refuse disposal permit", "ใบอนุญาตกำจัดสิ่งปฏิกูลและมูลฝอย",
     "Local authority", "Annual"),
    ("Licence", "Radio / walkie-talkie frequency licence",
     "ใบอนุญาตวิทยุคมนาคม", "NBTC", "Annual"),
    ("Licence", "Music / copyright public performance licence",
     "ใบอนุญาตเผยแพร่งานดนตรี (MCT / ลิขสิทธิ์)",
     "MCT / Phonorights / collecting societies", "Annual"),
    ("Tax / Registration", "Signage tax (billboard)", "ภาษีป้าย",
     "Local authority", "Annual"),
    ("Tax / Registration", "Land & building tax", "ภาษีที่ดินและสิ่งปลูกสร้าง",
     "Local authority", "Annual"),
    ("Tax / Registration", "VAT registration (ภ.พ.20)",
     "ใบทะเบียนภาษีมูลค่าเพิ่ม (ภ.พ.20)", "Revenue Department", "One-off"),
    ("Tax / Registration", "Company registration / affidavit",
     "หนังสือรับรองบริษัท", "Dept. of Business Development", "Other"),
    ("Tax / Registration", "Social security registration",
     "ทะเบียนนายจ้างประกันสังคม", "Social Security Office", "One-off"),
    ("Permit", "Work permits / visas — foreign staff",
     "ใบอนุญาตทำงาน / วีซ่าพนักงานต่างชาติ",
     "Dept. of Employment / Immigration", "Annual"),
    ("Certificate", "Food handler health certificates",
     "ใบรับรองสุขภาพผู้สัมผัสอาหาร", "Public health office", "Annual"),
    ("Certificate", "Pest control service certificate",
     "ใบรับรองการกำจัดแมลงและสัตว์พาหะ", "Licensed contractor", "Annual"),
    ("Insurance", "Property all-risk insurance",
     "ประกันภัยทรัพย์สินแบบ All Risk", "Insurer / broker", "Annual"),
    ("Insurance", "Public liability insurance",
     "ประกันภัยความรับผิดต่อบุคคลภายนอก", "Insurer / broker", "Annual"),
    ("Insurance", "Business interruption insurance",
     "ประกันภัยธุรกิจหยุดชะงัก", "Insurer / broker", "Annual"),
    ("Insurance", "Hotel guest / hotel-act compulsory insurance",
     "ประกันภัยตาม พ.ร.บ. โรงแรม", "Insurer / broker", "Annual"),
]

# Contracts common to all four properties.
COMMON_CONTRACTS = [
    ("Contract - Management", "Hotel management agreement",
     "สัญญาบริหารโรงแรม", "Ascott International Management (Thailand) Ltd.",
     "Rolling / evergreen"),
    ("Contract - Management", "Technical services / central fee agreement",
     "สัญญาค่าบริการส่วนกลาง", "Ascott International Management (Thailand) Ltd.",
     "Rolling / evergreen"),
    ("Contract - OTA", "Booking.com accommodation agreement",
     "สัญญากับ Booking.com", "Booking.com B.V.", "Rolling / evergreen"),
    ("Contract - OTA", "Expedia lodging agreement", "สัญญากับ Expedia",
     "Travelscape LLC / Expedia (Thailand) Co., Ltd.", "Rolling / evergreen"),
    ("Contract - OTA", "Agoda accommodation agreement", "สัญญากับ Agoda",
     "Agoda Company Pte. Ltd.", "Rolling / evergreen"),
    ("Contract - Service", "Laundry & linen services", "สัญญาบริการซักรีด",
     "", "Annual"),
    ("Contract - Service", "F&B outlet / catering operator",
     "สัญญาผู้ประกอบการอาหารและเครื่องดื่ม", "DK Wow Venture Co., Ltd.", "Annual"),
    ("Contract - Service", "Security services", "สัญญาบริการรักษาความปลอดภัย",
     "", "Annual"),
    ("Contract - Service", "Cleaning / housekeeping outsource",
     "สัญญาบริการทำความสะอาด", "", "Annual"),
    ("Contract - Service", "Pest control", "สัญญากำจัดแมลง", "", "Annual"),
    ("Contract - Service", "Landscaping / plant care", "สัญญาดูแลต้นไม้และภูมิทัศน์",
     "", "Annual"),
    ("Contract - Service", "Lift maintenance", "สัญญาบำรุงรักษาลิฟต์", "", "Annual"),
    ("Contract - Service", "HVAC / chiller maintenance",
     "สัญญาบำรุงรักษาระบบปรับอากาศ", "", "Annual"),
    ("Contract - Service", "Fire system maintenance",
     "สัญญาบำรุงรักษาระบบดับเพลิง", "", "Annual"),
    ("Contract - Service", "Waste collection", "สัญญาเก็บขนขยะ", "", "Annual"),
    ("Contract - Service", "PMS / property management system licence",
     "สัญญาระบบบริหารโรงแรม (PMS)", "", "Annual"),
    ("Contract - Service", "Internet / IPTV / telephone",
     "สัญญาอินเทอร์เน็ต / IPTV / โทรศัพท์", "", "Annual"),
    ("Contract - Service", "External audit & tax agent",
     "สัญญาผู้สอบบัญชีและที่ปรึกษาภาษี", "", "Annual"),
    ("Contract - Utility", "Electricity supply", "สัญญาไฟฟ้า", "", "Rolling / evergreen"),
    ("Contract - Utility", "Water supply", "สัญญาน้ำประปา", "", "Rolling / evergreen"),
    ("Contract - Lease / Finance", "Land / building lease",
     "สัญญาเช่าที่ดิน / อาคาร", "", "Other"),
    ("Contract - Lease / Finance", "Bank facility / loan agreement",
     "สัญญาสินเชื่อธนาคาร", "", "Other"),
]

# Counterparties already known per property from the payment-voucher work
# (WORKLOG Aug 2026) — pre-filled so Mae recognises the rows.
KNOWN_COUNTERPARTIES = {
    "SR9": {"Laundry & linen services": "Asset World Wex Co., Ltd."},
    "AES": {"Laundry & linen services": "I Klean Laundry Co., Ltd."},
    "SP": {"Laundry & linen services": "Laundry Pattaya",
           "Electricity supply": "Provincial Electricity Authority (PEA)"},
    "LYF": {"F&B outlet / catering operator": ""},
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
    """Build the seeded checklist rows: one block per property."""
    if blank:
        return []
    rows = []
    for code, _full, entity in PROPERTIES:
        items = list(COMMON_LICENCES) + list(COMMON_CONTRACTS) + \
            list(EXTRA_ROWS.get(code, []))
        known = KNOWN_COUNTERPARTIES.get(code, {})
        for category, item_en, item_th, authority, cycle in items:
            rows.append({
                "Property": code,
                "Legal entity": entity,
                "Category": category,
                "Item (EN)": item_en,
                "รายการ (TH)": item_th,
                "Authority / Counterparty": known.get(item_en, authority),
                "Renewal cycle": cycle,
            })
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

    ws.freeze_panes = f"E{FIRST_DATA_ROW}"
    ws.auto_filter.ref = f"A{HEADER_ROW}:{visible_last}{last_row}"
    ws.column_dimensions[C["SortKey"]].hidden = True
    ws.sheet_view.zoomScale = 90
    return ws, last_row


def build_dashboard(wb, last_row):
    ws = wb.create_sheet("Dashboard", 0)
    title_block(
        ws,
        "Licence & Contract Dashboard — 4 properties",
        "Everything below reads the Register tab. Reopen the file to refresh. "
        "ตัวเลขทั้งหมดดึงจากแท็บ Register อัตโนมัติ",
        9,
    )
    ws["A2"].font = SUB_FONT

    reg_status = f"Register!${C['Status']}${FIRST_DATA_ROW}:${C['Status']}${last_row}"
    reg_prop = f"Register!${C['Property']}${FIRST_DATA_ROW}:${C['Property']}${last_row}"
    reg_item = f"Register!${C['Item (EN)']}${FIRST_DATA_ROW}:${C['Item (EN)']}${last_row}"

    ws["A4"] = "As at"
    ws["A4"].font = Font(bold=True)
    ws["B4"] = "=TODAY()"
    ws["B4"].number_format = "DD-MMM-YYYY"

    # --- status matrix ---------------------------------------------------
    head_row = 6
    headers = ["Property", "Expired", "Urgent (≤30 d)", "Due soon (≤90 d)",
               "OK", "No date", "Total items"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=head_row, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        c.border = BOX
    ws.row_dimensions[head_row].height = 30

    status_cols = {"Expired": 2, "Urgent (≤30 d)": 3, "Due soon (≤90 d)": 4,
                   "OK": 5, "No date": 6}
    status_lookup = {"Expired": "Expired", "Urgent (≤30 d)": "Urgent",
                     "Due soon (≤90 d)": "Due soon", "OK": "OK",
                     "No date": "No date"}

    first = head_row + 1
    for i, (code, full, _entity) in enumerate(PROPERTIES):
        r = first + i
        # column A holds the bare code so COUNTIFS can match the Register
        ws.cell(row=r, column=1, value=code).font = Font(bold=True)
        ws.cell(row=r, column=9, value=full).font = Font(color="595959")
        for label, col in status_cols.items():
            ws.cell(
                row=r, column=col,
                value=(f'=COUNTIFS({reg_prop},$A{r},{reg_status},'
                       f'"{status_lookup[label]}")'),
            )
        ws.cell(row=r, column=7,
                value=f'=COUNTIFS({reg_prop},$A{r},{reg_item},"?*")')

    total_row = first + len(PROPERTIES)
    ws.cell(row=total_row, column=1, value="TOTAL").font = Font(bold=True)
    for col in range(2, 8):
        L = get_column_letter(col)
        ws.cell(row=total_row, column=col,
                value=f"=SUM({L}{first}:{L}{total_row - 1})").font = Font(bold=True)

    for r in range(first, total_row + 1):
        for col in range(1, 8):
            c = ws.cell(row=r, column=col)
            c.border = BOX
            if col > 1:
                c.alignment = Alignment(horizontal="center")

    # colour the matrix count cells
    ws.conditional_formatting.add(f"B{first}:B{total_row}",
                                  FormulaRule(formula=[f"B{first}>0"],
                                              fill=FILL_EXPIRED, font=FONT_EXPIRED))
    ws.conditional_formatting.add(f"C{first}:C{total_row}",
                                  FormulaRule(formula=[f"C{first}>0"],
                                              fill=FILL_URGENT, font=FONT_URGENT))
    ws.conditional_formatting.add(f"D{first}:D{total_row}",
                                  FormulaRule(formula=[f"D{first}>0"],
                                              fill=FILL_SOON, font=FONT_SOON))
    ws.conditional_formatting.add(f"E{first}:E{total_row}",
                                  FormulaRule(formula=[f"E{first}>0"],
                                              fill=FILL_OK, font=FONT_OK))

    # --- next 90 days list ----------------------------------------------
    lst_head = total_row + 3
    ws.cell(row=lst_head - 1, column=1,
            value="Expiring next — soonest first (already expired items included)"
            ).font = Font(bold=True, size=12, color=NAVY)
    ws.cell(row=lst_head - 1, column=1).alignment = Alignment(vertical="center")

    list_headers = ["Rank", "Property", "Category", "Item", "Counterparty",
                    "Expiry date", "Days left", "Status", "Responsible"]
    for i, h in enumerate(list_headers, start=1):
        c = ws.cell(row=lst_head, column=i, value=h)
        c.fill = HEAD_FILL
        c.font = HEAD_FONT
        c.alignment = Alignment(horizontal="center")
        c.border = BOX

    reg_sort = (f"Register!${C['SortKey']}${FIRST_DATA_ROW}:"
                f"${C['SortKey']}${last_row}")
    n_list = 30
    pull_cols = [
        (2, C["Property"]),
        (3, C["Category"]),
        (4, C["Item (EN)"]),
        (5, C["Authority / Counterparty"]),
        (6, C["Expiry date"]),
        (7, C["Days left"]),
        (8, C["Status"]),
        (9, C["Responsible"]),
    ]
    for k in range(1, n_list + 1):
        r = lst_head + k
        ws.cell(row=r, column=1, value=k).alignment = Alignment(horizontal="center")
        key = f"SMALL({reg_sort},$A{r})"
        for col, letter in pull_cols:
            src = f"Register!${letter}${FIRST_DATA_ROW}:${letter}${last_row}"
            ws.cell(
                row=r, column=col,
                value=(f'=IFERROR(INDEX({src},MATCH({key},{reg_sort},0)),"")'),
            )
        ws.cell(row=r, column=6).number_format = "DD-MMM-YYYY"
        for col in range(1, 10):
            c = ws.cell(row=r, column=col)
            c.border = BOX
            if col in (1, 6, 7, 8):
                c.alignment = Alignment(horizontal="center")

    list_body = f"A{lst_head + 1}:I{lst_head + n_list}"
    for label, fill, font in [("Expired", FILL_EXPIRED, FONT_EXPIRED),
                              ("Urgent", FILL_URGENT, FONT_URGENT),
                              ("Due soon", FILL_SOON, FONT_SOON),
                              ("OK", FILL_OK, FONT_OK)]:
        ws.conditional_formatting.add(
            list_body,
            FormulaRule(formula=[f'$H{lst_head + 1}="{label}"'], fill=fill,
                        font=font),
        )

    widths = [10, 11, 22, 42, 32, 14, 11, 12, 16]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # --- legend ----------------------------------------------------------
    leg = lst_head + n_list + 3
    ws.cell(row=leg, column=1, value="Colours / สีบอกสถานะ").font = Font(
        bold=True, color=NAVY)
    legend = [
        ("Expired", "past the expiry date — เลยกำหนดแล้ว", FILL_EXPIRED, FONT_EXPIRED),
        ("Urgent", "expires within 30 days — ครบกำหนดใน 30 วัน", FILL_URGENT, FONT_URGENT),
        ("Due soon", "expires within 90 days — ครบกำหนดใน 90 วัน", FILL_SOON, FONT_SOON),
        ("OK", "more than 90 days left — ยังมีเวลา", FILL_OK, FONT_OK),
        ("No date", "expiry date not filled in yet — ยังไม่ได้ใส่วันหมดอายุ",
         FILL_NODATE, FONT_NODATE),
    ]
    for i, (label, meaning, fill, font) in enumerate(legend, start=1):
        c = ws.cell(row=leg + i, column=1, value=label)
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal="center")
        c.border = BOX
        ws.cell(row=leg + i, column=2, value=meaning).font = SUB_FONT
        ws.merge_cells(start_row=leg + i, start_column=2, end_row=leg + i,
                       end_column=5)

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
        ("Step 3", "Leave the grey formula columns alone — Action by, Days left "
                   "and Status fill themselves in."),
        ("Step 4", "Open the Dashboard tab to see what is expiring."),
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
        ("The seeded rows", "The rows already in the Register are a checklist of "
                            "the licences and contracts a Thai hotel normally "
                            "holds — they have no dates yet. Delete anything a "
                            "property does not have, and add anything missing."),
        ("รายการที่ใส่ไว้ให้", "เป็นเช็คลิสต์ใบอนุญาต/สัญญาที่โรงแรมไทยมักมี "
                              "ยังไม่มีวันที่ ลบอันที่ไม่มี และเพิ่มอันที่ขาด"),
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
