"""Generate a sample of the automatic Excel export, in the department's existing layout.

Dates run down the rows and doctors across the columns, with GUH red, RGH blue, NHH green,
leave yellow and weekends dark. All names are fictional demo data.

Usage: python3 tools/sample_excel_export.py docs/samples/haematology-roster-export-sample.xlsx
"""
import sys
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FILL = {
    "GUH": "F28B82", "RGH": "6FB1EA", "NHH": "93C47D",
    "LEAVE": "FFE36E", "ONCALL": "D9D2E9", "OFF": "FFFFFF",
    "WEEKEND": "262626", "HEAD": "D9E1EA", "PENDING": "FFF2B3",
}
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTRE = Alignment(horizontal="center", vertical="center", wrap_text=True)
ONCALL = {"NIGHTS", "LONG DAY", "TWILIGHT", "WARD COVER", "OFF PRE ON-CALL", "EARLY TWILIGHT", "W/E SD"}
LEAVE_LABELS = {"AL": "ANNUAL LEAVE", "SL": "STUDY LEAVE", "SICK": "SICK LEAVE", "AL?": "ANNUAL LEAVE (PENDING)"}

START = date(2026, 11, 16)
RESIDENTS = [("F1", "E. Ward"), ("F2", "P. Nair"), ("F2 (on calls NHH)", "T. Hughes"),
             ("IMT2", "S. Malik"), ("IMT3", "B. Carter"), ("JCF", "L. Osei")]
REGISTRARS = ["Omar", "Hannah", "Daniel", "Zara", "Meera", "Yusuf"]
RES = [
    "GUH|GUH|OFF|GUH|GUH|-|-|GUH|AL|AL|AL|AL|-|-",
    "RGH|NHH|NHH|NHH|NIGHTS|NIGHTS|NIGHTS|OFF|OFF|NHH|NHH|NHH|-|-",
    "TWILIGHT|TWILIGHT|SL|TWILIGHT|OFF|-|-|GUH|GUH|GUH|GUH|OFF PRE ON-CALL|LONG DAY|LONG DAY",
    "GUH|RGH|RGH|RGH|RGH|-|-|RGH|RGH|WARD COVER|RGH|RGH|-|-",
    "GUH|GUH LAB|GUH|RGH DAY UNIT|GUH|-|-|GUH|NHH|RGH|RGH AM; PM ANNUAL LEAVE|GUH|-|-",
    "RGH|RGH|RGH|LONG DAY|LONG DAY|LONG DAY|LONG DAY|OFF|OFF|GUH|GUH|RGH|-|-",
]
REG = [
    "GUH|GUH|GUH|GUH|GUH|-|-|RGH|RGH|RGH|RGH|RGH|-|-",
    "RGH|RGH|AL|AL|AL|-|-|GUH|GUH|GUH|SpR TRAINING DAY|GUH|-|-",
    "NHH|NHH|NHH|NHH|NHH|-|-|NHH - DAY UNIT|NHH|NHH|NHH|OFF|-|-",
    "RGH - DAY UNIT|RGH|RGH|RGH|RGH|-|-|AL|AL|AL|AL|AL|-|-",
    "GUH LAB|GUH LAB|GUH|GUH|OFF|-|-|GUH LAB|GUH|OFF|AL?|GUH|-|-",
    "RGH MDCU|RGH|RGH|RGH|RGH|-|-|RGH|RGH MDCU|RGH|OFF|SICK|-|-",
]
CONSULTANT_OF_WEEK = {0: "Dr Fraser", 1: "Dr Lin"}
RESTRICTED = {0: "AVOID", 1: "AVOID", 7: "NO AL"}


def site(code):
    for s in ("GUH", "RGH", "NHH"):
        if code.startswith(s):
            return s
    return None


def style_cell(cell, code, weekend):
    cell.border = BORDER
    cell.alignment = CENTRE
    cell.font = Font(name="Calibri", size=10)
    if code == "-":
        cell.value = None
        if weekend:
            cell.fill = PatternFill("solid", fgColor=FILL["WEEKEND"])
        return
    cell.value = LEAVE_LABELS.get(code, code)
    s = site(code)
    if code == "AL?":
        cell.fill = PatternFill("lightUp", fgColor="C9A800", bgColor=FILL["PENDING"])
    elif s and "LEAVE" not in code:
        cell.fill = PatternFill("solid", fgColor=FILL[s])
    elif code in ONCALL:
        cell.fill = PatternFill("solid", fgColor=FILL["ONCALL"])
    elif code == "OFF":
        cell.fill = PatternFill("solid", fgColor=FILL["OFF"])
    else:
        cell.fill = PatternFill("solid", fgColor=FILL["LEAVE"])


def cover_text(codes, dow, residents):
    """Residents need GUH and RGH daily and NHH Tue-Fri; registrars must have GUH."""
    parts = []
    for s in ("GUH", "RGH", "NHH"):
        if residents and s == "NHH" and dow == 0:
            parts.append("NHH n/a")
        elif any(site(c) == s for c in codes):
            parts.append(f"{s} OK")
        elif "AL?" in codes and (s == "GUH" or residents):
            parts.append(f"{s} AT RISK (pending leave)")
        elif not residents and s != "GUH":
            parts.append(f"{s} none (ideal only)")
        else:
            parts.append(f"{s} GAP")
    return " · ".join(parts)


def build_roster(ws):
    headers = ["Date", "Day", "Consultant", "Leave restriction"]
    headers += [f"{g}\n{n}" for g, n in RESIDENTS] + [f"Registrar\n{n}" for n in REGISTRARS]
    headers += ["Resident cover", "Registrar cover"]
    ws.append(headers)
    for col, _ in enumerate(headers, start=1):
        c = ws.cell(row=1, column=col)
        c.font = Font(bold=True, size=10)
        c.alignment = CENTRE
        c.border = BORDER
        c.fill = PatternFill("solid", fgColor=FILL["HEAD"])
    ws.row_dimensions[1].height = 42

    res = [r.split("|") for r in RES]
    reg = [r.split("|") for r in REG]
    for i in range(14):
        d = START + timedelta(days=i)
        dow = i % 7
        weekend = dow >= 5
        row = i + 2
        ws.cell(row=row, column=1, value=d).number_format = "DD/MM/YYYY"
        ws.cell(row=row, column=2, value=d.strftime("%A"))
        ws.cell(row=row, column=3, value=None if weekend else CONSULTANT_OF_WEEK[i // 7])
        ws.cell(row=row, column=4, value=RESTRICTED.get(i))
        for col in range(1, 5):
            c = ws.cell(row=row, column=col)
            c.border = BORDER
            c.alignment = CENTRE
            c.font = Font(size=10, color="FFFFFF" if weekend else "000000")
            if weekend:
                c.fill = PatternFill("solid", fgColor=FILL["WEEKEND"])
        if RESTRICTED.get(i):
            colour = "C00000" if RESTRICTED[i] == "NO AL" else "7030A0"
            ws.cell(row=row, column=1).fill = PatternFill("solid", fgColor=colour)
            ws.cell(row=row, column=1).font = Font(size=10, color="FFFFFF", bold=True)
        for j, col_codes in enumerate(res + reg):
            style_cell(ws.cell(row=row, column=5 + j), col_codes[i], weekend)
        res_cov = ws.cell(row=row, column=17)
        reg_cov = ws.cell(row=row, column=18)
        for c in (res_cov, reg_cov):
            c.border = BORDER
            c.alignment = CENTRE
            c.font = Font(size=9)
            if weekend:
                c.fill = PatternFill("solid", fgColor=FILL["WEEKEND"])
        if not weekend:
            res_cov.value = cover_text([r[i] for r in res], dow, residents=True)
            reg_cov.value = cover_text([r[i] for r in reg], dow, residents=False)
            if "GAP" in (res_cov.value or ""):
                res_cov.font = Font(size=9, bold=True, color="C00000")
            if "GUH GAP" in reg_cov.value or "AT RISK" in reg_cov.value:
                reg_cov.font = Font(size=9, bold=True, color="C00000")
        ws.row_dimensions[row].height = 30

    widths = [12, 11, 14, 11] + [14] * 12 + [26, 26]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    ws.freeze_panes = "E2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{15}"


def build_leave_log(ws):
    ws.append(["Request ID", "Submitted", "Doctor", "Grade", "Leave type", "Start", "End", "Part of day",
               "Working days", "AL remaining after", "Proposed cover", "Status", "Decided by", "Decided at", "Reason / notes"])
    rows = [
        ["LR-0141", "10/10/2026 09:12", "Meera K.", "Registrar", "Annual leave", "26/11/2026", "26/11/2026", "Full day",
         1, "8 of 27", "Omar: RGH → GUH", "Pending", "", "", "Sister's wedding"],
        ["LR-0140", "08/10/2026 16:40", "L. Osei", "JCF", "Study leave", "03/12/2026", "03/12/2026", "Full day",
         1, "n/a (SL 17 of 30)", "Not needed", "Pending", "", "", "Course"],
        ["LR-0139", "07/10/2026 11:05", "T. Hughes", "F2", "Day off in lieu", "23/11/2026", "23/11/2026", "Full day",
         1, "n/a", "B. Carter stays at GUH", "Pending", "", "", "In lieu of SL 15/10. Date is marked NO AL"],
        ["LR-0131", "28/09/2026 10:02", "B. Carter", "IMT3", "Annual leave", "26/11/2026", "26/11/2026", "PM",
         0.5, "11.5 of 27", "Not needed", "Approved", "Dr Fraser (Rota lead)", "29/09/2026 08:15", ""],
        ["LR-0127", "21/09/2026 13:30", "E. Ward", "F1", "Annual leave", "24/11/2026", "27/11/2026", "Full day",
         4, "6 of 27", "T. Hughes to GUH", "Approved", "Dr Fraser (Rota lead)", "22/09/2026 17:48", ""],
    ]
    for r in rows:
        ws.append(r)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor=FILL["HEAD"])
    for idx, w in enumerate([10, 17, 12, 10, 16, 11, 11, 10, 8, 16, 24, 10, 22, 17, 36], start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    ws.freeze_panes = "A2"


def main(out):
    wb = Workbook()
    ws = wb.active
    ws.title = "Residents & Registrars"
    build_roster(ws)
    build_leave_log(wb.create_sheet("Leave log"))
    readme = wb.create_sheet("About this file")
    for line in [
        "This file is written automatically by the Haematology Rota website after every change.",
        "Please don't edit it here: changes will be overwritten. Make changes on the website instead.",
        "All names in this sample are fictional.",
    ]:
        readme.append([line])
    readme.column_dimensions["A"].width = 100
    wb.save(out)


if __name__ == "__main__":
    main(sys.argv[1])
