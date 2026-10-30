"""
Run this once to generate a blank professors.xlsx with the correct columns.
Then fill in your professor data.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

OUTPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "professors.xlsx")

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Professors"

# ── headers ────────────────────────────────────────────────────────────────
headers = ["Professor Name", "Email", "Department", "Sent", "Stop"]
header_fill  = PatternFill("solid", fgColor="1F4E79")   # dark blue
sent_fill    = PatternFill("solid", fgColor="C6EFCE")   # light green
stop_fill    = PatternFill("solid", fgColor="FFC7CE")   # light red

thin = Side(style="thin", color="AAAAAA")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

for i, h in enumerate(headers, 1):
    cell = ws.cell(row=1, column=i, value=h)
    cell.font      = Font(bold=True, color="FFFFFF", size=11)
    cell.fill      = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border    = border

# ── column widths ──────────────────────────────────────────────────────────
ws.column_dimensions["A"].width = 28   # Professor Name
ws.column_dimensions["B"].width = 35   # Email
ws.column_dimensions["C"].width = 25   # Department
ws.column_dimensions["D"].width = 10   # Sent
ws.column_dimensions["E"].width = 10   # Stop
ws.row_dimensions[1].height = 22

# ── sample rows ────────────────────────────────────────────────────────────
sample = [
    ["Dr. Jane Smith",   "j.smith@university.edu",   "Computer Science",    "",  ""],
    ["Prof. Alan Kumar", "a.kumar@university.edu",    "Electrical Eng.",     "",  ""],
    ["Dr. Maria Lopez",  "m.lopez@university.edu",    "Data Science",        "",  ""],
]

for r, row_data in enumerate(sample, 2):
    for c, val in enumerate(row_data, 1):
        cell = ws.cell(row=r, column=c, value=val)
        cell.alignment = Alignment(horizontal="center" if c >= 4 else "left",
                                    vertical="center")
        cell.border = border

# Freeze top row
ws.freeze_panes = "A2"

wb.save(OUTPUT)
print(f"Created: {OUTPUT}")
print("Fill in your professors and save before running email_sender.py")
