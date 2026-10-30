"""
add_deans.py
Adds college deans and department chairs to professors.xlsx.
Skips any email already present. Marks Rank = "Dean" or "Chair".
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Side
import os

EXCEL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "professors.xlsx")

LEADERS = [
    # ── College Deans ──────────────────────────────────────────────────────────
    # Cullen already in spreadsheet: psharma@uh.edu (Dean)
    # Grad School Dean already in spreadsheet: sclarsen@uh.edu / sclarsen@central.uh.edu
    {"name": "Gregory B. Dudley",   "email": "gbdudley@uh.edu",            "dept": "NSM",                              "rank": "Dean"},
    {"name": "Xianjun Geng",        "email": "xgeng@bauer.uh.edu",         "dept": "Bauer College of Business",        "rank": "Dean"},
    {"name": "Catherine Horn",      "email": "clhorn2@uh.edu",             "dept": "College of Education",             "rank": "Dean"},
    {"name": "Heidar Malki",        "email": "malki@uh.edu",               "dept": "Technology Division",              "rank": "Dean"},

    # ── NSM Department Chairs ─────────────────────────────────────────────────
    {"name": "Shaun Zhang",         "email": "shaunzhang@uh.edu",          "dept": "NSM - Biology & Biochemistry",     "rank": "Chair"},
    {"name": "David Hoffman",       "email": "hoffman@uh.edu",             "dept": "NSM - Chemistry",                  "rank": "Chair"},
    {"name": "Lennart Johnsson",    "email": "ljohnsson@uh.edu",           "dept": "NSM - Computer Science",           "rank": "Chair"},
    {"name": "Thomas Lapen",        "email": "tjlapen@uh.edu",             "dept": "NSM - Earth & Atmospheric Sciences","rank": "Chair"},
    {"name": "Bernhard Bodmann",    "email": "bgb@uh.edu",                 "dept": "NSM - Mathematics",                "rank": "Chair"},
    {"name": "Kevin Bassler",       "email": "bassler@uh.edu",             "dept": "NSM - Physics",                    "rank": "Chair"},

    # ── Cullen Engineering Dept Chairs ────────────────────────────────────────
    # ECE jchen18@uh.edu, ME karolos@uh.edu, IE ginolim@uh.edu,
    # ChE tjmountz@central.uh.edu already in spreadsheet
    {"name": "Dimitrios Hatzignatiou", "email": "dghatzignatious@uh.edu",  "dept": "Cullen - Petroleum Engineering",   "rank": "Chair"},
    {"name": "Kirill Larin",        "email": "klarin@uh.edu",              "dept": "Cullen - Biomedical Engineering",  "rank": "Chair"},
    {"name": "Roberto Ballarini",   "email": "rballari@central.uh.edu",    "dept": "Cullen - Civil Engineering",       "rank": "Chair"},

    # ── College of Education Dept Chairs ─────────────────────────────────────
    {"name": "Kristen Hassett",     "email": "khassett@uh.edu",            "dept": "Education - ELPS",                 "rank": "Chair"},
    {"name": "Milena Keller-Margulis","email": "mkmargulis@uh.edu",        "dept": "Education - PHLS",                 "rank": "Chair"},
    {"name": "Samuel Brower",       "email": "srbrower@uh.edu",            "dept": "Education - Curriculum & Instruction","rank": "Chair"},

    # ── Technology Division Dept Chairs ──────────────────────────────────────
    {"name": "Fatima Merchant",     "email": "fmerchant@uh.edu",           "dept": "Technology - Engineering Technology","rank": "Chair"},
    {"name": "Ahmed Senouci",       "email": "asenouci@uh.edu",            "dept": "Technology - Construction Mgmt",   "rank": "Chair"},
    {"name": "Amaury Lendasse",     "email": "alendasse@uh.edu",           "dept": "Technology - ILT",                 "rank": "Chair"},
    {"name": "Barbara L. Stewart",  "email": "blstewart@uh.edu",           "dept": "Technology - HDCS",                "rank": "Chair"},
]


def main():
    wb = openpyxl.load_workbook(EXCEL)
    ws = wb.active

    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
    col = {h: i for i, h in enumerate(headers)}

    existing = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[col["Email"]]:
            existing.add(str(row[col["Email"]]).strip().lower())

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    added = 0
    skipped = 0

    for leader in LEADERS:
        email = leader["email"].strip().lower()
        if email in existing:
            print(f"  [=] Already present: {leader['name']} <{email}>")
            skipped += 1
            continue

        values = {
            "Professor Name": leader["name"],
            "Email":          email,
            "Department":     leader["dept"],
            "AcadGroup":      "",
            "Rank":           leader["rank"],
            "Sent":           "",
            "Stop":           "",
        }
        if "Replied" in col:
            values["Replied"] = ""

        r = ws.max_row + 1
        for i, key in enumerate(headers, 1):
            cell = ws.cell(row=r, column=i, value=values.get(key, ""))
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = border

        existing.add(email)
        added += 1
        print(f"  [+] {leader['rank']:5s}  {leader['name']} <{email}>")

    wb.save(EXCEL)

    print(f"\n{'='*60}")
    print(f"  Deans/chairs added   : {added}")
    print(f"  Already present      : {skipped}")
    print(f"  Saved to             : {EXCEL}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
