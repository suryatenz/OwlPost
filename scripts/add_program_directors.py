"""
add_program_directors.py
Adds Graduate Program Directors / Department Chairs for STEM + Bauer.
Skips any email already present in professors.xlsx.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Side
import os

EXCEL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "professors.xlsx")

DIRECTORS = [
    # ── ECE ───────────────────────────────────────────────────────────────────
    {"name": "David R. Jackson",    "email": "djackson@uh.edu",         "dept": "ECE",                          "rank": "Grad Director"},

    # ── Mechanical & Aerospace Engineering ───────────────────────────────────
    {"name": "Dong Liu",            "email": "dongliu@uh.edu",          "dept": "Mechanical Engineering",       "rank": "Grad Director"},
    {"name": "Shailendra Joshi",    "email": "spjoshi3@central.uh.edu", "dept": "Mechanical Engineering",       "rank": "Grad Director"},

    # ── Industrial & Systems Engineering ─────────────────────────────────────
    {"name": "Qianmei Feng",        "email": "qfeng@central.uh.edu",    "dept": "Industrial Engineering",       "rank": "Grad Director"},
    {"name": "Zheyong Bian",        "email": "zbian2@central.uh.edu",   "dept": "Industrial Engineering",       "rank": "Director"},
    {"name": "Matthew Hu",          "email": "mhu9@central.uh.edu",     "dept": "Industrial Engineering",       "rank": "Director"},
    {"name": "Dennis Sherman",      "email": "dlsherma@central.uh.edu", "dept": "Industrial Engineering",       "rank": "Director"},

    # ── Chemical & Biomolecular Engineering ───────────────────────────────────
    {"name": "Jeffrey Rimer",       "email": "jrimer@central.uh.edu",   "dept": "Chemical Engineering",         "rank": "Grad Director"},

    # ── Civil & Environmental Engineering ────────────────────────────────────
    {"name": "Keh-Han Wang",        "email": "khwang@uh.edu",           "dept": "Civil Engineering",            "rank": "Grad Director"},
    {"name": "Hyongki Lee",         "email": "hlee45@central.uh.edu",   "dept": "Civil Engineering",            "rank": "Grad Director"},

    # ── Mathematics ───────────────────────────────────────────────────────────
    {"name": "Alan Haynes",         "email": "akhaynes@central.uh.edu", "dept": "NSM - Mathematics",            "rank": "Grad Director"},
    {"name": "Jennifer May",        "email": "jrmay@uh.edu",            "dept": "NSM - Mathematics",            "rank": "Grad Director"},

    # ── Computer Science (NSM) ────────────────────────────────────────────────
    {"name": "Zhigang Deng",        "email": "zdeng4@uh.edu",           "dept": "NSM - Computer Science",       "rank": "Grad Director"},

    # ── Physics ───────────────────────────────────────────────────────────────
    {"name": "Oomman K. Varghese",  "email": "okvarghese@uh.edu",       "dept": "NSM - Physics",                "rank": "Grad Advisor"},

    # ── Biology & Biochemistry ────────────────────────────────────────────────
    {"name": "Tasneem Bawa",        "email": "tbawakha@central.uh.edu", "dept": "NSM - Biology & Biochemistry", "rank": "Assoc Chair Grad"},

    # ── Bauer College of Business ─────────────────────────────────────────────
    {"name": "Praveen Kumar",       "email": "pkumar@uh.edu",           "dept": "Bauer - Finance",              "rank": "Chair"},
    {"name": "Kris Jacobs",         "email": "kjacobs@bauer.uh.edu",    "dept": "Bauer - Finance",              "rank": "Chair"},
    {"name": "Steve Werner",        "email": "werner@bauer.uh.edu",     "dept": "Bauer - Management",           "rank": "Chair"},
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

    for d in DIRECTORS:
        email = d["email"].strip().lower()
        if email in existing:
            print(f"  [=] Already present: {d['name']} <{email}>")
            skipped += 1
            continue

        values = {
            "Professor Name": d["name"],
            "Email":          email,
            "Department":     d["dept"],
            "AcadGroup":      "",
            "Rank":           d["rank"],
            "Sent":           "",
            "Stop":           "",
            "Replied":        "",
        }

        r = ws.max_row + 1
        for i, key in enumerate(headers, 1):
            cell = ws.cell(row=r, column=i, value=values.get(key, ""))
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = border

        existing.add(email)
        added += 1
        print(f"  [+] {d['rank']:<18} {d['name']} <{email}>")

    wb.save(EXCEL)
    print(f"\n{'='*60}")
    print(f"  Program directors added  : {added}")
    print(f"  Already present          : {skipped}")
    print(f"  Saved to                 : {EXCEL}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
