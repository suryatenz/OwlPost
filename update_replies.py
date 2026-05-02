"""
update_replies.py
Processes all professor email replies and updates professors.xlsx:
  1. Adds a "Replied" column
  2. Marks who replied (positive/negative)
  3. Sets Stop=Yes for negative replies
  4. Deletes professors who left UH entirely
  5. Adds new contacts referred by professors
  6. Flags bounced emails
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
from openpyxl.utils import get_column_letter
import os

EXCEL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "professors.xlsx")

# ── REPLY DATA ─────────────────────────────────────────────────────────────────
# Key = email as it appears in professors.xlsx (the TO address from outreach)
# "replied" = column value, "stop" = whether to also set Stop=Yes

POSITIVE_REPLIES = {
    # Replied with interest / referral to hiring manager / invited meeting
    "ejjones@uh.edu":           "Yes - Invited for Zoom interview",
    "dkalliontzis@uh.edu":      "Yes - Referred to lab website to apply",
    "dkallion@central.uh.edu":  "Yes - Referred to lab website to apply",
    "akhan7@uh.edu":            "Yes - Asked about publication, possible interest",
    "dtonofrei@uh.edu":         "Yes - Invited to office meeting",
    "sclarsen@uh.edu":          "Yes - Dean of Grad School, forwarded to hiring mgr",
    "sclarsen@central.uh.edu":  "Yes - Dean of Grad School, forwarded to hiring mgr",
    "xiamt@central.uh.edu":     "Yes - Open to discussing research ideas",
    "xiamt@uh.edu":             "Yes - Open to discussing research ideas",
}

NEGATIVE_REPLIES = {
    # Emailed address -> reply note (all get Stop=Yes)
    "szkamal@uh.edu":           "Yes - No openings",
    "jwbreaux@uh.edu":          "Yes - Nothing at this time",
    "prroberts@uh.edu":         "Yes - Not able to hire",
    "hdomjan@uh.edu":           "Yes - No job openings",
    "dthomas15@uh.edu":         "Yes - Not looking for interns",
    "adnordin@central.uh.edu":  "Yes - No paid positions",
    "hsalloum@central.uh.edu":  "Yes - Not working with students",
    "tjmountz@central.uh.edu":  "Yes - No openings",
    "mhhanke@uh.edu":           "Yes - Not a good fit",
    "wfu3@uh.edu":              "Yes - No positions",
    "iportillovazquez@uh.edu":  "Yes - No position to offer",
    "rglarsen@uh.edu":          "Yes - Not aware of anything",
    "nmalamataris@uh.edu":      "Yes - No availability",
    "czhang59@central.uh.edu":  "Yes - No extra funding",
    "hskrishn@uh.edu":          "Yes - No openings",
    "mkaspar2@uh.edu":          "Yes - TAs assigned by dept",
    "jnortiz5@uh.edu":          "Yes - No opportunities",
    "pmann@uh.edu":             "Yes - No RA positions",
    "ychoi6@uh.edu":            "Yes - Limited funding",
    "wwsager@uh.edu":           "Yes - Retiring, not taking students",
    "zren@uh.edu":              "Yes - No liberty to offer positions",
    "maperepelitsa@uh.edu":     "Yes - No positions",
    "srgilbertson@uh.edu":      "Yes - Professor Emeritus, no positions",
    "jhthurmond@uh.edu":        "Yes - Only temporary at UH",
    "mogletree@uh.edu":         "Yes - No positions",
    "vbsisson@uh.edu":          "Yes - EAS students only",
    "bfernandes@uh.edu":        "Yes - Adjunct, cannot assign TAs",
    "krcastle@central.uh.edu":  "Yes - No funding available",
    "cmneuhau@central.uh.edu":  "Yes - Full-time administrator",
    "sscott2@uh.edu":           "Yes - No permanent position at UH",
    "tqdang@uh.edu":            "Yes - Cannot take RA/TA/IA",
    "ckweber@uh.edu":           "Yes - No position",
    "jreed@uh.edu":             "Yes - Fine arts, wrong dept",
    "ddjohnson7@uh.edu":        "Yes - Directed to HR/COE",
    "besorensen@uh.edu":        "Yes - No RA/TA funds in economics",
    "mdchance@uh.edu":          "Yes - No positions",
    "bdauwalder@uh.edu":        "Yes - No positions",
    "mtblake@uh.edu":           "Yes - No open RA positions",
    "pcirino@uh.edu":           "Yes - Psychologist, wrong dept",
    "trlee@uh.edu":             "Yes - No RA/TA this year",
    "pahoward@uh.edu":          "Yes - History dept, wrong dept",
    "spennings@uh.edu":         "Yes - No positions",
    "acarrera@uh.edu":          "Yes - Languages dept",
    "svenkat@uh.edu":           "Yes - Part-time, not in Texas",
    "mmurphy4@uh.edu":          "Yes - No TAs or RAs",
    "long@uh.edu":              "Yes - No positions",
    "hwagan@uh.edu":            "Yes - No openings",
    "kbrem@uh.edu":             "Yes - Law Center",
    "gtaylor@uh.edu":           "Yes - All positions filled",
    "bolson@uh.edu":            "Yes - Nothing that fits",
    "ggillan@uh.edu":           "Yes - College hires TAs directly",
    "ltwood@uh.edu":            "Yes - No positions",
    "carlos@uh.edu":            "Yes - PhD only, no MS/non-CS support",
    "ljhs@uh.edu":              "Yes - No opportunities",
    "dpapell@uh.edu":           "Yes - Nothing",
    "capuano@uh.edu":           "Yes - No position",
    "lwilliams@uh.edu":         "Yes - No position",
    "ajjacob@uh.edu":           "Yes - Nothing",
    "shattuck@uh.edu":          "Yes - No assistantships",
    "dpearson@uh.edu":          "Yes - No opportunities",
    "ppeters@uh.edu":           "Yes - No position",
    "asekar4@central.uh.edu":   "Yes - Referred to CS TA coordinator",
    "dpattison@uh.edu":         "Yes - Referred to Math grad director",
    "thkazemi@central.uh.edu":  "Yes - Referred to BME grad advisor",
    "scpennin@central.uh.edu":  "Yes - No positions",
    # Already deleted from spreadsheet but marking for completeness
    "fhilterman@uh.edu":        "Yes - 84 yrs old, no research",
    "wdupre@uh.edu":            "Yes - Retired",
    "mfleischer@uh.edu":        "Yes - Retired",
    "jgarson@uh.edu":           "Yes - Retired",
    "srasmussen@uh.edu":        "Yes - Passed away April 2025",
    "srobbins@uh.edu":          "Yes - Retired (Emerita)",
    "rstorey@uh.edu":           "Yes - Retired for years",
    "ltrombetta@uh.edu":        "Yes - Retired (Emeritus)",
    "jjohnso2@uh.edu":          "Yes - Retired",
    "jjwaite@uh.edu":           "Yes - Retired",
    "semoreno@uh.edu":          "Yes - Retired Jan 2025",
    "bweber@uh.edu":            "Yes - No longer Music faculty",
    "jelawren@uh.edu":          "Yes - No longer at Law Center",
    "bdindoru@uh.edu":          "Yes - Left UH for Texas A&M",
    "mkidd@uh.edu":             "Yes - Transitioned from UH",
    "sschroed@uh.edu":          "Yes - Retired after 40 yrs",
}

# ── DELETE these (left UH entirely) ───────────────────────────────────────────
DELETE_NAME_PREFIXES = [
    "zwortzel",    # Zachary Wortzel - no longer at Bauer College
    "lmckinney",   # McKinney - no longer at UH (15 yrs, now gone)
    "eemay",       # Elizabeth May - lab moved to UW-Madison
]
# Also delete any emails we know exactly
DELETE_EXACT_EMAILS = {
    "zwortzel@bauer.uh.edu",
    "lmckinney@central.uh.edu",
    "eemay@central.uh.edu",
}

# ── BOUNCE (address not found) ─────────────────────────────────────────────────
BOUNCE_EMAILS = {
    "kglate@uh.edu": "Bounced - actually at NASA",
}

# ── NEW contacts referred by professors ───────────────────────────────────────
NEW_CONTACTS = [
    {
        "name":  "Carlos Porras",
        "email": "cporras@bauer.uh.edu",
        "dept":  "Bauer College of Business",
        "note":  "HIRING 2 TA positions Fall semester (BUSI 3302)",
    },
]


def main():
    wb = openpyxl.load_workbook(EXCEL)
    ws = wb.active

    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
    col = {h: i for i, h in enumerate(headers)}

    # ── Add Replied column if missing ─────────────────────────────────────────
    if "Replied" not in col:
        replied_col_idx = len(headers)
        ws.cell(row=1, column=replied_col_idx + 1, value="Replied")
        headers.append("Replied")
        col["Replied"] = replied_col_idx
        print("  Added 'Replied' column")
    else:
        print("  'Replied' column already exists")

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Fills for color-coding
    green_fill  = PatternFill("solid", fgColor="C6EFCE")  # positive reply
    yellow_fill = PatternFill("solid", fgColor="FFEB9C")  # negative reply
    red_fill    = PatternFill("solid", fgColor="FFC7CE")  # bounce

    stats = {
        "positive": 0, "negative": 0,
        "deleted": 0, "bounced": 0,
        "new_added": 0,
    }

    rows_to_delete = []

    for row in ws.iter_rows(min_row=2):
        email_cell = row[col["Email"]]
        if not email_cell.value:
            continue
        email = str(email_cell.value).strip().lower()

        # Check DELETE (left UH)
        prefix = email.split("@")[0].lower()
        if email in DELETE_EXACT_EMAILS or prefix in DELETE_NAME_PREFIXES:
            rows_to_delete.append(row[0].row)
            stats["deleted"] += 1
            continue

        replied_cell = row[col["Replied"]]
        stop_cell    = row[col["Stop"]]
        name = str(row[col["Professor Name"]].value or "").strip()

        # Positive reply
        if email in POSITIVE_REPLIES:
            replied_cell.value = POSITIVE_REPLIES[email]
            replied_cell.fill  = green_fill
            replied_cell.alignment = Alignment(horizontal="left", vertical="center")
            stats["positive"] += 1
            print(f"  [+] {name} <{email}>")
            continue

        # Negative reply
        if email in NEGATIVE_REPLIES:
            replied_cell.value = NEGATIVE_REPLIES[email]
            replied_cell.fill  = yellow_fill
            replied_cell.alignment = Alignment(horizontal="left", vertical="center")
            stop_cell.value = "Yes"
            stats["negative"] += 1
            continue

        # Bounce
        if email in BOUNCE_EMAILS:
            replied_cell.value = f"BOUNCED - {BOUNCE_EMAILS[email]}"
            replied_cell.fill  = red_fill
            replied_cell.alignment = Alignment(horizontal="left", vertical="center")
            stop_cell.value = "Yes"
            stats["bounced"] += 1
            print(f"  [!] BOUNCE: {name} <{email}>")
            continue

    # Delete rows in reverse order
    for row_num in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(row_num)

    # ── Add new contacts ──────────────────────────────────────────────────────
    existing_emails = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[col["Email"]]:
            existing_emails.add(str(row[col["Email"]]).strip().lower())

    for c in NEW_CONTACTS:
        email = c["email"].strip().lower()
        if email in existing_emails:
            print(f"  [=] Already exists: {c['name']} <{email}>")
            continue
        values = {
            "Professor Name": c["name"],
            "Email":          email,
            "Department":     c["dept"],
            "AcadGroup":      "",
            "Rank":           "",
            "Sent":           "",
            "Stop":           "",
            "Replied":        f"Yes - {c['note']}",
        }
        r = ws.max_row + 1
        for i, key in enumerate(headers, 1):
            cell = ws.cell(row=r, column=i, value=values.get(key, ""))
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = border
        # Green fill for Replied cell
        ws.cell(row=r, column=col["Replied"] + 1).fill = green_fill
        existing_emails.add(email)
        stats["new_added"] += 1
        print(f"  [NEW] {c['name']} <{email}> — {c['note']}")

    wb.save(EXCEL)

    print(f"\n{'='*60}")
    print(f"  Positive replies marked  : {stats['positive']}")
    print(f"  Negative replies / Stop  : {stats['negative']}")
    print(f"  Bounced emails flagged   : {stats['bounced']}")
    print(f"  Rows deleted (left UH)   : {stats['deleted']}")
    print(f"  New contacts added       : {stats['new_added']}")
    print(f"  Saved to                 : {EXCEL}")
    print(f"{'='*60}\n")

    # Quick count of what remains active
    wb2 = openpyxl.load_workbook(EXCEL)
    ws2 = wb2.active
    h2 = [str(c.value).strip() if c.value else "" for c in ws2[1]]
    c2 = {h: i for i, h in enumerate(h2)}
    active = 0
    for row in ws2.iter_rows(min_row=2, values_only=True):
        stop_val    = str(row[c2["Stop"]]    or "").strip().lower()
        replied_val = str(row[c2["Replied"]] or "").strip().lower()
        if stop_val not in ("yes", "true", "1") and not replied_val.startswith("yes - no") \
           and not replied_val.startswith("bounced"):
            active += 1
    print(f"  Active (not stopped, not replied-no): {active}")
    print(f"  Total rows in file                  : {ws2.max_row - 1}\n")


if __name__ == "__main__":
    main()
