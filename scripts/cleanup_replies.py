"""
cleanup_replies.py
Three-stage cleanup:
  1. DELETE professors who retired / passed away / left UH entirely
  2. STOP professors who replied saying no positions / wrong department
  3. DELETE all professors from departments that cannot offer RA/TA/IA
     to an Engineering Data Science graduate student
"""

import openpyxl
import os

EXCEL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "professors.xlsx")

# ── Stage 1: DELETE completely (retired / deceased / left UH) ─────────────────
DELETE_EMAILS = {
    # Confirmed from email replies or known status
    "jbelinne@uh.edu",       # Retired (Bauer)
    "wdupre@uh.edu",         # Retired
    "mfleischer@uh.edu",     # Retired
    "jgarson@uh.edu",        # Retired
    "fhilterman@uh.edu",     # 84 yrs old, no active research
    "srasmussen@uh.edu",     # Passed away April 13, 2025
    "bdindoru@uh.edu",       # Left UH — now at Texas A&M
    "jjohnso2@uh.edu",       # Retired (Theatre & Dance)
    "semoreno@uh.edu",       # Retired Jan 2025
    "bweber@uh.edu",         # No longer UH Music faculty
    "jelawren@uh.edu",       # No longer at Law Center
    "jjwaite@uh.edu",        # Retired (Digital Media)
    "ltrombetta@uh.edu",     # Retired, Professor Emeritus (ECE)
    "srobbins@uh.edu",       # Retired, Professor Emerita (Social Work)
    "rstorey@uh.edu",        # Retired for several years (Anthropology)
    "sschroed@uh.edu",       # Retired after 40 yrs teaching Mathematics
    "mkidd@uh.edu",          # Transitioned out of UH (no longer faculty)
}

# ── Stage 2: STOP (replied — no positions / wrong dept / not relevant) ────────
STOP_EMAILS = {
    # --- Already in list ---
    "jreed@uh.edu",          # Fine arts, not engineering
    "ddjohnson7@uh.edu",     # Directed to HR/COE web pages
    "besorensen@uh.edu",     # No RA/TA funds (Economics)
    "mdchance@uh.edu",       # No positions (IST, Lecturer)
    "bdauwalder@uh.edu",     # No positions
    "pcirino@uh.edu",        # Psychologist, wrong dept
    "rboswell@uh.edu",       # English dept
    "pahoward@uh.edu",       # History dept
    "spennings@uh.edu",      # No positions (Biology)
    "acarrera@uh.edu",       # Languages dept
    "svenkat@uh.edu",        # Part-time, not in Texas
    "mmurphy4@uh.edu",       # No positions (Earth & Atmospheric Sciences)
    "ltwood@uh.edu",         # No positions
    "bolson@uh.edu",         # Doesn't fit skill set
    "ggillan@uh.edu",        # College of Education
    "long@uh.edu",           # No openings (ECE)
    "hwagan@uh.edu",         # No openings
    "kbrem@uh.edu",          # Law Center
    "trlee@uh.edu",          # No RA/TA this year
    "gtaylor@uh.edu",        # All positions filled (Eng Technology)
    "mtblake@uh.edu",        # No open RA positions / on leave (CLASS/COMD)
    # --- New from April 2026 replies ---
    "carlos@uh.edu",         # ECE — assistantships for PhD only, no MS support
    "ckweber@uh.edu",        # Bauer College — no position
    "ppeters@uh.edu",        # No position
    "dpearson@uh.edu",       # Health & Human Performance, no positions
    "lwilliams@uh.edu",      # Pharmacy — not relevant dept, no positions
    "ajjacob@uh.edu",        # Chemistry — no positions
    "shattuck@uh.edu",       # ECE — no assistantships to offer
    "ljhs@uh.edu",           # IE — no opportunities at this time
    "dpapell@uh.edu",        # Economics — no positions
    "capuano@uh.edu",        # Geosciences — no position, none expected
}

# ── Stage 3: DELETE entire departments not relevant to Eng Data Science ────────
# Keep: ECE, ME, BME, ChE, Civil, IE, Pete, CS, Math, Chemistry,
#        Physics, Biochemistry, Geosciences, Eng Technology
# Delete everything else
DELETE_DEPARTMENTS = {
    # McGovern College of the Arts
    "H230ART",    # Art
    "H230MUS",    # Music
    "H230THE",    # Theatre & Dance
    # Hines College of Architecture
    "H15",        # Architecture & Design
    "H23",        # Architecture
    # College of Liberal Arts & Social Sciences (CLASS)
    "H46",        # CLASS general
    "H460COM",    # Communication
    "H460ENG",    # English
    "H460ECO",    # Economics
    "H460HHP",    # Health & Human Performance
    "H460HIS",    # History
    "H460HSP",    # Hispanic Studies
    "H460MCL",    # Modern & Classical Languages
    "H460PHI",    # Philosophy
    "H460POL",    # Political Science
    "H460PSY",    # Psychology (CLASS — NOT College of Education PHLS)
    "H460SOC",    # Sociology
    "H460CCS",    # Comparative Cultural Studies
    "H460CMD",    # Communication Sciences & Disorders
    # Law Center
    "H450LAW",    # Law
    # Hilton College of Hotel & Restaurant Management
    "H40",        # Hotel & Restaurant Management
    "H400HRM",    # Human Resource Management
    # College of Optometry
    "H500OPT",    # Optometry
    # College of Pharmacy
    "H550PHA",    # Pharmacy
    "H550CSA",    # Communication Sciences & Arts
    "H550PPS",    # Pharmacological & Pharmaceutical Sciences
    # Graduate College of Social Work
    "H600S_W",    # Social Work
    # Anthropology (standalone)
    "H70",        # Anthropology
    # Misc / non-academic
    "H35",        # Communication / Media Studies
    "H37",        # Business admin / non-dept
    "H48",        # Misc
    "H90",        # Student Affairs / non-faculty
    # NOTE: intentionally NOT in this list (these are KEPT):
    # Bauer College of Business : H250ACC, H250DIS, H250FIN, H250MGT, H250MKT, H25
    # College of Education      : H300C&I, H300ELPS, H300EPY, H30
    # Technology Division       : H650CNST, H650HEC, H650ILT, H650E_T
    # Cullen Engineering        : H350BIOE, H350CHE, H350C_E, H350ECE, H350I_E, H350M_E, H350PET
    # NSM                       : H470BCH, H470CHM, H470CSC, H470GEO, H470MTH, H470PHY, H47
}


def main():
    wb = openpyxl.load_workbook(EXCEL)
    ws = wb.active

    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
    col = {h: i for i, h in enumerate(headers)}

    stage1_delete   = []   # email-based deletes
    stage2_stopped  = []   # newly marked Stop
    stage2_already  = []   # already stopped
    stage3_delete   = []   # dept-based deletes

    rows_to_delete = []

    for row in ws.iter_rows(min_row=2):
        email_val = row[col["Email"]].value
        if not email_val:
            continue
        email = str(email_val).strip().lower()
        name  = str(row[col["Professor Name"]].value or "").strip()
        dept  = str(row[col["Department"]].value or "").strip()

        # Stage 1 — hard delete by email
        if email in DELETE_EMAILS:
            rows_to_delete.append(row[0].row)
            stage1_delete.append(f"{name} <{email}>")
            continue

        # Stage 3 — delete by irrelevant department
        if dept in DELETE_DEPARTMENTS:
            rows_to_delete.append(row[0].row)
            stage3_delete.append(f"{name} <{email}> [{dept}]")
            continue

        # Stage 2 — mark Stop by email
        if email in STOP_EMAILS:
            if str(row[col["Stop"]].value or "").strip().lower() not in ("yes", "true", "1"):
                row[col["Stop"]].value = "Yes"
                stage2_stopped.append(f"{name} <{email}>")
            else:
                stage2_already.append(f"{name} <{email}>")

    # Delete in reverse order to keep row indices valid
    for row_num in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(row_num)

    wb.save(EXCEL)

    # ── Report ───────────────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print(f"  CLEANUP RESULTS")
    print(f"{'='*60}")

    print(f"\n  STAGE 1 — DELETED ({len(stage1_delete)}) retired / deceased / left UH:")
    for p in stage1_delete:
        print(f"    - {p}")

    print(f"\n  STAGE 2 — MARKED STOP ({len(stage2_stopped)}) no positions / wrong dept:")
    for p in stage2_stopped:
        print(f"    - {p}")

    if stage2_already:
        print(f"\n  STAGE 2 — ALREADY STOPPED ({len(stage2_already)}):")
        for p in stage2_already:
            print(f"    - {p}")

    print(f"\n  STAGE 3 — DELETED ({len(stage3_delete)}) irrelevant department professors:")
    # Group by dept
    from collections import Counter
    dept_counts = Counter(p.split("[")[-1].rstrip("]") for p in stage3_delete)
    for dept, count in sorted(dept_counts.items(), key=lambda x: -x[1]):
        print(f"    {count:4d}  {dept}")

    total_deleted = len(stage1_delete) + len(stage3_delete)
    print(f"\n  Total rows deleted  : {total_deleted}")
    print(f"  Total newly stopped : {len(stage2_stopped)}")
    print(f"  Saved to            : {EXCEL}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()
