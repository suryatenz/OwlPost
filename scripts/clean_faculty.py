"""
clean_faculty.py
Reads the raw faculty Excel, removes non-professors / retirees / bad emails,
deduplicates, and writes a clean professors.xlsx ready for the email sender.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import os
import re
from collections import defaultdict

INPUT  = r"C:\Users\spraj\Downloads\Faculty List_final.xlsx"
OUTPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "professors.xlsx")

# ── filter rules ────────────────────────────────────────────────────────────

# Only keep these roles (TA = Teaching Assistant, SI = Supplemental Instruction)
KEEP_ROLES = {"PI"}

# Any rank that starts with these strings is excluded
EXCLUDE_RANK_PREFIXES = (
    "retiree",
    "post doctoral",
    "post-doctoral",
    "visiting scholar",      # typically no lab / not hiring
)

# Exact ranks to exclude (not real faculty)
EXCLUDE_RANKS_EXACT = {
    "peer facilitator",
    "teaching assistant",
    "teaching fellow",       # usually grad students — remove if you want only faculty
    "core lab manager",
    "sr researcher",
    "research scientist",
    "scholar-in-residence",
    "temporary faculty",
}

# Optional: keep only professors from specific AcadGroup numbers.
# Set to None to keep ALL departments.
KEEP_ACAD_GROUPS = None   # e.g. {23, 25, 35} to filter by dept code

# ── helpers ─────────────────────────────────────────────────────────────────

def is_valid_email(email):
    if not email:
        return False
    return bool(re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", str(email).strip()))

def clean_name(raw):
    """Convert 'LastName,FirstName MiddleName' → 'FirstName LastName'."""
    raw = str(raw).strip()
    if "," in raw:
        parts = raw.split(",", 1)
        last  = parts[0].strip()
        first = parts[1].strip()
        return f"{first} {last}"
    return raw

def should_exclude(role, rank):
    role_str = str(role or "").strip()
    rank_str = str(rank or "").strip().lower()

    if role_str not in KEEP_ROLES:
        return True, f"role={role_str}"

    for prefix in EXCLUDE_RANK_PREFIXES:
        if rank_str.startswith(prefix):
            return True, f"rank starts with '{prefix}'"

    if rank_str in EXCLUDE_RANKS_EXACT:
        return True, f"rank='{rank_str}'"

    return False, ""

# ── main ────────────────────────────────────────────────────────────────────

def main():
    print(f"Reading: {INPUT}")
    wb_in = openpyxl.load_workbook(INPUT)
    ws_in = wb_in.active

    kept     = []
    removed  = defaultdict(int)
    seen_emails = set()

    for row in ws_in.iter_rows(min_row=2, values_only=True):
        iid, name, email, role, rank, group, org = row

        # 1. Valid email?
        email_str = str(email or "").strip()
        if not is_valid_email(email_str):
            removed["no/invalid email"] += 1
            continue

        # 2. Duplicate email?
        email_lower = email_str.lower()
        if email_lower in seen_emails:
            removed["duplicate email"] += 1
            continue
        seen_emails.add(email_lower)

        # 3. Role / rank filter
        exclude, reason = should_exclude(role, rank)
        if exclude:
            removed[reason] += 1
            continue

        # 4. AcadGroup filter (optional)
        if KEEP_ACAD_GROUPS is not None and group not in KEEP_ACAD_GROUPS:
            removed[f"acad_group={group}"] += 1
            continue

        kept.append({
            "Professor Name": clean_name(name),
            "Email":          email_str,
            "Department":     str(org or "").strip(),
            "AcadGroup":      group,
            "Rank":           str(rank or "").strip(),
            "Sent":           "",
            "Stop":           "",
        })

    # ── write output ──────────────────────────────────────────────────────

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Professors"

    headers = ["Professor Name", "Email", "Department", "AcadGroup", "Rank", "Sent", "Stop"]

    # Styles
    header_fill = PatternFill("solid", fgColor="1F4E79")
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Header row
    for i, h in enumerate(headers, 1):
        c = ws_out.cell(row=1, column=i, value=h)
        c.font      = Font(bold=True, color="FFFFFF", size=11)
        c.fill      = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = border

    # Column widths
    widths = [30, 35, 22, 12, 30, 8, 8]
    for i, w in enumerate(widths, 1):
        ws_out.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws_out.row_dimensions[1].height = 22
    ws_out.freeze_panes = "A2"

    # Data rows
    alt_fill = PatternFill("solid", fgColor="F2F7FC")
    for r, prof in enumerate(kept, 2):
        row_fill = alt_fill if r % 2 == 0 else None
        for i, key in enumerate(headers, 1):
            c = ws_out.cell(row=r, column=i, value=prof[key])
            c.alignment = Alignment(horizontal="center" if i >= 6 else "left",
                                    vertical="center")
            c.border = border
            if row_fill:
                c.fill = row_fill

    wb_out.save(OUTPUT)

    # ── report ────────────────────────────────────────────────────────────
    print(f"\n{'='*50}")
    print(f"  CLEAN RESULTS")
    print(f"{'='*50}")
    print(f"  Total input rows : {ws_in.max_row - 1}")
    print(f"  Rows KEPT        : {len(kept)}")
    print(f"  Rows REMOVED     : {sum(removed.values())}")
    print(f"\n  Removed breakdown:")
    for reason, count in sorted(removed.items(), key=lambda x: -x[1]):
        print(f"    {count:>5}  {reason}")
    print(f"\n  Output saved to  : {OUTPUT}")
    print(f"{'='*50}\n")


if __name__ == "__main__":
    main()
