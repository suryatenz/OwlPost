"""
reply_checker.py
Connects to Gmail via IMAP, scans your inbox for replies from professors
in professors.xlsx, and automatically marks Stop = Yes for anyone who replied.
Run this anytime, or it runs automatically before each daily email batch.
"""

import imaplib
import email
import openpyxl
import json
import os
from datetime import datetime

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
LOG_FILE    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_log.txt")


def load_config():
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def log(msg):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def get_column_map(ws):
    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
    return {h: i for i, h in enumerate(headers)}

def is_yes(value):
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ("yes", "true", "1", "x", "sent")


def check_replies():
    config      = load_config()
    excel_path  = os.path.join(os.path.dirname(CONFIG_FILE), config["excel_path"])
    gmail_user  = config["sender_email"]
    gmail_pass  = config["sender_app_password"]

    if not os.path.exists(excel_path):
        log("reply_checker: Excel not found, skipping.")
        return

    # ── load spreadsheet ──────────────────────────────────────────────────
    wb  = openpyxl.load_workbook(excel_path)
    ws  = wb.active
    col = get_column_map(ws)

    # Build lookup maps:
    # 1. full email (lowercase)  -> row   e.g. "mkidd@uh.edu"
    # 2. username only           -> row   e.g. "mkidd"  (handles @Central.UH.EDU aliases)
    # 3. normalized full name    -> row   e.g. "reginamdcapuano" (handles dept aliases)
    prof_by_email    = {}
    prof_by_username = {}
    prof_by_name     = {}

    import re
    def normalize_name(name):
        return re.sub(r'[^a-z]', '', name.lower())

    for row in ws.iter_rows(min_row=2):
        email_val = row[col["Email"]].value
        name_val  = row[col["Professor Name"]].value
        if email_val:
            full = str(email_val).strip().lower()
            prof_by_email[full] = row
            prof_by_username[full.split("@")[0]] = row
        if name_val:
            prof_by_name[normalize_name(str(name_val))] = row

    if not prof_by_email:
        log("reply_checker: No professors found in spreadsheet.")
        return

    log(f"reply_checker: Checking Gmail inbox for replies from {len(prof_by_email)} professors...")

    # ── connect to Gmail IMAP ─────────────────────────────────────────────
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(gmail_user, gmail_pass)
        mail.select("inbox")
    except Exception as e:
        log(f"reply_checker: Could not connect to Gmail IMAP — {e}")
        return

    # Search emails since April 1 2026 (when sending started)
    # To do a full inbox scan run: python reply_checker.py --all
    import sys
    if "--all" in sys.argv:
        _, msg_ids = mail.search(None, "ALL")
    else:
        _, msg_ids = mail.search(None, 'SINCE "01-Apr-2026"')
    all_ids = msg_ids[0].split()

    log(f"reply_checker: Scanning {len(all_ids)} emails in inbox...")

    marked_count = 0
    already_stopped = 0

    for msg_id in all_ids:
        try:
            _, msg_data = mail.fetch(msg_id, "(RFC822)")
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)

            # Get sender email address
            from_header = msg.get("From", "")
            # Extract just the email part from "Name <email@domain.com>"
            sender_email = ""
            if "<" in from_header and ">" in from_header:
                sender_email = from_header.split("<")[1].split(">")[0].strip().lower()
            else:
                sender_email = from_header.strip().lower()

            # Match 1: full email  e.g. mkidd@uh.edu
            # Match 2: username   e.g. mkidd (handles @Central.UH.EDU vs @uh.edu)
            # Match 3: name       e.g. "Capuano, Regina M" -> normalize and lookup
            sender_username = sender_email.split("@")[0]

            # Extract display name from "LastName, FirstName <email>"
            # Rearrange "Last, First M" -> "First M Last" to match spreadsheet format
            sender_name_norm = ""
            if "<" in from_header:
                display = from_header.split("<")[0].strip().strip('"')
                if "," in display:
                    parts = display.split(",", 1)
                    display = parts[1].strip() + " " + parts[0].strip()
                sender_name_norm = normalize_name(display)

            row = (prof_by_email.get(sender_email)
                   or prof_by_username.get(sender_username)
                   or (prof_by_name.get(sender_name_norm) if sender_name_norm else None))

            if row:
                prof_name = str(row[col["Professor Name"]].value or "").strip()

                # Already stopped?
                if is_yes(row[col["Stop"]].value):
                    already_stopped += 1
                    continue

                # Mark Stop = Yes
                row[col["Stop"]].value = "Yes"
                wb.save(excel_path)
                marked_count += 1
                log(f"  REPLY DETECTED -> {prof_name} <{sender_email}> marked Stop = Yes")

        except Exception:
            continue

    mail.logout()

    if marked_count == 0:
        log(f"reply_checker: No new replies found. ({already_stopped} already marked Stop)")
    else:
        log(f"reply_checker: Done. {marked_count} new replies marked. {already_stopped} were already stopped.")


if __name__ == "__main__":
    check_replies()
