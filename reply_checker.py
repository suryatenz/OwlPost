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

    # Build a map of professor email (lowercase) -> row
    prof_rows = {}
    for row in ws.iter_rows(min_row=2):
        email_val = row[col["Email"]].value
        if email_val:
            prof_rows[str(email_val).strip().lower()] = row

    if not prof_rows:
        log("reply_checker: No professors found in spreadsheet.")
        return

    log(f"reply_checker: Checking Gmail inbox for replies from {len(prof_rows)} professors...")

    # ── connect to Gmail IMAP ─────────────────────────────────────────────
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(gmail_user, gmail_pass)
        mail.select("inbox")
    except Exception as e:
        log(f"reply_checker: Could not connect to Gmail IMAP — {e}")
        return

    # Search only emails received after the first email was sent (last 90 days)
    _, msg_ids = mail.search(None, 'SINCE "01-Apr-2026"')
    all_ids = msg_ids[0].split()

    log(f"reply_checker: Scanning {len(all_ids)} recent emails in inbox...")

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

            if sender_email in prof_rows:
                row = prof_rows[sender_email]
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
