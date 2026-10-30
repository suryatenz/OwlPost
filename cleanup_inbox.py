"""
cleanup_inbox.py
Cleans Gmail inbox after daily email batch.

Steps performed:
  1. Bounce detection  — finds mailer-daemon/postmaster notifications,
                         extracts the failed recipient email,
                         marks that professor Stop=Yes + Replied="Bounce" in spreadsheet,
                         then deletes the bounce notification.
  2. Promotion cleanup — moves everything in [Gmail]/Promotions to trash.
  3. Junk sender sweep — deletes emails from known non-professor domains.

Usage:
  python cleanup_inbox.py            # full cleanup
  python cleanup_inbox.py --dry-run  # preview only, nothing is deleted
"""

import imaplib
import email
import email.header
import json
import os
import re
import sys
from datetime import datetime

import openpyxl

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
LOG_FILE    = os.path.join(BASE_DIR, "email_log.txt")

DRY_RUN = "--dry-run" in sys.argv

# ── Known promotional / junk sender patterns (regex on From header) ──────────
JUNK_SENDER_PATTERNS = [
    r"transparentlabs\.com",
    r"m\.mcdonalds\.com",
    r"a\.grubhub\.com",
    r"grubhub\.com",
    r"doordash\.com",
    r"ubereats\.com",
    r"notifications\.unidays\.com",
    r"email\.hm\.com",
    r"usps\.com",
    r"mailer\.robinhood\.com",
    r"info\.robinhood\.com",
    r"noreply@robinhood\.com",
    r"hello@coinbase\.com",
    r"noreply@medium\.com",
    r"newsletter@",
    r"no-reply@marketing\.",
    r"deals@",
    r"offers@",
    r"promotions@",
    r"unsubscribe@",
    r"news@",
]
JUNK_PATTERNS_COMPILED = [re.compile(p, re.IGNORECASE) for p in JUNK_SENDER_PATTERNS]

# ── Regexes to extract failed recipient from bounce body ─────────────────────
BOUNCE_EMAIL_REGEXES = [
    re.compile(r"Final-Recipient:\s*rfc822;\s*<?([^\s>]+@[^\s>]+)>?", re.IGNORECASE),
    re.compile(r"Original-Recipient:\s*rfc822;\s*<?([^\s>]+@[^\s>]+)>?", re.IGNORECASE),
    re.compile(r"The\s+(?:email\s+)?address\s+<?([^\s>]+@[^\s>]+)>?\s+(?:could not|wasn't|was not)", re.IGNORECASE),
    re.compile(r"Your\s+message\s+wasn't\s+delivered\s+to\s+<?([^\s>]+@[^\s>]+)>?", re.IGNORECASE),
    re.compile(r"<([^\s>]+@[^\s>]+)>:\s+(?:user unknown|no such user|does not exist|invalid)", re.IGNORECASE),
    re.compile(r"failed\s+permanently.*?<([^\s>]+@[^\s>]+)>", re.IGNORECASE | re.DOTALL),
]


# ── helpers ──────────────────────────────────────────────────────────────────

def log(msg: str):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    safe = line.encode(sys.stdout.encoding or "utf-8", errors="replace").decode(
        sys.stdout.encoding or "utf-8", errors="replace"
    )
    print(safe)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_config() -> dict:
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def decode_header_value(raw: str) -> str:
    parts = email.header.decode_header(raw or "")
    decoded = []
    for part, charset in parts:
        if isinstance(part, bytes):
            decoded.append(part.decode(charset or "utf-8", errors="replace"))
        else:
            decoded.append(str(part))
    return "".join(decoded)


def get_plain_body(msg) -> str:
    body_parts = []
    if msg.is_multipart():
        for part in msg.walk():
            ct = part.get_content_type()
            cd = str(part.get("Content-Disposition", ""))
            if ct == "text/plain" and "attachment" not in cd:
                charset = part.get_content_charset() or "utf-8"
                try:
                    body_parts.append(part.get_payload(decode=True).decode(charset, errors="replace"))
                except Exception:
                    body_parts.append(str(part.get_payload()))
    else:
        charset = msg.get_content_charset() or "utf-8"
        try:
            body_parts.append(msg.get_payload(decode=True).decode(charset, errors="replace"))
        except Exception:
            body_parts.append(str(msg.get_payload()))
    return "\n".join(body_parts)


def extract_bounced_email(body: str) -> str | None:
    for pattern in BOUNCE_EMAIL_REGEXES:
        m = pattern.search(body)
        if m:
            addr = m.group(1).strip().lower().rstrip(".")
            if "@" in addr and "." in addr.split("@")[1]:
                return addr
    # Fallback: find any edu/university-looking email in the body
    for m in re.finditer(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.edu", body):
        return m.group(0).strip().lower()
    return None


def is_bounce_sender(from_header: str) -> bool:
    lower = from_header.lower()
    return any(x in lower for x in [
        "mailer-daemon",
        "postmaster@",
        "mail delivery subsystem",
        "mail delivery failure",
        "undelivered mail returned",
        "delivery status notification",
    ])


def is_junk_sender(from_header: str) -> bool:
    return any(p.search(from_header) for p in JUNK_PATTERNS_COMPILED)


def trash_message(mail: imaplib.IMAP4_SSL, msg_id: bytes, label: str = "bounce"):
    if DRY_RUN:
        return
    try:
        mail.copy(msg_id, "[Gmail]/Trash")
        mail.store(msg_id, "+FLAGS", "\\Deleted")
    except Exception as e:
        log(f"    WARNING: Could not trash {label} msg — {e}")


def mark_professor_bounce(excel_path: str, failed_email: str) -> bool:
    if DRY_RUN:
        return True
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
        col = {h: i for i, h in enumerate(headers)}

        if "Email" not in col or "Stop" not in col:
            return False

        found = False
        for row in ws.iter_rows(min_row=2):
            cell_email = str(row[col["Email"]].value or "").strip().lower()
            if cell_email == failed_email.lower():
                row[col["Stop"]].value = "Yes"
                if "Replied" in col:
                    row[col["Replied"]].value = "Bounce"
                found = True
                break

        if found:
            wb.save(excel_path)
        return found
    except Exception as e:
        log(f"    WARNING: Could not update spreadsheet for {failed_email} — {e}")
        return False


# ── core cleanup functions ───────────────────────────────────────────────────

def clean_bounces(mail: imaplib.IMAP4_SSL, excel_path: str) -> int:
    log("Scanning for bounce notifications...")
    mail.select("inbox")

    # Search for mailer-daemon and postmaster messages
    result, ids = mail.search(None, '(OR FROM "mailer-daemon" FROM "postmaster")')
    if result != "OK" or not ids[0]:
        # Try alternative search
        result, ids = mail.search(None, 'SUBJECT "Mail delivery failed"')
    if result != "OK" or not ids[0]:
        result, ids = mail.search(None, 'SUBJECT "Undelivered Mail Returned"')
    if result != "OK" or not ids[0]:
        result, ids = mail.search(None, 'SUBJECT "Delivery Status Notification"')

    all_ids = set()
    if result == "OK" and ids[0]:
        for id_str in ids[0].split():
            all_ids.add(id_str)

    # Also search for delivery failure subjects
    for subj in [
        'SUBJECT "Mail delivery failed"',
        'SUBJECT "Undelivered Mail"',
        'SUBJECT "Delivery Status"',
        'SUBJECT "Delivery Failure"',
        'SUBJECT "failure notice"',
    ]:
        r2, i2 = mail.search(None, subj)
        if r2 == "OK" and i2[0]:
            for id_str in i2[0].split():
                all_ids.add(id_str)

    if not all_ids:
        log("  No bounce notifications found.")
        return 0

    log(f"  Found {len(all_ids)} potential bounce emails.")
    cleaned = 0

    for msg_id in all_ids:
        _, msg_data = mail.fetch(msg_id, "(RFC822)")
        if not msg_data or not msg_data[0]:
            continue
        raw_msg = email.message_from_bytes(msg_data[0][1])
        from_hdr = decode_header_value(raw_msg.get("From", ""))
        subject  = decode_header_value(raw_msg.get("Subject", ""))

        if not is_bounce_sender(from_hdr) and "delivery" not in subject.lower() and "undeliver" not in subject.lower():
            continue

        body = get_plain_body(raw_msg)
        failed_addr = extract_bounced_email(body)

        if failed_addr:
            in_sheet = mark_professor_bounce(excel_path, failed_addr)
            status = "marked Stop=Yes in spreadsheet" if in_sheet else "not found in spreadsheet (may already be stopped)"
            action = "[DRY RUN] would delete" if DRY_RUN else "deleted"
            log(f"  BOUNCE >> {failed_addr} — {status} — {action}")
        else:
            action = "[DRY RUN] would delete" if DRY_RUN else "deleted"
            log(f"  BOUNCE (no address extracted) — {subject[:60]} — {action}")

        trash_message(mail, msg_id, label="bounce")
        cleaned += 1

    if not DRY_RUN:
        try:
            mail.expunge()
        except Exception:
            pass

    return cleaned


def clean_promotions(mail: imaplib.IMAP4_SSL) -> int:
    log("Scanning Gmail Promotions folder...")
    total = 0

    # Try the Gmail Promotions category folder
    for folder_name in ["[Gmail]/Promotions", "[Gmail]/Category/Promotions", "Promotions"]:
        try:
            status, _ = mail.select(folder_name)
            if status != "OK":
                continue

            _, ids = mail.search(None, "ALL")
            if not ids[0]:
                log(f"  {folder_name}: empty.")
                break

            count = len(ids[0].split())
            action = f"[DRY RUN] would trash {count}" if DRY_RUN else f"trashing {count}"
            log(f"  {folder_name}: {action} emails...")

            if not DRY_RUN:
                for msg_id in ids[0].split():
                    try:
                        mail.store(msg_id, "+FLAGS", "\\Deleted")
                    except Exception:
                        pass
                mail.expunge()

            total += count
            break
        except Exception:
            continue

    # Also search inbox for promotional patterns
    mail.select("inbox")
    for pattern in [
        'HEADER "List-Unsubscribe" ""',
        'HEADER "Precedence" "bulk"',
        'HEADER "Precedence" "list"',
    ]:
        try:
            _, ids = mail.search(None, pattern)
            if not ids[0]:
                continue
            count = len(ids[0].split())
            action = f"[DRY RUN] would trash {count}" if DRY_RUN else f"trashing {count}"
            log(f"  Bulk/list emails in inbox ({pattern}): {action}...")
            if not DRY_RUN:
                for msg_id in ids[0].split():
                    try:
                        mail.copy(msg_id, "[Gmail]/Trash")
                        mail.store(msg_id, "+FLAGS", "\\Deleted")
                    except Exception:
                        pass
                mail.expunge()
            total += count
        except Exception:
            continue

    return total


def clean_junk_senders(mail: imaplib.IMAP4_SSL) -> int:
    log("Scanning for known junk senders in inbox...")
    mail.select("inbox")
    cleaned = 0

    junk_search_terms = [
        'FROM "transparentlabs.com"',
        'FROM "m.mcdonalds.com"',
        'FROM "a.grubhub.com"',
        'FROM "grubhub.com"',
        'FROM "doordash.com"',
        'FROM "notifications.unidays.com"',
        'FROM "email.hm.com"',
        'FROM "hm.com"',
        'FROM "mailer.robinhood.com"',
        'FROM "info.robinhood.com"',
        'FROM "uber.com"',
        'FROM "ubereats.com"',
        'FROM "coinbase.com"',
    ]

    for term in junk_search_terms:
        try:
            _, ids = mail.search(None, term)
            if not ids[0]:
                continue
            count = len(ids[0].split())
            sender = term.replace('FROM "', "").rstrip('"')
            action = f"[DRY RUN] would trash {count}" if DRY_RUN else f"trashing {count}"
            log(f"  {sender}: {action} emails")
            if not DRY_RUN:
                for msg_id in ids[0].split():
                    try:
                        mail.copy(msg_id, "[Gmail]/Trash")
                        mail.store(msg_id, "+FLAGS", "\\Deleted")
                    except Exception:
                        pass
                mail.expunge()
            cleaned += count
        except Exception:
            continue

    return cleaned


# ── main entry point ─────────────────────────────────────────────────────────

def clean_inbox(config: dict | None = None):
    if config is None:
        config = load_config()

    excel_path = os.path.join(BASE_DIR, config["excel_path"])
    sender_email = config["sender_email"]
    password = config["sender_app_password"]

    mode = "DRY RUN" if DRY_RUN else "LIVE"
    log(f"=== Inbox Cleanup [{mode}] ===")

    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(sender_email, password)
        log("IMAP connected.")
    except Exception as e:
        log(f"ERROR: Could not connect to Gmail IMAP — {e}")
        return

    bounces_cleaned   = clean_bounces(mail, excel_path)
    promos_cleaned    = clean_promotions(mail)
    junk_cleaned      = clean_junk_senders(mail)

    mail.logout()

    total = bounces_cleaned + promos_cleaned + junk_cleaned
    log(f"=== Cleanup done: {bounces_cleaned} bounces, {promos_cleaned} promos, {junk_cleaned} junk ({total} total) ===")


if __name__ == "__main__":
    clean_inbox()
