"""
ai_reply_processor.py
AI-powered reply classifier for OwlPost professor outreach system.
Uses Claude to read each reply, classify it, write the Replied note,
color-code the spreadsheet, and draft personalized follow-ups for warm leads.

Setup:
  1. Get an Anthropic API key from https://console.anthropic.com/settings/keys
  2. Add to config.json:  "anthropic_api_key": "sk-ant-..."
  3. (Already done) pip install anthropic

Run:
  python ai_reply_processor.py              # process unclassified replies since Apr 2026
  python ai_reply_processor.py --all        # scan all inbox emails
  python ai_reply_processor.py --backfill   # classify all Stop=Yes rows with empty Replied
  python ai_reply_processor.py --dry-run    # show what would be classified, no changes
"""

import imaplib
import email
import email.header
import openpyxl
from openpyxl.styles import PatternFill
import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

try:
    import anthropic
except ImportError:
    print("ERROR: anthropic package not installed.  Run: pip install anthropic")
    sys.exit(1)

BASE_DIR   = Path(__file__).parent
CONFIG_FILE = BASE_DIR / "config.json"
LOG_FILE   = BASE_DIR / "email_log.txt"
DRAFTS_DIR = BASE_DIR / "drafts"

# ── cell background colors ────────────────────────────────────────────────────
FILL = {
    "POSITIVE":   PatternFill("solid", fgColor="C6EFCE"),  # light green
    "NEGATIVE":   PatternFill("solid", fgColor="FFC7CE"),  # light red
    "REFERRAL":   PatternFill("solid", fgColor="DDEBF7"),  # light blue
    "AUTO_REPLY": PatternFill("solid", fgColor="EDEDED"),  # light gray
    "BOUNCED":    PatternFill("solid", fgColor="FFEB9C"),  # light yellow
}


def log(msg: str):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line.encode(sys.stdout.encoding or "utf-8", errors="replace").decode(sys.stdout.encoding or "utf-8", errors="replace"))
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")


def load_config() -> dict:
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def normalize_name(name: str) -> str:
    return re.sub(r"[^a-z]", "", name.lower())


def decode_header_value(val: str) -> str:
    if not val:
        return ""
    parts = email.header.decode_header(val)
    out = []
    for text, charset in parts:
        if isinstance(text, bytes):
            out.append(text.decode(charset or "utf-8", errors="replace"))
        else:
            out.append(str(text))
    return " ".join(out)


def get_plain_body(msg) -> str:
    """Extract plain-text body; fall back to stripped HTML."""
    parts = []
    if msg.is_multipart():
        for part in msg.walk():
            ct  = part.get_content_type()
            cd  = str(part.get("Content-Disposition", ""))
            if ct == "text/plain" and "attachment" not in cd:
                payload = part.get_payload(decode=True)
                if payload:
                    cs = part.get_content_charset() or "utf-8"
                    parts.append(payload.decode(cs, errors="replace"))
    else:
        if msg.get_content_type() == "text/plain":
            payload = msg.get_payload(decode=True)
            if payload:
                cs = msg.get_content_charset() or "utf-8"
                parts.append(payload.decode(cs, errors="replace"))

    body = "\n".join(parts).strip()

    # fallback: strip HTML
    if not body and msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/html":
                payload = part.get_payload(decode=True)
                if payload:
                    cs = part.get_content_charset() or "utf-8"
                    html = payload.decode(cs, errors="replace")
                    body = re.sub(r"<[^>]+>", " ", html)
                    body = re.sub(r"\s+", " ", body).strip()
                    break

    return body


def classify_with_claude(client: anthropic.Anthropic,
                          prof_name: str,
                          prof_email: str,
                          subject: str,
                          body: str) -> dict:
    """Single Claude call: classify reply + optionally draft a reply."""

    body_trunc = body[:3500] + ("\n[...truncated...]" if len(body) > 3500 else "")

    system_prompt = (
        "You are an AI assistant helping Surya Prajyesh Eatha "
        "(Engineering Data Science graduate student at University of Houston) "
        "who sent cold outreach emails to professors inquiring about "
        "RA/TA/IA assistantship positions for Fall 2026 / Spring 2027.\n\n"
        "Classify each professor reply and return ONLY a JSON object "
        "(no markdown, no explanation) with these exact fields:\n"
        "{\n"
        '  "category": "POSITIVE" | "NEGATIVE" | "REFERRAL" | "AUTO_REPLY" | "BOUNCED",\n'
        '  "replied_note": "<spreadsheet note, max 60 chars>",\n'
        '  "action_needed": "<what Surya should do next, or null>",\n'
        '  "draft_reply": "<full email body for Surya to send, or null>",\n'
        '  "referral_name": "<name if referred to someone, or null>",\n'
        '  "referral_email": "<email if referred to someone, or null>"\n'
        "}\n\n"
        "CATEGORIES:\n"
        "  POSITIVE   — Interest shown: wants to meet, invited to apply, asked for CV, open to volunteering, etc.\n"
        "  NEGATIVE   — No positions, wrong field, retiring, not taking students, out of funding, etc.\n"
        "  REFERRAL   — Directed Surya to contact another person or department.\n"
        "  AUTO_REPLY — Out of office or fully automated vacation/OOO message.\n"
        "  BOUNCED    — Delivery failure or bounce notification (usually from mailer-daemon).\n\n"
        "replied_note examples (keep under 60 chars):\n"
        '  "Yes - Interested, asked to send CV"\n'
        '  "Yes - Invited to apply via Google Form"\n'
        '  "Yes - Wants to meet and discuss"\n'
        '  "No - No openings currently"\n'
        '  "No - Retiring / leaving UH"\n'
        '  "No - Wrong department / no match"\n'
        '  "Ref - Referred to Dr. X in Y dept"\n'
        '  "Auto - Out of office until [date]"\n'
        '  "Bounce - Invalid address"\n\n'
        "For POSITIVE: draft_reply = warm professional follow-up confirming interest "
        "and asking about next steps. Sign: Surya Prajyesh Eatha | Engineering Data Science | "
        "University of Houston | seatha@cougarnet.uh.edu\n"
        "For REFERRAL: draft_reply = thank professor and clarify the referral contact "
        "(ask for their email if not provided).\n"
        "For NEGATIVE / AUTO_REPLY / BOUNCED: draft_reply = null.\n"
        "Be professional, concise, and genuine."
    )

    user_msg = (
        f"Professor: {prof_name} ({prof_email})\n"
        f"Subject: {subject}\n"
        "---\n"
        f"{body_trunc}\n"
        "---\n"
        "Classify this reply and return the JSON."
    )

    response = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=1024,
        system=system_prompt,
        messages=[{"role": "user", "content": user_msg}],
    )

    raw = response.content[0].text.strip()

    # strip accidental code fences
    if "```json" in raw:
        raw = raw.split("```json")[1].split("```")[0].strip()
    elif "```" in raw:
        raw = raw.split("```")[1].split("```")[0].strip()

    return json.loads(raw)


def save_draft(prof_name: str, prof_email: str, category: str,
               subject: str, original_body: str, result: dict) -> None:
    if not result.get("draft_reply"):
        return
    DRAFTS_DIR.mkdir(exist_ok=True)
    safe = re.sub(r"[^a-zA-Z0-9_]", "_", prof_name)[:28]
    fname = DRAFTS_DIR / f"{category}_{safe}_{datetime.now().strftime('%Y%m%d_%H%M')}.txt"
    content = (
        f"=== DRAFT REPLY ===\n"
        f"To      : {prof_name} <{prof_email}>\n"
        f"Subject : Re: {subject}\n"
        f"Category: {category}\n"
        f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n"
        f"\nACTION NEEDED: {result.get('action_needed') or 'None'}\n"
    )
    if result.get("referral_name"):
        content += f"REFERRAL: {result['referral_name']} <{result.get('referral_email') or 'unknown'}>\n"
    content += (
        f"\n{'─'*56}\n"
        f"ORIGINAL EMAIL:\n{'─'*56}\n"
        f"{original_body[:1200]}{'...' if len(original_body) > 1200 else ''}\n"
        f"\n{'─'*56}\n"
        f"DRAFT:\n{'─'*56}\n"
        f"{result['draft_reply']}\n"
    )
    fname.write_text(content, encoding="utf-8")
    log(f"    Draft saved → drafts/{fname.name}")


def color_row(ws, row_idx: int, num_cols: int, category: str) -> None:
    fill = FILL.get(category)
    if fill:
        for c in range(1, num_cols + 1):
            ws.cell(row=row_idx, column=c).fill = fill


# ── main ──────────────────────────────────────────────────────────────────────

def process_replies(backfill: bool = False, scan_all: bool = False, dry_run: bool = False):
    config = load_config()

    api_key = config.get("anthropic_api_key") or os.environ.get("ANTHROPIC_API_KEY", "")
    if not dry_run and (not api_key or api_key.startswith("YOUR_")):
        print("\n" + "=" * 60)
        print("  SETUP REQUIRED — add your Anthropic API key to config.json:")
        print('  "anthropic_api_key": "sk-ant-..."')
        print("  Get one free at: https://console.anthropic.com/settings/keys")
        print("=" * 60 + "\n")
        sys.exit(1)

    client = anthropic.Anthropic(api_key=api_key) if not dry_run else None

    excel_path = BASE_DIR / config["excel_path"]
    if not excel_path.exists():
        log("ERROR: professors.xlsx not found")
        return

    wb = openpyxl.load_workbook(str(excel_path))
    ws = wb.active
    headers  = [str(c.value).strip() if c.value else "" for c in ws[1]]
    col      = {h: i for i, h in enumerate(headers)}
    num_cols = len(headers)

    # ── build professor lookup maps ───────────────────────────────────────────
    prof_by_email    = {}   # "mkidd@uh.edu" -> (row_tuple, row_idx)
    prof_by_username = {}   # "mkidd"        -> (row_tuple, row_idx)
    prof_by_name     = {}   # "marcusdkidd"  -> (row_tuple, row_idx)

    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        ev = row[col["Email"]].value
        nv = row[col["Professor Name"]].value
        if ev:
            full = str(ev).strip().lower()
            prof_by_email[full]                 = (row, i)
            prof_by_username[full.split("@")[0]] = (row, i)
        if nv:
            prof_by_name[normalize_name(str(nv))] = (row, i)

    log(f"ai_processor: {len(prof_by_email)} professors loaded from spreadsheet")

    # ── connect to Gmail IMAP ─────────────────────────────────────────────────
    gmail_user = config["sender_email"]
    gmail_pass = config["sender_app_password"]
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(gmail_user, gmail_pass)
        mail.select("inbox")
    except Exception as e:
        log(f"ERROR: Gmail IMAP connect failed — {e}")
        return

    if scan_all:
        _, msg_ids = mail.search(None, "ALL")
    else:
        _, msg_ids = mail.search(None, 'SINCE "01-Apr-2026"')

    all_ids = msg_ids[0].split()
    log(f"ai_processor: Scanning {len(all_ids)} emails (newest-first)...")

    processed = skipped = errors = 0
    seen_rows = set()   # avoid double-processing one professor

    for msg_id in reversed(all_ids):    # newest first → most recent reply wins
        try:
            _, msg_data = mail.fetch(msg_id, "(RFC822)")
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)

            from_header = msg.get("From", "")
            subject     = decode_header_value(msg.get("Subject", ""))

            # extract sender address
            if "<" in from_header and ">" in from_header:
                sender_email = from_header.split("<")[1].split(">")[0].strip().lower()
            else:
                sender_email = from_header.strip().lower()
            sender_user = sender_email.split("@")[0] if "@" in sender_email else sender_email

            # display name → normalized for name-match
            sender_name_norm = ""
            if "<" in from_header:
                display = from_header.split("<")[0].strip().strip('"')
                if "," in display:        # "Last, First" → "First Last"
                    parts   = display.split(",", 1)
                    display = parts[1].strip() + " " + parts[0].strip()
                sender_name_norm = normalize_name(display)

            match = (prof_by_email.get(sender_email)
                     or prof_by_username.get(sender_user)
                     or (prof_by_name.get(sender_name_norm) if sender_name_norm else None))

            if not match:
                continue

            row, row_idx = match

            # skip if already processed in this run (keep only newest email per prof)
            if row_idx in seen_rows:
                continue
            seen_rows.add(row_idx)

            prof_name    = str(row[col["Professor Name"]].value or "").strip()
            replied_cell = row[col["Replied"]]
            stop_cell    = row[col["Stop"]]
            replied_val  = str(replied_cell.value or "").strip()

            if backfill:
                stop_val = str(stop_cell.value or "").strip().lower()
                if stop_val != "yes":
                    continue
                if replied_val:
                    skipped += 1
                    continue
            else:
                # normal mode: only process rows without a Replied note yet
                if replied_val:
                    skipped += 1
                    continue

            body = get_plain_body(msg)
            if not body or len(body) < 10:
                log(f"  SKIP {prof_name}: empty body")
                continue

            log(f"  >> {prof_name} <{sender_email}>")

            if dry_run:
                preview = body[:120].replace("\n", " ").strip()
                log(f"    [DRY-RUN] Subject: {subject[:60]}")
                log(f"    [DRY-RUN] Preview: {preview}...")
                processed += 1
                continue

            try:
                result = classify_with_claude(client, prof_name, sender_email, subject, body)
            except Exception as e:
                log(f"    ERROR calling Claude: {e}")
                errors += 1
                continue

            category     = result.get("category", "NEGATIVE")
            replied_note = (result.get("replied_note") or f"Classified: {category}")[:60]

            replied_cell.value = replied_note
            stop_cell.value    = "Yes"
            color_row(ws, row_idx, num_cols, category)
            wb.save(str(excel_path))

            save_draft(prof_name, sender_email, category, subject, body, result)

            log(f"    [{category}] {replied_note}")
            if result.get("action_needed"):
                log(f"    ACTION: {result['action_needed']}")
            if result.get("referral_name"):
                log(f"    REFERRAL: {result['referral_name']} <{result.get('referral_email') or ''}>")

            processed += 1
            time.sleep(0.4)     # brief pause between API calls

        except Exception as e:
            log(f"  ERROR processing msg {msg_id}: {e}")
            errors += 1
            continue

    mail.logout()

    log(f"\nai_processor: Done.")
    log(f"  Classified  : {processed}")
    log(f"  Skipped     : {skipped}  (already had a Replied note)")
    log(f"  Errors      : {errors}")
    if processed and not dry_run:
        log(f"  Drafts      : {DRAFTS_DIR}/")


if __name__ == "__main__":
    args = set(sys.argv[1:])
    backfill = "--backfill" in args
    scan_all = "--all"      in args
    dry_run  = "--dry-run"  in args

    if dry_run:
        log("ai_processor: DRY-RUN — no spreadsheet changes will be made")
    elif backfill:
        log("ai_processor: BACKFILL mode — classifying all Stop=Yes rows with empty Replied")
    elif scan_all:
        log("ai_processor: Scanning ALL inbox emails")
    else:
        log("ai_processor: Processing replies since April 1 2026")

    process_replies(backfill=backfill, scan_all=scan_all, dry_run=dry_run)
