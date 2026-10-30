"""
Professor Email Sender
Sends personalized emails via Gmail SMTP or Outlook/UH SMTP.
Reads professors from Excel, marks sent rows, loops when done.
"""

import openpyxl
import smtplib
import ssl
import json
import os
import random
import time
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

CONFIG_FILE   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
LOG_FILE      = os.path.join(os.path.dirname(os.path.abspath(__file__)), "email_log.txt")
LAST_RUN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "last_run.txt")

# SMTP profiles — pick one via "email_provider" in config.json
SMTP_PROFILES = {
    "gmail": {
        "server": "smtp.gmail.com",
        "port":   587,
    },
    "outlook": {
        "server": "smtp.office365.com",
        "port":   587,
    },
}


# ── helpers ────────────────────────────────────────────────────────────────

def already_ran_today():
    """Return True if the script already sent emails today."""
    if not os.path.exists(LAST_RUN_FILE):
        return False
    with open(LAST_RUN_FILE, "r") as f:
        return f.read().strip() == datetime.now().strftime("%Y-%m-%d")

def mark_ran_today():
    """Save today's date so we don't run twice in one day."""
    with open(LAST_RUN_FILE, "w") as f:
        f.write(datetime.now().strftime("%Y-%m-%d"))

def load_config():
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)

def is_yes(value):
    if value is None:
        return False
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ("yes", "true", "1", "x", "✓", "sent")

def log(msg):
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(line)
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

def get_column_map(ws):
    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
    return {h: i for i, h in enumerate(headers)}

def load_pending_rows(ws, col):
    pending = []
    for row in ws.iter_rows(min_row=2):
        sent  = row[col["Sent"]].value
        stop  = row[col["Stop"]].value
        email = row[col["Email"]].value
        if email and not is_yes(stop) and not is_yes(sent):
            pending.append(row)
    return pending

def reset_cycle(ws, col):
    log("All professors emailed! Resetting for next cycle (skipping stopped rows).")
    for row in ws.iter_rows(min_row=2):
        if not is_yes(row[col["Stop"]].value):
            row[col["Sent"]].value = None

def build_message(sender, to_email, subject, body, resume_path):
    msg = MIMEMultipart()
    msg["From"]    = sender
    msg["To"]      = to_email
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    if resume_path and os.path.exists(resume_path):
        with open(resume_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={os.path.basename(resume_path)}"
        )
        msg.attach(part)

    return msg


# ── core ───────────────────────────────────────────────────────────────────

def send_emails():
    # ── guard: don't run twice in one day ──
    if already_ran_today():
        log("Already ran today. Skipping. Will run again tomorrow.")
        return

    # ── auto check for replies first ──
    from reply_checker import check_replies
    check_replies()

    config = load_config()

    # ── validate config ──
    provider = config.get("email_provider", "").lower()
    if provider not in SMTP_PROFILES:
        log(f"ERROR: 'email_provider' must be 'gmail' or 'outlook' in config.json (got: '{provider}')")
        return

    sender_email   = config.get("sender_email", "").strip()
    sender_password = config.get("sender_app_password", "").strip()
    if not sender_email or not sender_password:
        log("ERROR: 'sender_email' and 'sender_app_password' must be set in config.json")
        return

    excel_path = os.path.join(os.path.dirname(CONFIG_FILE), config["excel_path"])
    if not os.path.exists(excel_path):
        log(f"ERROR: Excel file not found at {excel_path}")
        return

    resume_path = os.path.join(os.path.dirname(CONFIG_FILE), config.get("resume_path", ""))
    if not os.path.exists(resume_path):
        log(f"WARNING: Resume not found at {resume_path} — emails will send without attachment.")
        resume_path = None

    # ── load spreadsheet ──
    try:
        wb  = openpyxl.load_workbook(excel_path)
    except Exception as e:
        log(f"ERROR: Cannot open Excel ({e}). Close Excel and rerun.")
        return
    ws  = wb.active
    col = get_column_map(ws)

    for required in ("Professor Name", "Email", "Department", "Sent", "Stop"):
        if required not in col:
            log(f"ERROR: Column '{required}' not found in Excel.")
            return

    pending = load_pending_rows(ws, col)
    if not pending:
        reset_cycle(ws, col)
        wb.save(excel_path)
        wb  = openpyxl.load_workbook(excel_path)
        ws  = wb.active
        col = get_column_map(ws)
        pending = load_pending_rows(ws, col)

    if not pending:
        log("No emails to send (all rows may be stopped).")
        return

    daily_count = random.randint(config["daily_min"], config["daily_max"])
    batch = pending[:daily_count]
    log(f"Provider : {provider.upper()} ({sender_email})")
    log(f"Batch    : {len(batch)} emails today")

    # ── connect to SMTP ──
    profile = SMTP_PROFILES[provider]
    try:
        context = ssl.create_default_context()
        server  = smtplib.SMTP(profile["server"], profile["port"])
        server.ehlo()
        server.starttls(context=context)
        server.login(sender_email, sender_password)
        log("SMTP connection established.")
    except Exception as e:
        log(f"ERROR: Could not connect to {provider} SMTP — {e}")
        return

    # ── send loop ──
    sent_count  = 0
    failed_count = 0
    sent_log    = []   # track for daily report

    for row in batch:
        prof_name  = str(row[col["Professor Name"]].value or "Professor").strip()
        prof_email = str(row[col["Email"]].value or "").strip()
        dept       = str(row[col["Department"]].value or "").strip()

        name_parts = prof_name.split()
        prof_last  = name_parts[-1] if name_parts else prof_name

        if not prof_email:
            log(f"  SKIP {prof_name}: no email address")
            continue

        subject = (config["subject"]
                   .replace("{name}", prof_name)
                   .replace("{last_name}", prof_last)
                   .replace("{department}", dept))
        body    = (config["template"]
                   .replace("{name}", prof_name)
                   .replace("{last_name}", prof_last)
                   .replace("{department}", dept))

        try:
            msg = build_message(sender_email, prof_email, subject, body, resume_path)
            server.sendmail(sender_email, prof_email, msg.as_string())

            row[col["Sent"]].value = "Yes"
            wb.save(excel_path)
            sent_count += 1
            sent_log.append((sent_count, prof_name, prof_email, dept))
            log(f"  SENT -> {prof_name} <{prof_email}>")

            if sent_count < len(batch):
                delay = random.randint(config.get("delay_min", 10),
                                       config.get("delay_max", 30))
                time.sleep(delay)

        except Exception as e:
            failed_count += 1
            log(f"  FAILED -> {prof_name} <{prof_email}>: {e}")

    server.quit()
    mark_ran_today()
    log(f"Done. {sent_count}/{len(batch)} emails sent today.")

    # ── daily report ──────────────────────────────────────────────────────
    send_daily_report(sender_email, sender_password, sent_log, failed_count, config)

    # ── inbox cleanup (bounces + promos + junk) ───────────────────────────
    try:
        from cleanup_inbox import clean_inbox
        clean_inbox(config)
    except Exception as e:
        log(f"WARNING: Inbox cleanup failed — {e}")


def send_daily_report(sender_email, sender_password, sent_log, failed_count, config):
    today      = datetime.now().strftime("%B %d, %Y")
    total_sent = len(sent_log)

    rows = "\n".join(
        f"  {i:>3}.  {name:<35} {email:<40} {dept}"
        for i, name, email, dept in sent_log
    )

    body = f"""Daily Email Report — {today}
{"=" * 60}

  Emails sent today : {total_sent}
  Failed            : {failed_count}
  Sent from         : {sender_email}

{"=" * 60}
{"  #    Professor Name                      Email                                    Department" }
{"-" * 110}
{rows if rows else "  No emails were sent."}
{"=" * 60}

This is an automated report from your Professor Email Sender.
"""

    msg = MIMEMultipart()
    msg["From"]    = sender_email
    msg["To"]      = sender_email
    msg["Subject"] = f"Email Report: {total_sent} professors contacted on {today}"
    msg.attach(MIMEText(body, "plain"))

    try:
        profile = SMTP_PROFILES[config.get("email_provider", "gmail")]
        ctx     = ssl.create_default_context()
        s       = smtplib.SMTP(profile["server"], profile["port"])
        s.ehlo()
        s.starttls(context=ctx)
        s.login(sender_email, sender_password)
        s.sendmail(sender_email, sender_email, msg.as_string())
        s.quit()
        log(f"Daily report sent to {sender_email}")
    except Exception as e:
        log(f"WARNING: Could not send daily report — {e}")


if __name__ == "__main__":
    send_emails()
