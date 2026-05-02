import imaplib, email, json

config = json.load(open(r'C:\Users\spraj\professor-emailer\config.json'))
mail = imaplib.IMAP4_SSL('imap.gmail.com')
mail.login(config['sender_email'], config['sender_app_password'])

# Scan All Mail for .edu replies since Apr 1
mail.select('"[Gmail]/All Mail"')
_, msg_ids = mail.search(None, 'SINCE "01-Apr-2026"')
all_ids = msg_ids[0].split()
print(f'Emails since Apr 1 in All Mail: {len(all_ids)}')
print()

edu_senders = []
for msg_id in all_ids:
    try:
        _, msg_data = mail.fetch(msg_id, '(RFC822)')
        raw = msg_data[0][1]
        msg = email.message_from_bytes(raw)
        from_header = msg.get('From', '')
        subject = msg.get('Subject', '')
        to_header = msg.get('To', '')
        # Only show emails RECEIVED (not sent by us) from .edu
        if '.edu' in from_header.lower() and config['sender_email'] not in from_header.lower():
            edu_senders.append((from_header, subject))
    except Exception:
        continue

print(f'=== EMAILS FROM .EDU ADDRESSES ({len(edu_senders)}) ===')
for sender, subj in edu_senders:
    print(f'  FROM: {sender}')
    print(f'  SUBJ: {subj}')
    print()

mail.logout()
