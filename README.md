# Professor Email Sender

An automated tool that sends personalized RA/TA assistantship inquiry emails to university professors every day. You give it an Excel sheet of professors, it emails them 30–40 per day via your Gmail, marks who was contacted, skips anyone who replied, and sends you a daily report — all on autopilot.

---

## What It Does

- Reads professors from an Excel spreadsheet (Name, Email, Department)
- Sends 30–40 personalized emails per day automatically at 9 AM
- Attaches your resume to every email
- Marks each professor as `Sent` in the spreadsheet right after emailing
- Skips any professor you mark as `Stop` (when they reply to you)
- When the full list is done, automatically resets and starts the next cycle
- Sends you a daily report email listing every professor contacted that day
- Runs via Windows Task Scheduler — no manual work after setup

---

## Requirements

- Windows 10 or 11
- Python (Anaconda) — free download
- A Gmail account
- Your resume as a PDF file

---

## Step-by-Step Setup

### Step 1 — Install Python (Anaconda)

1. Go to [anaconda.com/download](https://www.anaconda.com/download)
2. Download the Windows installer
3. Run the installer — click **Next** through everything, keep all defaults
4. When asked "Add Anaconda to PATH" — check that box
5. Finish installation

To verify it worked, open **Command Prompt** (search "cmd" in Start menu) and type:
```
python --version
```
You should see something like `Python 3.11.x`

---

### Step 2 — Download This Project

Click the green **Code** button on this GitHub page → **Download ZIP**

Extract the ZIP to anywhere on your computer, for example:
```
C:\Users\YourName\professor-emailer\
```

---

### Step 3 — Install the Required Library

Open **Command Prompt** and run:
```
pip install openpyxl
```
Wait for it to finish. That is the only library needed.

---

### Step 4 — Set Up Your Config File

1. In the project folder, find `config.example.json`
2. Make a copy of it and rename the copy to `config.json`
3. Open `config.json` with Notepad (right-click → Open with → Notepad)
4. Fill in your details:

```json
{
    "email_provider": "gmail",
    "sender_email": "your@gmail.com",
    "sender_app_password": "paste-your-app-password-here",
    "resume_path": "resume.pdf",
    "subject": "Inquiry Regarding RA/TA Assistantship Opportunities",
    "template": "Dear Professor {last_name},\n\n...",
    "daily_min": 30,
    "daily_max": 40,
    "send_time": "09:00"
}
```

Replace `your@gmail.com` with your Gmail address.
For the `sender_app_password` — see Step 5 below.
Edit the `template` body with your own name and details.

---

### Step 5 — Get a Gmail App Password

Google does not allow scripts to use your normal Gmail password. You need a free **App Password** — a separate 16-character password just for this tool.

1. Go to [myaccount.google.com](https://myaccount.google.com) and sign in
2. Click **Security** in the left sidebar
3. Under "How you sign in to Google", make sure **2-Step Verification is ON**
   - If it is off, click it and follow the steps to turn it on (required)
4. After 2-Step Verification is on, go back to Security
5. In the search bar at the top of the page, type **App passwords** and click it
6. Give it any name (e.g. `ProfessorEmailer`) and click **Create**
7. Google shows you a 16-character password like `abcd efgh ijkl mnop`
8. Copy it and paste it into `config.json` as `sender_app_password` (remove the spaces):
```json
"sender_app_password": "abcdefghijklmnop"
```

> Keep this password private. Never share your `config.json` file with anyone.

---

### Step 6 — Add Your Resume

Copy your resume PDF into the project folder and name it `resume.pdf`.

If your resume has a different name, update the `resume_path` line in `config.json`:
```json
"resume_path": "My_Resume_2024.pdf"
```

---

### Step 7 — Prepare Your Professor List

Run this once to generate a blank Excel file with the correct columns:
```
python create_excel_template.py
```

This creates `professors.xlsx` in the folder. Open it and fill in your professors:

| Column | What to enter |
|---|---|
| `Professor Name` | Full name (e.g. `Jane Smith`) |
| `Email` | Professor's email address |
| `Department` | Department name or code |
| `Sent` | Leave blank — the script fills this automatically |
| `Stop` | Leave blank — you fill this when someone replies |

Save the file when done.

---

### Step 8 — (Optional) Clean a Raw Faculty List

If your university gave you a full faculty export (Excel or CSV), you can clean it automatically instead of entering professors manually.

1. Open `clean_faculty.py` in Notepad
2. Change the `INPUT` line at the top to point to your file:
```python
INPUT = r"C:\Users\YourName\Downloads\faculty_list.xlsx"
```
3. Run it:
```
python clean_faculty.py
```

It will automatically remove teaching assistants, retirees, duplicate emails, and rows with no email, then save a clean `professors.xlsx` ready to use.

---

### Step 9 — Schedule Automatic Daily Sending

Run this once to set up Windows Task Scheduler:
```
python setup_scheduler.py
```

This registers a task that runs `email_sender.py` every day at the time set in `send_time` in your `config.json` (default 9:00 AM).

To verify the task was created:
1. Search **Task Scheduler** in the Start menu and open it
2. In the left panel click **Task Scheduler Library**
3. You should see **ProfessorEmailSender** in the list with Status: **Ready**

---

### Step 10 — Test Before It Runs Automatically

Send today's batch right now to make sure everything works:
```
python email_sender.py
```

Check your Gmail inbox — you should receive a daily report email listing every professor contacted. Open `professors.xlsx` and confirm the `Sent` column is filled in.

---

## How the Excel Columns Work

| Column | Who fills it | What it means |
|---|---|---|
| `Professor Name` | You | Name of the professor |
| `Email` | You | Professor's email address |
| `Department` | You | Their department |
| `Sent` | Script (automatic) | Set to `Yes` after email is sent — do not edit |
| `Stop` | You | Put `Yes` here if a professor replies — they will never be emailed again |

---

## How the Daily Cycle Works

1. Every day at 9 AM the script picks 30–40 unsent, non-stopped professors
2. Sends each a personalized email with your resume attached
3. Marks them `Sent = Yes` in the spreadsheet
4. After the batch, sends you a report email
5. When ALL professors in the sheet have been sent → automatically clears `Sent` and starts over from the top (professors with `Stop = Yes` are permanently skipped)

---

## Useful Commands

```bash
# Run manually right now
python email_sender.py

# Check if the scheduled task exists
schtasks /query /tn ProfessorEmailSender

# Pause automatic sending
schtasks /change /tn ProfessorEmailSender /disable

# Resume automatic sending
schtasks /change /tn ProfessorEmailSender /enable

# Remove the scheduled task completely
schtasks /delete /tn ProfessorEmailSender /f
```

---

## Troubleshooting

**"Gmail authentication failed"**
Your app password is wrong. Re-generate it from Google Account → Security → App passwords.

**"Excel file not found"**
Make sure `professors.xlsx` is in the same folder as `email_sender.py` and the name matches `excel_path` in `config.json`.

**"Resume not found"**
Make sure your PDF is in the project folder and the filename matches `resume_path` in `config.json`.

**Script ran but no emails sent**
All professors may already be marked `Sent`. Either add more rows or delete the contents of the `Sent` column to reset manually.

**Task Scheduler ran but nothing happened**
Make sure your PC was on and connected to the internet at the scheduled time. If it was asleep, the task will run the next time you turn the PC on.

---

## Project Files

| File | Purpose |
|---|---|
| `email_sender.py` | Main script — reads Excel, sends emails, writes log, sends daily report |
| `clean_faculty.py` | Cleans a raw university faculty export into a usable professor list |
| `create_excel_template.py` | Generates a blank `professors.xlsx` with the correct columns |
| `setup_scheduler.py` | Registers the Windows Task Scheduler job for automatic daily sending |
| `config.example.json` | Safe template — copy to `config.json` and fill in your details |
| `requirements.txt` | Python dependencies (`pip install -r requirements.txt`) |

---

## Security Notes

- `config.json` is in `.gitignore` and will never be uploaded to GitHub — it contains your password
- `professors.xlsx` is also excluded — it contains personal email addresses
- Never share your `config.json` or App Password with anyone
- Your Gmail App Password can be revoked anytime from Google Account → Security → App passwords

---

## License

MIT License — free to use, modify, and share.
