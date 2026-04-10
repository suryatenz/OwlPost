# OwlPost

An automated tool that sends personalized RA/TA assistantship inquiry emails to university professors every day. You give it an Excel sheet of professors, it emails them 30 to 40 per day via your Gmail, marks who was contacted, automatically detects replies and skips those professors, and sends you a daily report. All on autopilot.

## What It Does

OwlPost handles the entire outreach process for you. Once it is set up, you never have to touch it again.

It reads your professor list from an Excel file, sends 30 to 40 personalized emails every morning at 9 AM with your resume attached, and marks each professor as Sent in the spreadsheet right after emailing them. When a professor replies, it automatically detects the reply and marks them as Stop so they are never emailed again. It even handles professors who reply from department alias emails that look completely different from what you have in the spreadsheet. Once it has emailed everyone on your list, it resets and starts the next cycle automatically.

Every morning after sending, you get a report email in your Gmail showing exactly who was contacted that day.

If your PC was off at 9 AM, no emails are lost. The moment you turn your PC on and log in, it waits 2 minutes for your internet to connect and then runs immediately. It also has built-in protection to make sure it never sends twice in the same day.

## Requirements

You need a Windows 10 or 11 PC, Python installed via Anaconda, a Gmail account, and your resume saved as a PDF.

## Step by Step Setup

### Step 1. Install Python (Anaconda)

Go to [anaconda.com/download](https://www.anaconda.com/download) and download the Windows installer. Run it and click Next through everything with the default settings. When it asks about adding Anaconda to PATH, make sure that box is checked.

Once it is done, open Command Prompt (search "cmd" in the Start menu) and type:
```
python --version
```
You should see something like Python 3.11.x which means it worked.

### Step 2. Download This Project

Click the green Code button at the top of this GitHub page and select Download ZIP. Extract it anywhere on your computer, for example:
```
C:\Users\YourName\professor-emailer\
```

### Step 3. Install the Required Library

Open Command Prompt and run:
```
pip install openpyxl
```
That is the only library needed. Wait for it to finish.

### Step 4. Set Up Your Config File

Find `config.example.json` in the project folder. Make a copy of it and rename the copy to `config.json`. Open it with Notepad (right-click and choose Open with Notepad) and fill in your details:

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

Replace `your@gmail.com` with your Gmail address. For the app password, follow Step 5 below. Edit the template body with your own name and details.

### Step 5. Get a Gmail App Password

Google does not let scripts use your regular Gmail password, so you need to create a free App Password. It takes about 2 minutes.

Go to [myaccount.google.com](https://myaccount.google.com) and sign in. Click Security in the left sidebar. Under "How you sign in to Google", make sure 2-Step Verification is turned ON. If it is off, click it and follow the steps to enable it since this is required before you can create an App Password.

Once 2-Step Verification is on, go back to Security and search for "App passwords" in the search bar at the top of the page. Give it any name like ProfessorEmailer and click Create. Google will show you a 16-character password like `abcd efgh ijkl mnop`. Copy it, remove the spaces, and paste it into `config.json`:

```json
"sender_app_password": "abcdefghijklmnop"
```

Keep this password private and never share your `config.json` file with anyone.

### Step 6. Add Your Resume

Copy your resume PDF into the project folder and name it `resume.pdf`. If your file has a different name, update the resume_path line in `config.json` to match:

```json
"resume_path": "My_Resume_2024.pdf"
```

### Step 7. Prepare Your Professor List

Run this once to generate a blank Excel file with the correct columns:
```
python create_excel_template.py
```

This creates `professors.xlsx` in the folder. Open it and fill in your professors:

| Column | What to enter |
|---|---|
| Professor Name | Full name like Jane Smith |
| Email | Their email address |
| Department | Department name or code |
| Sent | Leave blank, the script fills this in automatically |
| Stop | Leave blank, the script or you can fill this when someone replies |

### Step 8. (Optional) Clean a Raw Faculty List

If your university gave you a full faculty export, you can clean it automatically instead of entering professors one by one.

Open `clean_faculty.py` in Notepad and change the INPUT line at the top to point to your file:
```python
INPUT = r"C:\Users\YourName\Downloads\faculty_list.xlsx"
```
Then run it:
```
python clean_faculty.py
```

It will automatically remove teaching assistants, retirees, duplicate emails, and rows with no email address, then save a clean `professors.xlsx` ready to use.

### Step 9. Schedule Automatic Daily Sending

The recommended way is to use the included `task.xml` file. First open it in Notepad and replace `REGAL\spraj` with your own PC name and Windows username. To find these, open Command Prompt and type `whoami`, it will show `PCNAME\username`.

Then run:
```
schtasks /create /tn ProfessorEmailSender /xml task.xml
```

To verify it worked, search for Task Scheduler in the Start menu and open it. Click Task Scheduler Library in the left panel and you should see ProfessorEmailSender in the list with Status: Ready.

You can also use the Python script instead if you prefer:
```
python setup_scheduler.py
```

### Step 10. Test Before It Runs Automatically

Send today's batch right now to confirm everything is working:
```
python email_sender.py
```

Check your Gmail inbox and you should receive a daily report email listing every professor contacted. Open `professors.xlsx` and confirm the Sent column is filled in.

## How the Excel Columns Work

| Column | Who fills it | What it means |
|---|---|---|
| Professor Name | You | Full name of the professor |
| Email | You | Their email address |
| Department | You | Their department |
| Sent | Script (automatic) | Set to Yes after email is sent, do not edit this |
| Stop | You or the script | Put Yes here when a professor replies, or let reply_checker.py do it automatically |

## How the Daily Cycle Works

Every day at 9 AM the script checks for any new professor replies first, then picks 30 to 40 professors who have not been emailed yet and sends each one a personalized email with your resume. After sending, it marks them Sent in the spreadsheet and emails you a report.

When every professor on the list has been contacted, the script automatically clears the Sent column and starts the cycle over again. Anyone marked Stop is permanently skipped every cycle.

### How Reply Detection Works

Before sending each daily batch, OwlPost automatically scans your Gmail inbox for replies from professors and marks them Stop so they are never emailed again.

It uses 3 matching methods so no reply ever slips through:

| Method | Example |
|---|---|
| Exact email match | mkidd@uh.edu matches mkidd@uh.edu |
| Username match | mkidd@central.uh.edu matches mkidd@uh.edu even though the domain is different |
| Name match | "Capuano, Regina M" matched by name when their reply alias like geos2j@central.uh.edu looks completely different |

To manually scan your inbox for replies at any time:
```
python reply_checker.py
```

To scan your entire inbox history including older emails:
```
python reply_checker.py --all
```

### What Happens if Your PC is Off at 9 AM

No emails are lost. If your PC was sleeping at 9 AM, Windows wakes it and runs the task as normal. If your PC was fully off, the moment you turn it on and log in, the task waits 2 minutes for your internet to connect and then runs immediately.

The script checks a file called `last_run.txt` before doing anything. If it already sent emails today it skips and waits until tomorrow so you will never get duplicate sends. The counter resets automatically at midnight because the script compares today's date. After midnight it is a new day and the next run goes through normally.

## Useful Commands

```bash
# Run manually right now (also checks for replies before sending)
python email_sender.py

# Check for replies only without sending any emails
python reply_checker.py

# Scan your entire inbox history for replies
python reply_checker.py --all

# Check the full task status including Next Run Time and Last Run Time
schtasks /query /tn ProfessorEmailSender /fo LIST

# Quick status check
schtasks /query /tn ProfessorEmailSender /fo LIST | findstr "Status Next Last"

# See the last 20 lines of the email log
powershell -Command "Get-Content email_log.txt -Tail 20"

# Check if the script already ran today
type last_run.txt

# Pause automatic sending
schtasks /change /tn ProfessorEmailSender /disable

# Resume automatic sending
schtasks /change /tn ProfessorEmailSender /enable

# Remove the scheduled task completely
schtasks /delete /tn ProfessorEmailSender /f
```

### What the Status Values Mean

| Status | What it means |
|---|---|
| Ready | Task is active and waiting for the next scheduled time |
| Running | Emails are being sent right now |
| Disabled | Task is paused and will not run until you enable it |
| Next Run Time | When it will fire next, usually tomorrow at 9 AM |
| Last Run Time | When it last ran successfully |

## Troubleshooting

**Gmail authentication failed**
Your app password is wrong or expired. Go to Google Account, Security, App passwords and generate a new one.

**Excel file not found**
Make sure `professors.xlsx` is in the same folder as `email_sender.py` and the name matches excel_path in `config.json`.

**Resume not found**
Make sure your PDF is in the project folder and the filename exactly matches resume_path in `config.json`.

**Script ran but no emails were sent**
All professors may already be marked Sent. Add more rows or clear the Sent column to reset manually.

**A professor replied but was not marked Stop automatically**
Run `python reply_checker.py --all` to scan your full inbox. This handles professors who reply from department aliases or completely different email addresses using name matching.

**Task Scheduler ran but nothing happened**
Open `last_run.txt` and check if it contains today's date. If it does, the script already ran earlier and skipped on purpose to prevent duplicates. Check `email_log.txt` for the full history of what happened.

**PC was off at 9 AM and emails did not send when it turned on**
The task needs a LogonTrigger to fire on startup. Re-run the setup by updating your username in `task.xml` and running `schtasks /create /tn ProfessorEmailSender /xml task.xml` again.

**"Already ran today. Skipping."**
This is normal. It means the script already sent emails today. It will run again tomorrow automatically. To force a re-run today, delete `last_run.txt` and run `python email_sender.py` manually.

## Project Files

| File | Purpose |
|---|---|
| email_sender.py | Main script that checks for replies, sends emails, writes the log, and sends the daily report |
| reply_checker.py | Scans your Gmail inbox for professor replies and marks Stop = Yes automatically |
| clean_faculty.py | Cleans a raw university faculty export into a usable professor list |
| create_excel_template.py | Generates a blank professors.xlsx with the correct columns |
| setup_scheduler.py | Registers the Windows Task Scheduler job for automatic daily sending |
| task.xml | Task Scheduler config with 9 AM daily trigger and a logon trigger with 2 minute delay |
| config.example.json | Safe template to copy to config.json and fill in your own details |
| requirements.txt | Python dependencies, install with pip install -r requirements.txt |
| last_run.txt | Auto-generated file that stores today's date to prevent duplicate sends |
| email_log.txt | Auto-generated log with the full history of every email sent and every reply detected |

## Security Notes

`config.json` is listed in `.gitignore` and will never be uploaded to GitHub since it contains your password. `professors.xlsx` is also excluded because it contains personal email addresses. Never share your `config.json` or App Password with anyone. You can revoke your Gmail App Password at any time from Google Account, Security, App passwords.

## License

MIT License, free to use, modify, and share.
