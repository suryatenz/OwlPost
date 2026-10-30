"""
Run this ONCE (as Administrator) to register a Windows Task Scheduler job
that runs email_sender.py every day at the time set in config.json.

Right-click this script → "Run as administrator" or run from an elevated terminal.
"""

import subprocess
import sys
import os
import json

SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
SENDER      = os.path.join(SCRIPT_DIR, "email_sender.py")
CONFIG      = os.path.join(SCRIPT_DIR, "config.json")
TASK_NAME   = "ProfessorEmailSender"

def main():
    with open(CONFIG, encoding="utf-8") as f:
        config = json.load(f)

    send_time   = config.get("send_time", "09:00")
    python_exe  = r"C:\Users\spraj\AppData\Local\Programs\Python\Python311\python.exe"

    print(f"Python   : {python_exe}")
    print(f"Script   : {SENDER}")
    print(f"Schedule : daily at {send_time}")
    print()

    # Remove old task if it exists (ignore errors)
    subprocess.run(
        ["schtasks", "/delete", "/tn", TASK_NAME, "/f"],
        capture_output=True
    )

    cmd = [
        "schtasks", "/create",
        "/tn",  TASK_NAME,
        "/tr",  f'"{python_exe}" "{SENDER}"',
        "/sc",  "DAILY",
        "/st",  send_time,
        "/rl",  "HIGHEST",          # run with elevated rights
        "/f",                        # force overwrite
    ]

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode == 0:
        print(f"SUCCESS  : Task '{TASK_NAME}' scheduled daily at {send_time}.")
        print("           Outlook must be logged in on this machine for emails to send.")
        print()
        print("To change the send time: edit 'send_time' in config.json then re-run this script.")
        print("To remove the task:      schtasks /delete /tn ProfessorEmailSender /f")
    else:
        print(f"FAILED: {result.stderr.strip()}")
        print()
        print("Try running this script as Administrator:")
        print("  Right-click setup_scheduler.py → Run as administrator")


if __name__ == "__main__":
    main()
