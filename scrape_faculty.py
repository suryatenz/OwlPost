"""
scrape_faculty.py
Scrapes UH department faculty pages for relevant departments,
extracts professor names and emails, and adds them to professors.xlsx.

Targets departments relevant to Engineering Data Science:
CS, ECE, IE, ME, ChE, Civil, BME, Math, Chemistry.

Uses requests+BeautifulSoup for NSM pages (server-side rendered).
Uses Selenium headless Chrome for Cullen Engineering dept pages (JS-rendered).
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
import re
import os
import time
import random
from openpyxl.styles import Alignment, Border, Side
import urllib3

urllib3.disable_warnings()

EXCEL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "professors.xlsx")

# NSM pages — server-side rendered, use requests
STATIC_DEPTS = {
    "Computer Science":  "https://uh.edu/nsm/computer-science/people/faculty/",
    "Mathematics":       "https://uh.edu/nsm/math/people/faculty/",
    "Chemistry":         "https://uh.edu/nsm/chemistry/people/faculty/",
}

# Cullen Engineering dept pages — JS-rendered, require Selenium
JS_DEPTS = {
    "Electrical & Computer Eng": "https://www.ece.uh.edu/faculty/",
    "Mechanical Engineering":     "https://www.me.uh.edu/faculty/",
    "Biomedical Engineering":     "https://www.bme.uh.edu/faculty/",
    "Chemical & Biomolecular Eng":"https://www.chee.uh.edu/faculty/",
    "Civil & Environmental Eng":  "https://www.cive.uh.edu/faculty/",
    "Industrial Engineering":     "https://www.ie.uh.edu/faculty/",
}

SESSION = requests.Session()
SESSION.verify = False
SESSION.headers.update({"User-Agent": "Mozilla/5.0"})


# ── Selenium setup ────────────────────────────────────────────────────────────

def make_driver():
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from webdriver_manager.chrome import ChromeDriverManager

    opts = Options()
    opts.add_argument("--headless")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-gpu")
    opts.add_argument("--ignore-certificate-errors")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--log-level=3")
    opts.add_experimental_option("excludeSwitches", ["enable-logging"])
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)


# ── Extraction helpers ────────────────────────────────────────────────────────

def clean_name(raw):
    """Strip titles, affiliations, extra info — keep just the name."""
    # Remove "Dr." prefix
    raw = re.sub(r"^Dr\.\s*", "", raw.strip())
    # Keep only up to first comma (strip rank/title after it)
    if "," in raw:
        raw = raw.split(",")[0].strip()
    # Remove IEEE/AIMBE/NAE suffixes like "Fellow IEEE"
    raw = re.sub(r"\s+(Fellow|Member|Senior Member|Life Fellow|NAE|AIMBE|IEEE).*", "", raw, flags=re.I)
    return raw.strip()


def extract_from_html(html, dept):
    """Parse static-rendered HTML for email + name pairs."""
    soup = BeautifulSoup(html, "html.parser")
    faculty = []
    seen = set()

    for a in soup.find_all("a", href=re.compile(r"mailto:")):
        email = a["href"].replace("mailto:", "").strip().lower()
        if "@uh.edu" not in email and "@central.uh.edu" not in email:
            continue
        if email in seen:
            continue
        seen.add(email)

        name = ""
        # Try going up to field-content span (Drupal pattern)
        fc = a.find_parent("span", class_="field-content")
        if fc:
            name = clean_name(fc.get_text(" ", strip=True))
        else:
            parent = a.find_parent(["div", "li", "tr", "article", "section"])
            if parent:
                for tag in ["h2", "h3", "h4", "strong", "b"]:
                    heading = parent.find(tag)
                    if heading and len(heading.get_text(strip=True)) > 3:
                        name = clean_name(heading.get_text(strip=True))
                        break
                if not name:
                    name = clean_name(parent.get_text(" ", strip=True)[:80])

        faculty.append({"name": name, "email": email, "dept": dept})

    # Fallback if no mailto links found
    if not faculty:
        for email in set(re.findall(r"[a-zA-Z0-9_.+-]+@(?:uh\.edu|central\.uh\.edu)", html)):
            faculty.append({"name": "", "email": email.lower(), "dept": dept})

    return faculty


def scrape_static(dept, url):
    """Scrape a server-side-rendered faculty page with requests."""
    try:
        r = SESSION.get(url, timeout=15)
        if r.status_code == 200:
            return extract_from_html(r.text, dept)
    except Exception as e:
        print(f"  requests error: {e}")
    return []


def scrape_js(driver, dept, url):
    """Scrape a JS-rendered faculty page with Selenium."""
    try:
        driver.get(url)
        time.sleep(5)  # wait for JS to render
        return extract_from_html(driver.page_source, dept)
    except Exception as e:
        print(f"  selenium error: {e}")
    return []


# ── Spreadsheet helpers ───────────────────────────────────────────────────────

def load_existing(ws, col):
    existing_emails = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[col["Email"]]:
            existing_emails.add(str(row[col["Email"]]).strip().lower())
    return existing_emails


def add_professor(ws, headers, col, border, f):
    values = {
        "Professor Name": f["name"] if f["name"] else f["email"].split("@")[0],
        "Email":          f["email"],
        "Department":     f["dept"],
        "AcadGroup":      "",
        "Rank":           "",
        "Sent":           "",
        "Stop":           "",
    }
    next_row = ws.max_row + 1
    for i, key in enumerate(headers, 1):
        cell = ws.cell(row=next_row, column=i, value=values.get(key, ""))
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = border


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    wb = openpyxl.load_workbook(EXCEL)
    ws = wb.active
    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
    col = {h: i for i, h in enumerate(headers)}

    existing_emails = load_existing(ws, col)

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    SKIP_KEYWORDS = ["admin", "office", "dept", "grad", "help", "info", "web",
                     "support", "access", "parking", "uit", "nsmoffice"]

    added_total = 0
    skipped_total = 0

    def process_faculty(faculty_list):
        nonlocal added_total, skipped_total
        added = 0
        skipped = 0

        for f in faculty_list:
            email = f["email"].strip().lower()

            if email in existing_emails:
                skipped += 1
                skipped_total += 1
                continue
            if "@uh.edu" not in email and "@central.uh.edu" not in email:
                skipped += 1
                continue
            if any(x in email for x in SKIP_KEYWORDS):
                skipped += 1
                continue

            add_professor(ws, headers, col, border, f)
            existing_emails.add(email)
            added += 1
            added_total += 1

        print(f"  Added: {added} | Skipped: {skipped}")

    # ── Static pages (no browser needed) ─────────────────────────────────────
    for dept, url in STATIC_DEPTS.items():
        print(f"\nScraping (static): {dept}")
        faculty = scrape_static(dept, url)
        print(f"  Found {len(faculty)} entries on page")
        process_faculty(faculty)
        time.sleep(random.uniform(1, 2))

    # ── JS-rendered pages (Selenium) ──────────────────────────────────────────
    print("\nStarting headless Chrome for JS-rendered pages...")
    driver = make_driver()

    try:
        for dept, url in JS_DEPTS.items():
            print(f"\nScraping (JS): {dept}")
            faculty = scrape_js(driver, dept, url)
            print(f"  Found {len(faculty)} entries on page")
            process_faculty(faculty)
            time.sleep(random.uniform(1, 2))
    finally:
        driver.quit()

    # ── Save ─────────────────────────────────────────────────────────────────
    wb.save(EXCEL)
    print(f"\n{'='*50}")
    print(f"  Total new professors added : {added_total}")
    print(f"  Total skipped              : {skipped_total}")
    print(f"  Saved to                   : {EXCEL}")
    print(f"{'='*50}\n")


if __name__ == "__main__":
    main()
