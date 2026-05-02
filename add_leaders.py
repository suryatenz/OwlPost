"""
add_leaders.py
Scrapes and adds deans / associate deans / department chairs / directors
for every college currently represented in professors.xlsx.

Colleges covered:
  - Cullen College of Engineering
  - NSM (Natural Sciences & Mathematics)
  - C.T. Bauer College of Business
  - College of Education
  - Technology Division
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Alignment, Border, Side
import re
import os
import time
import random
import urllib3

urllib3.disable_warnings()

EXCEL = os.path.join(os.path.dirname(os.path.abspath(__file__)), "professors.xlsx")

SESSION = requests.Session()
SESSION.verify = False
SESSION.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"})

# Known leaders (hardcoded from prior research)
KNOWN_LEADERS = [
    {"name": "Pradeep Sharma",  "email": "psharma@uh.edu",          "dept": "Cullen College of Engineering"},
    {"name": "Megan Vernon",    "email": "mavernon@central.uh.edu", "dept": "Cullen College of Engineering"},
]

# Static HTML pages (requests)
STATIC_PAGES = [
    ("Cullen College of Engineering", "https://uh.edu/engineering/about/people/administration/"),
    ("Cullen College of Engineering", "https://uh.edu/engineering/about/people/"),
    ("NSM",  "https://uh.edu/nsm/about/people/administration/"),
    ("NSM",  "https://uh.edu/nsm/about/people/"),
    ("Bauer - Accountancy",         "https://bauer.uh.edu/departments/accy/people/"),
    ("Bauer - Decision & Info Sci", "https://bauer.uh.edu/departments/disc/directory/"),
    ("Bauer - Finance",             "https://bauer.uh.edu/departments/finance/faculty/"),
    ("Bauer - Management",          "https://bauer.uh.edu/departments/management-leadership/directory/faculty.php"),
    ("Bauer - Marketing",           "https://bauer.uh.edu/departments/marketing-entrepreneurship/faculty/"),
    ("Bauer College of Business",   "https://bauer.uh.edu/about/leadership/"),
    ("Bauer College of Business",   "https://bauer.uh.edu/about/people/"),
    ("College of Education",        "https://uh.edu/education/about/people/"),
    ("College of Education",        "https://uh.edu/education/about/administration/"),
    ("Technology Division",         "https://uh.edu/technology/about/people/"),
    ("Technology Division",         "https://uh.edu/technology/about/administration/"),
    ("Physics",      "https://uh.edu/nsm/physics/people/faculty/"),
    ("Biochemistry", "https://uh.edu/nsm/biology-biochemistry/people/faculty/"),
    ("Earth & Atmospheric Sciences", "https://uh.edu/nsm/earth-atmospheric/people/faculty/"),
]

# JS-rendered pages (Selenium)
JS_PAGES = [
    ("College of Education",  "https://uh.edu/education/about/people/"),
    ("Technology Division",   "https://uh.edu/technology/about/people/"),
    ("Petroleum Engineering", "https://www.pete.uh.edu/faculty/"),
    ("Cullen - ECE",          "https://www.ece.uh.edu/about/leadership/"),
    ("Cullen - ME",           "https://www.me.uh.edu/about/"),
    ("Cullen - IE",           "https://www.ie.uh.edu/about/"),
]


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


def clean_name(raw):
    raw = re.sub(r"^Dr\.\s*", "", raw.strip())
    if "," in raw:
        raw = raw.split(",")[0].strip()
    raw = re.sub(r"\s+(Fellow|Member|Senior Member|Life Fellow|NAE|AIMBE|IEEE|Ph\.?D\.?).*", "", raw, flags=re.I)
    raw = re.sub(r"\s{2,}", " ", raw)
    return raw.strip()


def extract_emails_from_html(html, dept):
    soup = BeautifulSoup(html, "html.parser")
    faculty = []
    seen = set()

    for a in soup.find_all("a", href=re.compile(r"mailto:", re.I)):
        email = a["href"].replace("mailto:", "").strip().lower()
        email = email.split("?")[0]
        if "@uh.edu" not in email and "@central.uh.edu" not in email:
            continue
        if email in seen:
            continue
        seen.add(email)

        name = ""
        fc = a.find_parent("span", class_="field-content")
        if fc:
            name = clean_name(fc.get_text(" ", strip=True))
        else:
            parent = a.find_parent(["div", "li", "tr", "article", "section", "td"])
            if parent:
                for tag in ["h2", "h3", "h4", "strong", "b"]:
                    heading = parent.find(tag)
                    if heading and len(heading.get_text(strip=True)) > 3:
                        name = clean_name(heading.get_text(strip=True))
                        break
                if not name:
                    name = clean_name(parent.get_text(" ", strip=True)[:80])

        faculty.append({"name": name, "email": email, "dept": dept})

    if not faculty:
        for email in set(re.findall(r"[a-zA-Z0-9_.+-]+@(?:uh\.edu|central\.uh\.edu)", html)):
            if email.lower() not in seen:
                faculty.append({"name": "", "email": email.lower(), "dept": dept})

    return faculty


def scrape_static(dept, url):
    try:
        r = SESSION.get(url, timeout=15)
        if r.status_code == 200:
            results = extract_emails_from_html(r.text, dept)
            print(f"    {len(results)} emails found")
            return results
        else:
            print(f"    HTTP {r.status_code}")
    except Exception as e:
        print(f"    error: {e}")
    return []


def scrape_js(driver, dept, url):
    try:
        driver.get(url)
        time.sleep(4)
        results = extract_emails_from_html(driver.page_source, dept)
        print(f"    {len(results)} emails found")
        return results
    except Exception as e:
        print(f"    error: {e}")
    return []


def load_workbook_data():
    wb = openpyxl.load_workbook(EXCEL)
    ws = wb.active
    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
    col = {h: i for i, h in enumerate(headers)}
    existing = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[col["Email"]]:
            existing.add(str(row[col["Email"]]).strip().lower())
    return wb, ws, headers, col, existing


def add_row(ws, headers, col, border, name, email, dept):
    values = {
        "Professor Name": name if name else email.split("@")[0],
        "Email":          email,
        "Department":     dept,
        "AcadGroup":      "",
        "Rank":           "",
        "Sent":           "",
        "Stop":           "",
    }
    r = ws.max_row + 1
    for i, key in enumerate(headers, 1):
        cell = ws.cell(row=r, column=i, value=values.get(key, ""))
        cell.alignment = Alignment(horizontal="left", vertical="center")
        cell.border = border


SKIP_KEYWORDS = ["admin", "office", "dept", "grad", "help", "info", "web",
                 "support", "access", "parking", "uit", "nsmoffice",
                 "webmaster", "contact", "enrollment"]


def main():
    wb, ws, headers, col, existing = load_workbook_data()

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    added = 0
    skipped = 0

    def process(entry_list):
        nonlocal added, skipped
        for f in entry_list:
            email = f["email"].strip().lower()
            if email in existing:
                skipped += 1
                continue
            if "@uh.edu" not in email and "@central.uh.edu" not in email:
                skipped += 1
                continue
            if any(x in email for x in SKIP_KEYWORDS):
                skipped += 1
                continue
            add_row(ws, headers, col, border, f["name"], email, f["dept"])
            existing.add(email)
            print(f"    + {f['name']} <{email}>")
            added += 1

    # Known leaders
    print("\n-- Known leaders --")
    process(KNOWN_LEADERS)

    # Static pages
    print("\n-- Static pages --")
    seen_urls = set()
    for dept, url in STATIC_PAGES:
        if url in seen_urls:
            continue
        seen_urls.add(url)
        print(f"\n  {dept}")
        print(f"  {url}")
        faculty = scrape_static(dept, url)
        process(faculty)
        time.sleep(random.uniform(0.8, 1.5))

    # JS-rendered pages
    print("\n-- JS-rendered pages (Selenium) --")
    driver = make_driver()
    try:
        js_seen = set()
        for dept, url in JS_PAGES:
            if url in js_seen:
                continue
            js_seen.add(url)
            print(f"\n  {dept}")
            print(f"  {url}")
            faculty = scrape_js(driver, dept, url)
            process(faculty)
            time.sleep(random.uniform(1, 2))
    finally:
        driver.quit()

    wb.save(EXCEL)
    print(f"\n{'='*55}")
    print(f"  Leaders/directors added : {added}")
    print(f"  Already present/skipped : {skipped}")
    print(f"  Saved to                : {EXCEL}")
    print(f"{'='*55}\n")


if __name__ == "__main__":
    main()
