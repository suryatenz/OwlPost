"""
fix_names.py
Fixes rows in professors.xlsx where the Professor Name column is
garbled (starts with 'Office:', is an email address, or is too long).

Strategy:
  1. Re-scrape faculty pages with smarter name extraction
  2. Build email -> name lookup
  3. Update all bad-name rows
"""

import requests
from bs4 import BeautifulSoup
import openpyxl
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

# Pages to re-scrape for name lookup
PAGES = [
    # Cullen Engineering (JS) - handled via Selenium below
    # NSM static
    ("https://uh.edu/nsm/computer-science/people/faculty/",     "static"),
    ("https://uh.edu/nsm/chemistry/people/faculty/",            "static"),
    ("https://uh.edu/nsm/physics/people/faculty/",              "static"),
    ("https://uh.edu/nsm/biology-biochemistry/people/faculty/", "static"),
    ("https://uh.edu/nsm/earth-atmospheric/people/faculty/",    "static"),
    # Bauer static
    ("https://bauer.uh.edu/departments/marketing-entrepreneurship/faculty/", "static"),
    # Cullen Engineering (JS-rendered)
    ("https://www.ece.uh.edu/faculty/",  "js"),
    ("https://www.me.uh.edu/faculty/",   "js"),
    ("https://www.bme.uh.edu/faculty/",  "js"),
    ("https://www.chee.uh.edu/faculty/", "js"),
    ("https://www.cive.uh.edu/faculty/", "js"),
    ("https://www.ie.uh.edu/faculty/",   "js"),
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


def looks_like_name(text):
    """Rough check: 2-5 words, no digits, no special chars."""
    t = text.strip()
    if not t or len(t) < 4 or len(t) > 50:
        return False
    if any(c.isdigit() for c in t):
        return False
    if any(c in t for c in ["@", "/", ":", "(", ")"]):
        return False
    parts = t.split()
    return 2 <= len(parts) <= 5


def extract_name_email_pairs(html):
    """
    Returns dict of {email: name} from a faculty page HTML.
    Tries multiple extraction strategies:
      1. Drupal field-content span (name is heading inside it)
      2. Walking up from mailto link to find nearest heading
      3. UH People-style: h3/h4 directly above the email link
    """
    soup = BeautifulSoup(html, "html.parser")
    lookup = {}

    for a in soup.find_all("a", href=re.compile(r"mailto:", re.I)):
        email = a["href"].replace("mailto:", "").strip().lower().split("?")[0]
        if "@uh.edu" not in email and "@central.uh.edu" not in email:
            continue
        if email in lookup:
            continue

        name = ""

        # Strategy 1: Drupal field-content (span wrapping the whole card)
        fc = a.find_parent("span", class_="field-content")
        if fc:
            for tag in ["h3", "h4", "h2", "strong"]:
                h = fc.find(tag)
                if h:
                    candidate = clean_name(h.get_text(strip=True))
                    if looks_like_name(candidate):
                        name = candidate
                        break

        # Strategy 2: walk up to article/div/li, find nearest heading
        if not name:
            parent = a.find_parent(["article", "div", "li", "tr", "td", "section"])
            if parent:
                for tag in ["h2", "h3", "h4", "strong", "b"]:
                    h = parent.find(tag)
                    if h:
                        candidate = clean_name(h.get_text(strip=True))
                        if looks_like_name(candidate):
                            name = candidate
                            break

        # Strategy 3: previous sibling text nodes
        if not name:
            for sib in a.find_all_previous(["h2", "h3", "h4", "strong"]):
                candidate = clean_name(sib.get_text(strip=True))
                if looks_like_name(candidate):
                    name = candidate
                    break

        if name:
            lookup[email] = name

    return lookup


def build_lookup():
    lookup = {}

    js_urls = [url for url, kind in PAGES if kind == "js"]
    static_urls = [url for url, kind in PAGES if kind == "static"]

    # Static pages
    for url in static_urls:
        try:
            r = SESSION.get(url, timeout=15)
            if r.status_code == 200:
                pairs = extract_name_email_pairs(r.text)
                print(f"  {len(pairs)} names  {url}")
                lookup.update(pairs)
        except Exception as e:
            print(f"  error {url}: {e}")
        time.sleep(random.uniform(0.5, 1.0))

    # JS pages
    if js_urls:
        driver = make_driver()
        try:
            for url in js_urls:
                try:
                    driver.get(url)
                    time.sleep(4)
                    pairs = extract_name_email_pairs(driver.page_source)
                    print(f"  {len(pairs)} names  {url}")
                    lookup.update(pairs)
                except Exception as e:
                    print(f"  error {url}: {e}")
                time.sleep(random.uniform(1, 2))
        finally:
            driver.quit()

    return lookup


def is_bad_name(name):
    if not name:
        return True
    if name.startswith("Office") or name.startswith("Email:"):
        return True
    if "@" in name:
        return True
    if len(name) > 60:
        return True
    return False


def username_to_display(email):
    """Last resort: turn 'jsmith3@uh.edu' into a placeholder."""
    u = email.split("@")[0]
    # strip trailing digits
    u = re.sub(r"\d+$", "", u)
    # split camelCase or underscores
    parts = re.sub(r"([a-z])([A-Z])", r"\1 \2", u).replace("_", " ").split()
    return " ".join(p.capitalize() for p in parts) if parts else u


def main():
    print("Building name lookup from faculty pages...")
    lookup = build_lookup()
    print(f"  Total names in lookup: {len(lookup)}\n")

    wb = openpyxl.load_workbook(EXCEL)
    ws = wb.active
    headers = [str(c.value).strip() if c.value else "" for c in ws[1]]
    col = {h: i for i, h in enumerate(headers)}

    fixed = 0
    still_bad = 0

    for row in ws.iter_rows(min_row=2):
        name_cell = row[col["Professor Name"]]
        email_cell = row[col["Email"]]

        name = str(name_cell.value or "").strip()
        email = str(email_cell.value or "").strip().lower()

        if not is_bad_name(name):
            continue

        # Look up from scraped data
        new_name = lookup.get(email, "")

        if not new_name:
            # fallback: derive from username
            new_name = username_to_display(email)
            still_bad += 1

        name_cell.value = new_name
        fixed += 1

    wb.save(EXCEL)
    print(f"Fixed  : {fixed}")
    print(f"  - from scrape : {fixed - still_bad}")
    print(f"  - from email  : {still_bad} (username-derived fallback)")
    print(f"Saved  : {EXCEL}")


if __name__ == "__main__":
    main()
