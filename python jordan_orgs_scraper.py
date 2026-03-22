"""
Jordan Civil Society Organizations Scraper
Source: http://www.civilsociety-jo.net
Output: jordan_organizations.xlsx

Requirements:
    pip install requests beautifulsoup4 openpyxl

Usage:
    python jordan_orgs_scraper.py

The script scrapes every organization page (IDs 1–1500),
extracts name, governorate, category, phone, email, website,
and address, then saves everything to an Excel file.
"""

import time
import re
import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_URL = "http://www.civilsociety-jo.net/en/organization/{id}/"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    )
}

# IDs observed go up to ~25000; the script skips 404s automatically
ID_RANGE = range(1, 25001)

# Delay between requests (seconds) — be polite to the server
REQUEST_DELAY = 0.8


def scrape_org(org_id: int, session: requests.Session) -> dict | None:
    url = BASE_URL.format(id=org_id)
    try:
        resp = session.get(url, timeout=15, verify=False)
    except requests.RequestException:
        return None

    if resp.status_code != 200:
        return None

    soup = BeautifulSoup(resp.text, "html.parser")

    # If the page has no org title it's a redirect / empty page
    title_tag = soup.find("h1")
    if not title_tag:
        return None
    name = title_tag.get_text(strip=True)
    if not name or name.lower() in ("home", ""):
        return None

    # ── Breadcrumb gives us category ──────────────────────────────────────────
    breadcrumbs = soup.select("ol.breadcrumb li, .breadcrumb li, nav li")
    category = ""
    if len(breadcrumbs) >= 2:
        # last breadcrumb before the org name is the category
        category = breadcrumbs[-2].get_text(strip=True)

    # ── Right-hand contact card ────────────────────────────────────────────────
    # The contact block is a <div> that contains the org name as a link,
    # then plain text lines with phone / fax / address, and <a> tags for
    # email and website.
    contact_block = soup.find("div", class_=re.compile(r"contact|info|card", re.I))
    if not contact_block:
        # fall back: look for the section that immediately follows the h1
        contact_block = soup.find("section") or soup.find("article") or soup.body

    full_text = contact_block.get_text("\n", strip=True) if contact_block else ""

    # Website
    website = ""
    for a in (contact_block or soup).find_all("a", href=True):
        href = a["href"]
        if href.startswith("http") and "civilsociety-jo.net" not in href:
            website = href
            break

    # Email
    email = ""
    mailto = (contact_block or soup).find("a", href=re.compile(r"^mailto:", re.I))
    if mailto:
        email = mailto["href"].replace("mailto:", "").strip()
    else:
        # Try to find raw email in text
        match = re.search(r"[\w.+-]+@[\w-]+\.[a-zA-Z]{2,}", full_text)
        if match:
            email = match.group()

    # Phone — lines that look like Jordanian numbers: 06-…, 079…, +962…, 07…
    phone_pattern = re.compile(
        r"(?:\+962[\s\-]?|0)(?:6|7\d)[\s\-]?\d{3,4}[\s\-]?\d{4}"
    )
    phones = list(dict.fromkeys(phone_pattern.findall(full_text)))  # unique, ordered
    phone = " / ".join(phones[:3]) if phones else ""

    # Fax — same pattern but preceded by "fax"
    fax = ""
    fax_match = re.search(
        r"[Ff]ax[:\s]*(" + phone_pattern.pattern + r")", full_text
    )
    if fax_match:
        fax = fax_match.group(1)

    # Governorate — shown as a small badge/link near the top
    governorate = ""
    gov_link = (contact_block or soup).find(
        "a",
        href=re.compile(r"governorate|amman|irbid|zarqa|mafraq|ajloun|jerash|"
                        r"madaba|balqa|karak|tafileh|maan|aqaba", re.I)
    )
    if gov_link:
        governorate = gov_link.get_text(strip=True)
    else:
        # Try the page text: look for a governorate keyword near the title
        gov_pattern = re.compile(
            r"\b(Amman|Irbid|Zarqa|Mafraq|Ajloun|Jerash|Madaba|Balqa|"
            r"Karak|Tafileh|Ma.an|Aqaba)\b", re.I
        )
        gm = gov_pattern.search(full_text)
        if gm:
            governorate = gm.group(1).title()

    # Address — everything after the last phone number until end of block
    address = ""
    addr_match = re.search(
        r"(?:Jordan|P\.?O\.?\s*Box|Street|St\.|Amman|Irbid)[^\n]{5,80}", full_text
    )
    if addr_match:
        address = addr_match.group().strip()

    return {
        "ID": org_id,
        "Name": name,
        "Category": category,
        "Governorate": governorate,
        "Phone": phone,
        "Fax": fax,
        "Email": email,
        "Website": website,
        "Address": address,
        "Source URL": url,
    }


def build_excel(records: list[dict], output_path: str) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Organizations"

    # ── Styles ─────────────────────────────────────────────────────────────────
    header_font = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", start_color="1F4E79")  # dark blue
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    alt_fill = PatternFill("solid", start_color="EBF3FB")  # light blue stripe

    columns = [
        ("No.",          5),
        ("Organization Name", 40),
        ("Category",     25),
        ("Governorate",  14),
        ("Phone",        22),
        ("Fax",          18),
        ("Email",        32),
        ("Website",      35),
        ("Address",      40),
        ("Source URL",   45),
    ]

    # Header row
    ws.row_dimensions[1].height = 30
    for col_idx, (header, width) in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Freeze header
    ws.freeze_panes = "A2"

    # Data rows
    fields = ["ID", "Name", "Category", "Governorate",
              "Phone", "Fax", "Email", "Website", "Address", "Source URL"]

    for row_idx, rec in enumerate(records, start=2):
        fill = alt_fill if row_idx % 2 == 0 else PatternFill()
        ws.row_dimensions[row_idx].height = 18
        for col_idx, field in enumerate(fields, start=1):
            val = rec.get(field, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Arial", size=10)
            cell.alignment = left_wrap
            cell.border = border
            if fill.fill_type:
                cell.fill = fill

    # Auto-filter
    ws.auto_filter.ref = ws.dimensions

    # Summary sheet
    ws2 = wb.create_sheet("Summary")
    ws2["A1"] = "Total Organizations Scraped"
    ws2["B1"] = len(records)
    ws2["A1"].font = Font(bold=True, name="Arial")

    gov_counts: dict[str, int] = {}
    for r in records:
        g = r.get("Governorate") or "Unknown"
        gov_counts[g] = gov_counts.get(g, 0) + 1

    ws2["A3"] = "Governorate"
    ws2["B3"] = "Count"
    ws2["A3"].font = Font(bold=True, name="Arial")
    ws2["B3"].font = Font(bold=True, name="Arial")
    for i, (gov, cnt) in enumerate(sorted(gov_counts.items()), start=4):
        ws2.cell(row=i, column=1, value=gov).font = Font(name="Arial", size=10)
        ws2.cell(row=i, column=2, value=cnt).font = Font(name="Arial", size=10)
    ws2.column_dimensions["A"].width = 20
    ws2.column_dimensions["B"].width = 10

    wb.save(output_path)
    print(f"\n✅  Saved {len(records)} organizations → {output_path}")


def main():
    output_file = "jordan_organizations.xlsx"
    session = requests.Session()
    session.headers.update(HEADERS)

    records = []
    consecutive_misses = 0

    print("Starting scrape of civilsociety-jo.net …")
    print(f"Checking IDs 1 – {max(ID_RANGE)}  (delay: {REQUEST_DELAY}s per request)\n")

    for org_id in ID_RANGE:
        result = scrape_org(org_id, session)
        if result:
            records.append(result)
            consecutive_misses = 0
            gov = result["Governorate"] or "?"
            print(f"  [{org_id:4d}] ✓  {result['Name'][:55]:<55}  [{gov}]")
        else:
            consecutive_misses += 1
            print(f"  [{org_id:4d}] –  (no data)")
            # If 100 IDs in a row return nothing, we're past the end
            if consecutive_misses >= 2000:
                print(f"\nNo data found for 2000 consecutive IDs. Stopping early.")
                break

        time.sleep(REQUEST_DELAY)

    if records:
        build_excel(records, output_file)
    else:
        print("No records found. Check your internet connection or the site URL.")


if __name__ == "__main__":
    main()