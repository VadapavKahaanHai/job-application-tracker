#!/usr/bin/env python3
"""
Werkstudent and internship data-role refresh for Werkstudent_Job_Tracker.xlsx.

Priority order in the sheet:
  1 Requested German cities
  2 Verify location (search result didn't show a city)
  3 Other cities
Within each group: High fit first, then newest posting first.

What a run does
  1. Searches and requests page content through Firecrawl for LinkedIn,
     Hiring Cafe, XING and StepStone, limited to postings less than 3 days old.
  2. Keeps working-student and internship data roles in the requested cities,
     excluding known old jobs.
     Queries and skill matches are fixed from the resume; no resume is read.
  3. Adds only NEW jobs (deduplicated by link), then re-sorts the whole Jobs sheet.
     Your Status / Date Applied / Notes / Fit edits are preserved.

Setup
  pip install requests openpyxl
  export FIRECRAWL_API_KEY="fc-..."
  python job_refresh.py
"""
import os
import re
import sys
from datetime import date, datetime, timedelta
from urllib.parse import unquote, urlsplit, urlunsplit

import requests
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, Side

API_KEY = os.environ.get("FIRECRAWL_API_KEY", "")
TRACKER_PATH = os.environ.get("TRACKER_PATH", "Werkstudent_Job_Tracker.xlsx")
SEARCH_URL = "https://api.firecrawl.dev/v2/search"
MAX_AGE_DAYS = 3
CITIES = {
    "düsseldorf": "Düsseldorf", "duesseldorf": "Düsseldorf", "dusseldorf": "Düsseldorf",
    "köln": "Köln", "koeln": "Köln", "koln": "Köln", "cologne": "Köln",
    "hannover": "Hannover", "hanover": "Hannover", "hamburg": "Hamburg",
    "aachen": "Aachen", "münchen": "München", "muenchen": "München",
    "munich": "München", "frankfurt": "Frankfurt am Main",
}
TARGET_CITIES = tuple(dict.fromkeys(CITIES.values()))
SITES = (
    "site:linkedin.com/jobs/view",
    "(site:hiring.cafe/viewjob OR site:hiringcafe.com/viewjob)",
    "site:xing.com/jobs",
    "site:stepstone.de/stellenangebote",
)
# Search by role; resume skills only affect fit ratings, not search eligibility.
ROLE_QUERIES = (
    '("Data Analyst" OR "Data Analysis" OR "Data Analytics" OR Datenanalyse '
    'OR "Business Analytics" OR "Business Analyst" OR "Revenue Analytics" '
    'OR "Marketing Analytics" OR Datenanalytik)',
    '("Business Intelligence" OR "Power BI" OR Reporting OR "BI Analyst" '
    'OR "BI Developer" OR Dashboard OR Datenvisualisierung OR "Reporting Analyst")',
    '("Data Engineering" OR "Data Engineer" OR Datenengineering '
    'OR "Analytics Engineering" OR "Data Platform" OR "Data Warehouse" '
    'OR Datenmanagement OR Datenintegration OR Datenqualität)',
)
QUERIES = [
    (f'{site} (Werkstudent OR Werkstudentin OR "Working Student" OR Praktikum '
     f'OR Praktikant OR Praktikantin OR Intern OR Internship) {role} "{city}"',
     f"{city}, Germany")
    for city in TARGET_CITIES for site in SITES for role in ROLE_QUERIES
]
STUDENT = re.compile(
    r"\b(werkstudent\w*|working\s+student|student\s+(?:employee|assistant)|"
    r"praktikum|pflichtpraktikum|praktikant\w*|intern|internship)\b"
)
NOT_RELEVANT_TITLE = re.compile(
    r"\b(senior|principal|architect|director|head of|lead|vp|vice president)\b"
)
MUST_MATCH = re.compile(
    r"\b(data\s+(?:analyst\w*|analys\w*|analytics|engineer\w*)|datenanalys\w*|"
    r"datenengineering|business intelligence|bi|power\s*bi|reporting|"
    r"business\s+(?:analytics|analyst\w*)|(?:revenue|marketing)\s+analytics|"
    r"datenanalytik|dashboard\w*|datenvisualisierung|analytics\s+engineer\w*|"
    r"data\s+(?:platform\w*|warehous\w*)|datenmanagement|datenintegration|datenqualität)\b"
)
RESUME_SKILLS = (
    r"python", r"sql", r"r", r"pyspark", r"pandas", r"postgresql", r"duckdb",
    r"etl|elt", r"data pipelines?|datenpipelines?", r"data warehous\w*", r"aws",
    r"power\s*bi", r"excel", r"kpi\w*", r"data validation|datenvalidierung",
)
GERMAN_ABOVE_B1 = re.compile(
    r"(?:deutsch|german)[^\n.!?]{0,50}\b(?:b2|c1|c2|fluent|native)\b|"
    r"\b(?:b2|c1|c2|fluent|native)[^\n.!?]{0,50}(?:deutsch|german)|"
    r"(?:fließend\w*|fliessend\w*|verhandlungssicher\w*|sehr gut\w*)\s+deutsch"
)

# ---------- sheet layout (shared with the tracker file) ----------
HEADERS = ["#", "Priority", "Company", "Role", "Location", "Experience", "Source",
           "Approx. Posted", "Age (days)", "Fit", "Status", "Date Applied", "Link",
           "Notes", "Date Found"]
(COL_NUM, COL_PRI, COL_COMP, COL_ROLE, COL_LOC, COL_EXP, COL_SRC, COL_POSTED, COL_AGE,
 COL_FIT, COL_STATUS, COL_APPLIED, COL_LINK, COL_NOTE, COL_FOUND) = range(1, 16)
LAST_COL = COL_FOUND
FONT = "Arial"
THIN = Side(style="thin", color="BFBFBF")
FIT_RANK = {"High": 0, "Medium": 1, "Low": 2}


# ---------- helpers ----------
def norm(text):
    return re.sub(r"[-_/]+", " ", unquote(text or "").lower())


def clean_url(url):
    parts = urlsplit(url)
    return urlunsplit((parts.scheme, parts.netloc.lower(), parts.path.rstrip("/"), "", ""))


def to_date(v):
    if isinstance(v, datetime):
        return v.date()
    return v if isinstance(v, date) else None


def source_of(url):
    host = urlsplit(url).hostname or ""
    for domain, source in (("linkedin.com", "LinkedIn"), ("hiring.cafe", "Hiring Cafe"),
                           ("hiringcafe.com", "Hiring Cafe"), ("xing.com", "XING"),
                           ("stepstone.de", "StepStone")):
        if host == domain or host.endswith("." + domain):
            return source
    return "Other"


def is_job_url(url):
    if urlsplit(url).scheme not in ("http", "https"):
        return False
    patterns = {
        "LinkedIn": r"^/jobs/view/[^/]+$",
        "Hiring Cafe": r"^/viewjob/[^/]+$",
        "XING": r"^/jobs/[^/]+-\d+$",
        "StepStone": r"^/stellenangebote--.+--\d+[^/]*\.html$",
    }
    return bool(re.search(patterns.get(source_of(url), r"(?!)"), urlsplit(url).path, re.I))


def posted_date(title, desc, url, today):
    """Best-effort posting date; None if unknown."""
    text = f"{title} {desc}".lower()
    m = re.search(r"(\d+)\s*(hour|hr|minute|min|day|week|month|year)s?\s+ago", text)
    if m:
        n, unit = int(m.group(1)), m.group(2)
        if unit in ("hour", "hr", "minute", "min"):
            return today - timedelta(days=n // (24 if unit in ("hour", "hr") else 1440))
        return today - timedelta(days=n * {"day": 1, "week": 7, "month": 30, "year": 365}[unit])
    m = re.search(r"vor\s+(\d+)\s*(minute\w*|stunde\w*|tag\w*|woche\w*|monat\w*|jahr\w*)", text)
    if m:
        n, unit = int(m.group(1)), m.group(2)
        days = n * next((factor for prefix, factor in
                         (("tag", 1), ("woche", 7), ("monat", 30), ("jahr", 365),
                          ("stunde", 1 / 24), ("minute", 1 / 1440))
                         if unit.startswith(prefix)), 0)
        return today - timedelta(days=int(days))
    if re.search(r"\b(gestern|yesterday)\b", text):
        return today - timedelta(days=1)
    if re.search(r"\b(heute|today|just posted)\b", text):
        return today
    return None


def parse_experience(title, desc, url):
    """Returns (lo, hi) years, (lo, None) for 'N+', or None."""
    raw = f"{title} {desc} {url.replace('-', ' ')}".lower()
    m = re.search(r"\b(\d{1,2})\s*(?:to|bis|-|–)\s*(\d{1,2})\s*\+?\s*(?:yrs|years|yr|jahre[n]?)\b", raw)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.search(r"\b(\d{1,2})\s*\+\s*(?:yrs|years|yr|jahre[n]?)\b", raw)
    if m:
        return int(m.group(1)), None
    return None


def exp_label(exp):
    if not exp:
        return "Not stated"
    lo, hi = exp
    return f"{lo}-{hi} yrs" if hi else f"{lo}+ yrs"


def classify_location(title, url, desc):
    """Require an explicit target city; a search location bias is not evidence."""
    for text in (norm(f"{title} {url}"), norm(desc)):
        text = re.sub(r"\bfrankfurt\s+(?:(?:an\s+der|am)\s+)?oder\b", "", text)
        found = sorted({v for k, v in CITIES.items() if re.search(rf"\b{k}\b", text)})
        if found:
            return " / ".join(found)
    return "Verify"


def tier(location):
    """1 = requested cities, 2 = unknown, 3 = other city."""
    loc = (location or "").lower()
    if any(re.search(rf"\b{k}\b", loc) for k in CITIES):
        return 1
    if loc in ("", "verify"):
        return 2
    return 3


def guess_company_role(title):
    """Very rough; the full title still ends up in Role. Fix company by hand if needed."""
    t = re.sub(r"\s*[-|]\s*LinkedIn( Jobs)?\s*$", "", title).strip()
    m = re.match(r"(.+?)\s+hiring\s+(.+?)\s+in\s+", t, re.I)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    m = re.match(r"(.+?)\s+at\s+(.+?)(?:\s+[-—–]\s+.*)?$", t, re.I)
    if m:
        return m.group(2).strip(), m.group(1).strip()
    return "See link", t


def fit_of(text):
    t = text.lower()
    if GERMAN_ABOVE_B1.search(t):
        return "Low"
    matches = sum(bool(re.search(rf"\b(?:{skill})\b", t)) for skill in RESUME_SKILLS)
    return "High" if matches >= 2 else "Medium"


# ---------- search ----------
def run_search(query, location):
    today = date.today()
    since = today - timedelta(days=MAX_AGE_DAYS - 1)
    resp = requests.post(
        SEARCH_URL,
        headers={"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"},
        json={"query": query, "limit": 20, "sources": ["web"],
              "tbs": f"cdr:1,cd_min:{since:%m/%d/%Y},cd_max:{today:%m/%d/%Y}",
              "location": location, "country": "DE",
              "scrapeOptions": {"formats": [{"type": "markdown"}]}},
        timeout=90,
    )
    resp.raise_for_status()
    payload = resp.json()
    if payload.get("success") is False:
        raise RuntimeError("Firecrawl reported an unsuccessful search")
    data = payload.get("data", {})
    return data.get("web", []) if isinstance(data, dict) else data


def process_results(results, seen_links, today):
    """Filter raw search results into new job records. Pure function (easy to test)."""
    out = []
    for r in results:
        url = clean_url(r.get("url", ""))
        title = r.get("title", "") or ""
        desc = r.get("description", "") or r.get("snippet", "") or ""
        content = r.get("markdown", "") or ""
        blob = norm(f"{title} {desc} {content} {url}")

        if not url or url in seen_links:
            continue
        if not is_job_url(url):
            continue  # category / company pages
        if NOT_RELEVANT_TITLE.search(norm(title)) or not STUDENT.search(norm(title)):
            continue
        if not MUST_MATCH.search(norm(title)):
            continue

        # shortcut: location is text-based; use structured job metadata if false matches appear.
        location = classify_location(title, url, f"{desc}\n{content}")
        if location == "Verify":
            continue

        posted = posted_date(title, desc, url, today)
        if posted is None or not 0 <= (today - posted).days < MAX_AGE_DAYS:
            continue
        exp = parse_experience(title, desc, url)
        notes = ["Auto-added; verify company, location and requirements"]
        if not content:
            notes.append("page content unavailable; matched search snippet only")
        if GERMAN_ABOVE_B1.search(blob):
            notes.append("check German requirement against B1")
        company, role = guess_company_role(title)
        out.append(dict(
            company=company, role=role, location=location,
            exp=exp_label(exp), source=source_of(url), posted=posted, fit=fit_of(blob),
            status="Not applied", applied=None, link=url, found=today,
            note="; ".join(notes),
        ))
        seen_links.add(url)
    return out


# ---------- Excel ----------
def read_records(ws):
    recs = []
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, COL_COMP).value in (None, ""):
            continue
        link_cell = ws.cell(r, COL_LINK)
        recs.append(dict(
            company=ws.cell(r, COL_COMP).value, role=ws.cell(r, COL_ROLE).value,
            location=ws.cell(r, COL_LOC).value or "Verify",
            exp=ws.cell(r, COL_EXP).value, source=ws.cell(r, COL_SRC).value,
            posted=to_date(ws.cell(r, COL_POSTED).value), fit=ws.cell(r, COL_FIT).value,
            status=ws.cell(r, COL_STATUS).value or "Not applied",
            applied=to_date(ws.cell(r, COL_APPLIED).value),
            link=clean_url(link_cell.hyperlink.target) if link_cell.hyperlink else
                 clean_url(link_cell.value) if str(link_cell.value).startswith(("https://", "http://")) else "",
            note=ws.cell(r, COL_NOTE).value, found=to_date(ws.cell(r, COL_FOUND).value),
        ))
    return recs


def sort_key(rec):
    posted = rec["posted"] or date.min
    return (tier(rec["location"]), FIT_RANK.get(rec["fit"], 1), -posted.toordinal())


def style_row(ws, r):
    border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
    wrap_cols = (COL_COMP, COL_ROLE, COL_LOC, COL_NOTE)
    center_cols = (COL_NUM, COL_PRI, COL_EXP, COL_SRC, COL_POSTED, COL_AGE, COL_FIT,
                   COL_STATUS, COL_APPLIED, COL_LINK, COL_FOUND)
    for c in range(1, LAST_COL + 1):
        cell = ws.cell(r, c)
        cell.font = Font(name=FONT, size=10)
        cell.border = border
        cell.alignment = Alignment(
            vertical="top", wrap_text=c in wrap_cols,
            horizontal="center" if c in center_cols else None)
    for c in (COL_POSTED, COL_APPLIED, COL_FOUND):
        ws.cell(r, c).number_format = "dd-mmm-yy"
    ws.cell(r, COL_AGE).number_format = "0"
    link = ws.cell(r, COL_LINK)
    link.font = Font(name=FONT, size=10, color="0563C1", underline="single")


def write_records(ws, recs):
    recs = sorted(recs, key=sort_key)
    for r in range(2, max(ws.max_row, len(recs) + 1) + 1):
        for c in range(1, LAST_COL + 1):
            cell = ws.cell(r, c)
            cell.hyperlink = None
            cell.value = None
    for i, rec in enumerate(recs, start=2):
        style_row(ws, i)
        city_checks = ",".join(f'ISNUMBER(SEARCH("{city}",E{i}))' for city in CITIES)
        vals = {
            COL_NUM: i - 1,
            COL_PRI: (f'=IF(OR({city_checks}),"1 Target city",'
                      f'IF(OR(E{i}="",E{i}="Verify"),"2 Verify location","3 Other city"))'),
            COL_COMP: rec["company"], COL_ROLE: rec["role"], COL_LOC: rec["location"],
            COL_EXP: rec["exp"] or "Not stated", COL_SRC: rec["source"],
            COL_POSTED: rec["posted"], COL_AGE: f'=IF(H{i}="","",TODAY()-H{i})',
            COL_FIT: rec["fit"], COL_STATUS: rec["status"], COL_APPLIED: rec["applied"],
            COL_NOTE: rec["note"], COL_FOUND: rec["found"],
        }
        for c, v in vals.items():
            ws.cell(i, c).value = v
            if isinstance(v, str) and c not in (COL_PRI, COL_AGE):
                ws.cell(i, c).data_type = "s"  # Job text must never become an Excel formula.
        if rec["link"]:
            link = ws.cell(i, COL_LINK)
            link.hyperlink = rec["link"]
            link.value = "Open"
    ws.auto_filter.ref = f"A1:{ws.cell(1, LAST_COL).column_letter}{len(recs) + 1}"
    return len(recs)


def main():
    if not API_KEY:
        sys.exit("Set FIRECRAWL_API_KEY first.")
    today = date.today()
    wb = load_workbook(TRACKER_PATH)
    ws = wb["Jobs"]
    if ws.cell(1, COL_PRI).value != "Priority":
        sys.exit("Tracker layout is outdated - use Werkstudent_Job_Tracker.xlsx.")
    recs = read_records(ws)
    seen = {r["link"] for r in recs if r["link"]}
    new = []
    successful = 0
    for query, location in QUERIES:
        try:
            results = run_search(query, location)
            added = process_results(results, seen, today)
            new += added
            successful += 1
            print(f"[search] {len(results)} results, {len(added)} new matches: {query}", flush=True)
        except Exception as e:  # keep going if one query fails
            print(f"[warn] query failed: {query!r}: {e}")
    if not successful:
        sys.exit("All searches failed; tracker was not changed.")
    write_records(ws, recs + new)
    if "Summary" in wb.sheetnames:
        summary = wb["Summary"]
        summary["B2"] = "Werkstudent / Internship: Data Analysis / Business Intelligence / Data Engineering"
        summary["B3"] = today
        summary["B4"] = f"Less than {MAX_AGE_DAYS} days old; unknown posting dates excluded"
        summary["B5"] = "Target cities first; then fit and newest posting"
        summary["A8"] = "1 Target city"
        summary["B8"] = '=COUNTIF(Jobs!B2:B1000,"1 Target city")'
        summary["A18"] = "Target cities"
        summary["B18"] = ", ".join(TARGET_CITIES)
        summary["B19"] = "Fixed resume-based queries. Verify dates, language and job availability."
    wb.save(TRACKER_PATH)  # close the file in Excel before running locally
    print(f"{datetime.now():%Y-%m-%d %H:%M} - added {len(new)} new jobs "
          f"({successful}/{len(QUERIES)} searches succeeded)")


if __name__ == "__main__":
    main()
