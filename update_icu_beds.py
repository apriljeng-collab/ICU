#!/usr/bin/env python3
"""
Fetch ICU bed numbers from openinfo.mohw.gov.tw and update the Excel file.

Install dependencies first:
    pip install requests openpyxl beautifulsoup4 lxml

Run:
    python update_icu_beds.py

The script reads hospital codes/names from ALL sheet column D/E,
queries openinfo.mohw.gov.tw for current bed counts, and writes
new values to column K (加護病床數量) and L (燒傷加護病床數).
Results are also saved to icu_results.json so you can re-run
update_excel_from_cache() without re-scraping.
"""

import time
import re
import json
import sys
from pathlib import Path

import requests
from bs4 import BeautifulSoup
import openpyxl

EXCEL_FILE = Path("Hospital_Chair_20240731_update.xlsx")
CACHE_FILE = Path("icu_results.json")
BASE_URL = "https://openinfo.mohw.gov.tw"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
}

SESSION = requests.Session()
SESSION.headers.update(HEADERS)

# ---------------------------------------------------------------------------
# Site-structure discovery
# ---------------------------------------------------------------------------

def discover_search_endpoint() -> dict:
    """
    Visit the home page and find the form / AJAX endpoint used for
    hospital search.  Returns a dict with keys:
      method  : "GET" or "POST"
      url     : full URL to submit to
      fields  : {field_name: default_value} for hidden/pre-filled inputs
    """
    print("Discovering search endpoint …")
    resp = SESSION.get(BASE_URL + "/", timeout=20)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "lxml")

    # Look for a <form> that seems hospital-search related
    for form in soup.find_all("form"):
        action = form.get("action", "")
        text = form.get_text(strip=True)
        if any(kw in text or kw in action for kw in ("醫院", "機構", "hosp", "search")):
            method = (form.get("method") or "GET").upper()
            url = action if action.startswith("http") else BASE_URL + "/" + action.lstrip("/")
            fields = {}
            for inp in form.find_all("input"):
                name = inp.get("name")
                val = inp.get("value", "")
                if name and inp.get("type") not in ("submit", "button", "image"):
                    fields[name] = val
            return {"method": method, "url": url, "fields": fields}

    # Fallback: look for XHR hints in script tags
    for script in soup.find_all("script"):
        src = script.get("src", "")
        content = script.string or ""
        # Common patterns in AJAX calls
        for pattern in (r'"(/[^"]*search[^"]*)"', r"'(/[^']*search[^']*)'"):
            m = re.search(pattern, content, re.IGNORECASE)
            if m:
                return {"method": "POST", "url": BASE_URL + m.group(1), "fields": {}}

    # Last resort: well-known paths on this domain
    for path in ("/T1/T1D1104.htm", "/api/hospital/search", "/search"):
        url = BASE_URL + path
        try:
            r = SESSION.get(url, timeout=10)
            if r.status_code == 200:
                return {"method": "GET", "url": url, "fields": {}}
        except Exception:
            pass

    raise RuntimeError(
        "Could not find search endpoint on openinfo.mohw.gov.tw. "
        "Please inspect the site manually and update SEARCH_ENDPOINT below."
    )


# Hard-code the endpoint here once you've run the script once and found it,
# to avoid repeated discovery.
# SEARCH_ENDPOINT = {"method": "POST", "url": "https://openinfo.mohw.gov.tw/...", "fields": {}}
SEARCH_ENDPOINT: dict | None = None


# ---------------------------------------------------------------------------
# Bed-count extraction helpers
# ---------------------------------------------------------------------------

# These regex patterns match common field names in both HTML and JSON
_ICU_PATTERNS = [
    re.compile(r"加護病床[數量]*\s*[：:]\s*(\d+)"),
    re.compile(r"ICU\s*[：:]\s*(\d+)", re.IGNORECASE),
    re.compile(r'"加護病床[^"]*"\s*[：:]\s*["\']?(\d+)'),
    re.compile(r"icu[_\-]?beds?\s*[=:]\s*(\d+)", re.IGNORECASE),
]
_BURN_PATTERNS = [
    re.compile(r"燒傷加護病床[數量]*\s*[：:]\s*(\d+)"),
    re.compile(r'"燒傷[^"]*"\s*[：:]\s*["\']?(\d+)'),
]


def _extract_numbers_from_text(text: str) -> dict | None:
    icu = None
    for pat in _ICU_PATTERNS:
        m = pat.search(text)
        if m:
            icu = int(m.group(1))
            break
    if icu is None:
        return None

    burn = 0
    for pat in _BURN_PATTERNS:
        m = pat.search(text)
        if m:
            burn = int(m.group(1))
            break

    return {"icu": icu, "burn_icu": burn}


def parse_html_page(html: str) -> dict | None:
    """Extract ICU beds from an HTML response."""
    soup = BeautifulSoup(html, "lxml")

    # Strategy 1: look for table cells adjacent to 加護病床 label
    for cell in soup.find_all(["td", "th"]):
        if "加護病床" in (cell.get_text(strip=True)):
            # value should be in the next sibling cell
            next_cell = cell.find_next_sibling(["td", "th"])
            if next_cell:
                val_text = next_cell.get_text(strip=True)
                nums = re.findall(r"\d+", val_text)
                if nums:
                    icu = int(nums[0])
                    # look for burn ICU nearby
                    burn = 0
                    for row in soup.find_all("tr"):
                        rtxt = row.get_text(strip=True)
                        if "燒傷" in rtxt:
                            bns = re.findall(r"\d+", rtxt)
                            if bns:
                                burn = int(bns[0])
                    return {"icu": icu, "burn_icu": burn}

    # Strategy 2: embedded JSON
    for script in soup.find_all("script"):
        content = script.string or ""
        if "加護" in content:
            result = _extract_numbers_from_text(content)
            if result:
                return result

    # Strategy 3: full-text scan
    return _extract_numbers_from_text(soup.get_text())


def parse_json_response(data) -> dict | None:
    if isinstance(data, list):
        data = data[0] if data else {}
    if not isinstance(data, dict):
        return None

    icu_keys = ("加護病床數量", "加護病床", "ICUBeds", "icuBeds", "icu_beds", "icu")
    burn_keys = ("燒傷加護病床數", "燒傷加護病床", "burnICUBeds", "burn_icu")

    for k in icu_keys:
        if k in data:
            icu = int(data[k] or 0)
            burn = 0
            for bk in burn_keys:
                if bk in data:
                    burn = int(data[bk] or 0)
            return {"icu": icu, "burn_icu": burn}
    return None


# ---------------------------------------------------------------------------
# Per-hospital search
# ---------------------------------------------------------------------------

def search_hospital(endpoint: dict, hosp_code: str, hosp_name: str) -> dict | None:
    """Try code-based and name-based queries; return {"icu":int,"burn_icu":int} or None."""

    for query_params in [
        {"hospcode": hosp_code},
        {"HospCode": hosp_code},
        {"code": hosp_code},
        {"hospname": hosp_name},
        {"HospName": hosp_name},
        {"keyword": hosp_name},
    ]:
        fields = {**endpoint.get("fields", {}), **query_params}
        try:
            if endpoint["method"] == "POST":
                resp = SESSION.post(endpoint["url"], data=fields, timeout=15)
            else:
                resp = SESSION.get(endpoint["url"], params=fields, timeout=15)

            if resp.status_code != 200:
                continue

            # Try JSON first
            try:
                data = resp.json()
                result = parse_json_response(data)
                if result:
                    return result
            except ValueError:
                pass

            # Fall back to HTML
            result = parse_html_page(resp.text)
            if result:
                return result

        except requests.RequestException:
            continue

    return None


# ---------------------------------------------------------------------------
# Excel I/O
# ---------------------------------------------------------------------------

def read_hospitals() -> list[dict]:
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["ALL"]
    hospitals = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        code = row[3]   # col D
        name = row[4]   # col E
        if code:
            hospitals.append({"code": str(code), "name": str(name) if name else ""})
    return hospitals


def update_excel(results: dict[str, dict]):
    """Write ICU counts from *results* into ALL sheet columns K and L."""
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["ALL"]

    updated = skipped = 0
    for row in ws.iter_rows(min_row=3):
        code = str(row[3].value) if row[3].value else None
        if not code:
            continue
        if code in results:
            row[10].value = results[code]["icu"]        # col K
            row[11].value = results[code]["burn_icu"]   # col L
            updated += 1
        else:
            skipped += 1

    wb.save(EXCEL_FILE)
    print(f"Saved {EXCEL_FILE}  — updated {updated} rows, {skipped} rows unchanged.")


def update_excel_from_cache():
    """Re-apply saved results without re-scraping."""
    if not CACHE_FILE.exists():
        print(f"Cache file {CACHE_FILE} not found. Run the full scrape first.")
        sys.exit(1)
    with open(CACHE_FILE) as f:
        results = json.load(f)
    update_excel(results)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    global SEARCH_ENDPOINT

    # Re-apply from cache if requested
    if "--from-cache" in sys.argv:
        update_excel_from_cache()
        return

    if SEARCH_ENDPOINT is None:
        SEARCH_ENDPOINT = discover_search_endpoint()
        print(f"Using endpoint: {SEARCH_ENDPOINT['method']} {SEARCH_ENDPOINT['url']}")

    hospitals = read_hospitals()
    print(f"Processing {len(hospitals)} hospitals …\n")

    # Load partial cache if script was interrupted
    results: dict[str, dict] = {}
    if CACHE_FILE.exists():
        with open(CACHE_FILE) as f:
            results = json.load(f)
        print(f"Resuming from cache ({len(results)} already done).\n")

    try:
        for i, hosp in enumerate(hospitals, 1):
            code = hosp["code"]
            name = hosp["name"]

            if code in results:
                print(f"[{i:3d}/{len(hospitals)}] {code:12s} {name[:30]:<30}  (cached)")
                continue

            print(f"[{i:3d}/{len(hospitals)}] {code:12s} {name[:30]:<30}", end="  ", flush=True)
            data = search_hospital(SEARCH_ENDPOINT, code, name)

            if data:
                results[code] = data
                print(f"✓  ICU={data['icu']:4d}  Burn={data['burn_icu']}")
            else:
                print("✗  not found")

            time.sleep(0.4)  # polite delay

            # Auto-save every 20 hospitals
            if i % 20 == 0:
                with open(CACHE_FILE, "w") as f:
                    json.dump(results, f, ensure_ascii=False, indent=2)

    finally:
        with open(CACHE_FILE, "w") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print(f"\nResults cached to {CACHE_FILE}")

    update_excel(results)


if __name__ == "__main__":
    main()
