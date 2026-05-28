#!/usr/bin/env python3
"""
Fetch ICU bed numbers from openinfo.mohw.gov.tw and update the Excel file.

Install dependencies:
    pip install requests openpyxl beautifulsoup4 lxml

Run:
    python update_icu_beds.py
    python update_icu_beds.py --from-cache   # re-apply saved results without re-scraping

URL pattern (discovered from MOHW website):
    https://openinfo.mohw.gov.tw/web/C02?type=1&basAgencyId=XXXXXXXXXX&zoneCode=A
    basAgencyId = hospital code padded to 10 digits (e.g. 101090517 → 0101090517)
    zoneCode    = A for 本院, B/C/D/... for other campuses (tried in order)

The script reads all rows from ALL sheet, fetches current bed counts per
hospital code + campus, then writes values to column K (加護病床數量)
and column L (燒傷加護病床數).  Progress is cached to icu_results.json
so you can resume if interrupted.
"""

import time
import re
import json
import sys
from pathlib import Path
from collections import defaultdict

import requests
from bs4 import BeautifulSoup
import openpyxl

EXCEL_FILE = Path("Hospital_Chair_20240731_update.xlsx")
CACHE_FILE = Path("icu_results.json")
BASE_URL   = "https://openinfo.mohw.gov.tw"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
    "Referer": BASE_URL + "/",
}
SESSION = requests.Session()
SESSION.headers.update(HEADERS)

# ---------------------------------------------------------------------------
# URL / zone helpers
# ---------------------------------------------------------------------------

ZONE_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

# Campus-name → zone letter mapping built from the website's own listing.
# Populated at runtime by discover_zones(); pre-fill known values here.
#
# Format: {hospital_code_str: {campus_name: zone_letter}}
# When unknown, we iterate A→Z until the page returns valid data.
KNOWN_ZONES: dict[str, dict[str, str]] = {
    "101090517": {
        "仁愛院區": "A",
        "中興院區": "B",
        "忠孝院區": "C",
        "陽明院區": "D",
        "和平婦幼院區": "E",
        "松德院區": "F",
        "林森中醫昆明院區": "G",
    },
}


def pad_code(code: str) -> str:
    """Left-pad hospital code to 10 digits."""
    return code.zfill(10)


def hospital_url(code: str, zone: str = "A") -> str:
    return (
        f"{BASE_URL}/web/C02"
        f"?type=1&basAgencyId={pad_code(code)}&zoneCode={zone}"
    )


# ---------------------------------------------------------------------------
# Bed-count extraction
# ---------------------------------------------------------------------------

_RE_ICU_NUMBER = re.compile(r"加護病床[^數量\d]*[數量：:\s]*(\d+)", re.S)
_RE_BURN_NUMBER = re.compile(r"燒傷加護病床[^數量\d]*[數量：:\s]*(\d+)", re.S)


def _first_match(pattern: re.Pattern, text: str) -> int | None:
    m = pattern.search(text)
    return int(m.group(1)) if m else None


def parse_page(html: str) -> dict | None:
    """
    Return {"icu": int, "burn_icu": int} or None if not found.
    Tries multiple strategies on the page HTML.
    """
    soup = BeautifulSoup(html, "lxml")

    # Strategy 1: find table rows where the header cell contains 加護病床
    for cell in soup.find_all(["th", "td"]):
        cell_text = cell.get_text(strip=True)
        if "加護病床" in cell_text and "燒傷" not in cell_text:
            # value should be in the next sibling td
            sibling = cell.find_next_sibling(["td", "th"])
            if sibling:
                nums = re.findall(r"\d+", sibling.get_text())
                if nums:
                    icu = int(nums[0])
                    burn = _find_burn_in_soup(soup)
                    return {"icu": icu, "burn_icu": burn or 0}

    # Strategy 2: scan full text with regex
    full_text = soup.get_text(" ", strip=True)
    icu = _first_match(_RE_ICU_NUMBER, full_text)
    if icu is not None:
        burn = _first_match(_RE_BURN_NUMBER, full_text) or 0
        return {"icu": icu, "burn_icu": burn}

    # Strategy 3: embedded JSON (some MOHW pages embed data as JSON)
    for script in soup.find_all("script"):
        content = script.string or ""
        if "加護" in content:
            icu = _first_match(_RE_ICU_NUMBER, content)
            if icu is not None:
                burn = _first_match(_RE_BURN_NUMBER, content) or 0
                return {"icu": icu, "burn_icu": burn}

    return None


def _find_burn_in_soup(soup: BeautifulSoup) -> int | None:
    for cell in soup.find_all(["th", "td"]):
        if "燒傷" in cell.get_text(strip=True):
            sibling = cell.find_next_sibling(["td", "th"])
            if sibling:
                nums = re.findall(r"\d+", sibling.get_text())
                if nums:
                    return int(nums[0])
    return None


# ---------------------------------------------------------------------------
# Zone-code discovery
# ---------------------------------------------------------------------------

def discover_zones(code: str, expected_campuses: list[str]) -> dict[str, str]:
    """
    For a multi-campus hospital, probe A, B, C, … until we've found all
    expected_campuses or run out of letters.  Returns {campus_name: zone}.
    """
    if code in KNOWN_ZONES:
        return KNOWN_ZONES[code]

    mapping: dict[str, str] = {}
    found_campuses: set[str] = set()

    for letter in ZONE_LETTERS:
        if len(found_campuses) >= len(expected_campuses):
            break
        url = hospital_url(code, letter)
        try:
            resp = SESSION.get(url, timeout=15)
        except requests.RequestException:
            continue
        if resp.status_code != 200:
            if letter != "A":  # A might just be wrong, keep going
                break
            continue

        soup = BeautifulSoup(resp.text, "lxml")
        full_text = soup.get_text(" ", strip=True)

        # Find which campus this page is for
        for campus in expected_campuses:
            # Strip 院區/分院/etc. for fuzzy match
            core = re.sub(r"[院區分院本]+$", "", campus).strip()
            if core and core in full_text:
                mapping[campus] = letter
                found_campuses.add(campus)
                break

        time.sleep(0.3)

    KNOWN_ZONES[code] = mapping
    return mapping


# ---------------------------------------------------------------------------
# Per-hospital fetching
# ---------------------------------------------------------------------------

def fetch_hospital(code: str, campus: str) -> dict | None:
    """
    Fetch bed data for one (code, campus) pair.
    Returns {"icu": int, "burn_icu": int} or None.
    """
    # Determine zone letter
    zone = "A"
    if campus and campus not in ("本院", ""):
        zones = KNOWN_ZONES.get(code, {})
        if campus in zones:
            zone = zones[campus]
        else:
            # Probe until we find the right campus page
            for letter in ZONE_LETTERS:
                url = hospital_url(code, letter)
                try:
                    resp = SESSION.get(url, timeout=15)
                except requests.RequestException:
                    continue
                if resp.status_code != 200:
                    break
                soup = BeautifulSoup(resp.text, "lxml")
                full_text = soup.get_text(" ", strip=True)
                core = re.sub(r"[院區分院本]+$", "", campus).strip()
                if core and core in full_text:
                    zone = letter
                    KNOWN_ZONES.setdefault(code, {})[campus] = letter
                    break
                time.sleep(0.2)

    url = hospital_url(code, zone)
    try:
        resp = SESSION.get(url, timeout=15)
    except requests.RequestException as e:
        print(f"      Network error: {e}", end="")
        return None

    if resp.status_code != 200:
        print(f"      HTTP {resp.status_code}", end="")
        return None

    return parse_page(resp.text)


# ---------------------------------------------------------------------------
# Excel I/O
# ---------------------------------------------------------------------------

def read_hospitals() -> list[dict]:
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["ALL"]
    hospitals = []
    for i, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
        code = row[3]   # col D
        name = row[4]   # col E
        campus = row[5] # col F
        if code:
            hospitals.append({
                "row":    i,
                "code":   str(code),
                "name":   str(name)   if name   else "",
                "campus": str(campus) if campus else "本院",
            })
    return hospitals


def _eval_icu(val) -> int | None:
    """Evaluate an Excel formula or return the numeric value."""
    if isinstance(val, (int, float)):
        return int(val)
    if isinstance(val, str) and val.startswith("="):
        try:
            return int(eval(val[1:]))
        except Exception:
            pass
    return None


def update_excel(results: dict[str, dict]):
    """
    results key = "{code}|{campus}"
    Write ICU counts to ALL sheet columns K and L.
    """
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["ALL"]

    updated = skipped = unchanged = 0
    for row in ws.iter_rows(min_row=3):
        code   = str(row[3].value) if row[3].value else None
        campus = str(row[5].value) if row[5].value else "本院"
        if not code:
            continue

        key = f"{code}|{campus}"
        if key in results:
            new_icu  = results[key]["icu"]
            new_burn = results[key]["burn_icu"]
            old_icu  = _eval_icu(row[10].value)
            old_burn = row[11].value if isinstance(row[11].value, int) else None

            if old_icu != new_icu or old_burn != new_burn:
                row[10].value = new_icu   # col K
                row[11].value = new_burn  # col L
                updated += 1
            else:
                unchanged += 1
        else:
            skipped += 1

    wb.save(EXCEL_FILE)
    print(f"\nSaved {EXCEL_FILE}")
    print(f"  Updated:   {updated}")
    print(f"  Unchanged: {unchanged}")
    print(f"  No data:   {skipped}")


def update_excel_from_cache():
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
    if "--from-cache" in sys.argv:
        update_excel_from_cache()
        return

    hospitals = read_hospitals()
    print(f"Processing {len(hospitals)} hospital-campus rows …\n")

    # Load partial cache (resume support)
    results: dict[str, dict] = {}
    if CACHE_FILE.exists():
        with open(CACHE_FILE) as f:
            results = json.load(f)
        print(f"Resuming from cache ({len(results)} already done).\n")

    # Pre-discover zones for known multi-campus hospitals
    multi_campus_codes = defaultdict(list)
    for h in hospitals:
        multi_campus_codes[h["code"]].append(h["campus"])
    for code, campuses in multi_campus_codes.items():
        if len(campuses) > 1 and code not in KNOWN_ZONES:
            print(f"Discovering zones for code {code} ({campuses[0][:20]}…)")
            discover_zones(code, campuses)

    try:
        for i, hosp in enumerate(hospitals, 1):
            code   = hosp["code"]
            campus = hosp["campus"]
            name   = hosp["name"]
            key    = f"{code}|{campus}"

            if key in results:
                d = results[key]
                print(f"[{i:3d}/{len(hospitals)}] {code:12s} {campus:10s} {name[:22]:<22}  (cached  ICU={d['icu']:4d} Burn={d['burn_icu']})")
                continue

            print(f"[{i:3d}/{len(hospitals)}] {code:12s} {campus:10s} {name[:22]:<22}", end="  ", flush=True)
            data = fetch_hospital(code, campus)

            if data:
                results[key] = data
                print(f"✓  ICU={data['icu']:4d}  Burn={data['burn_icu']}")
            else:
                print("✗  not found")

            time.sleep(0.4)

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
