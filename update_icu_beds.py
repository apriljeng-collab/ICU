#!/usr/bin/env python3
"""
Fetch ICU bed numbers from openinfo.mohw.gov.tw and update the Excel file.

Install dependencies:
    pip install requests openpyxl beautifulsoup4 lxml

Run:
    python update_icu_beds.py                   # full scrape + update Excel
    python update_icu_beds.py --test            # verify 5 sample hospitals first
    python update_icu_beds.py --test 10         # verify 10 sample hospitals
    python update_icu_beds.py --from-cache      # re-apply saved results without re-scraping
    python update_icu_beds.py --diff            # show changes without writing Excel (dry run)
    python update_icu_beds.py --stats           # distribution summary of current Excel data
    python update_icu_beds.py --export-csv      # export current Excel K/L data to icu_export.csv
    python update_icu_beds.py --import-csv FILE # update Excel from a CSV (code,campus,icu,burn_icu)

URL pattern (discovered from MOHW website):
    https://openinfo.mohw.gov.tw/web/C02?type=1&basAgencyId=XXXXXXXXXX&zoneCode=A
    basAgencyId = hospital code padded to 10 digits (e.g. 101090517 -> 0101090517)
    zoneCode    = A for 本院, B/C/D/... for other campuses (tried in order)

The script reads all rows from ALL sheet, fetches current bed counts per
hospital code + campus, then writes values to column K (加護病床數量)
and column L (燒傷加護病床數).  Progress is cached to icu_results.json
so you can resume if interrupted.
"""

import csv
import time
import re
import json
import sys
import random
from pathlib import Path
from collections import defaultdict
from typing import Dict, List, Optional, Tuple

import requests
from bs4 import BeautifulSoup
import openpyxl

RETRY_DELAYS = (2, 5, 10)  # seconds between retries (3 attempts total)

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

# Campus-name -> zone letter mapping built from the website's own listing.
# Format: {hospital_code_str: {campus_name: zone_letter}}
KNOWN_ZONES: Dict[str, Dict[str, str]] = {
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
    return code.zfill(10)


def hospital_url(code: str, zone: str = "A") -> str:
    return (
        f"{BASE_URL}/web/C02"
        f"?type=1&basAgencyId={pad_code(code)}&zoneCode={zone}"
    )


# ---------------------------------------------------------------------------
# Bed-count extraction
# ---------------------------------------------------------------------------

_RE_ICU_NUMBER  = re.compile(r"加護病床[^數量\d]*[數量：:\s]*(\d+)", re.S)
_RE_BURN_NUMBER = re.compile(r"燒傷加護病床[^數量\d]*[數量：:\s]*(\d+)", re.S)


def _first_match(pattern: "re.Pattern[str]", text: str) -> Optional[int]:
    m = pattern.search(text)
    return int(m.group(1)) if m else None


def parse_page(html: str) -> Optional[dict]:
    """
    Return {"icu": int, "burn_icu": int} or None if not found.
    Tries multiple strategies on the page HTML.
    """
    soup = BeautifulSoup(html, "lxml")

    # Strategy 1: table th/td pairs  (e.g. <th>加護病床數量</th><td>177</td>)
    for cell in soup.find_all(["th", "td"]):
        cell_text = cell.get_text(strip=True)
        if "加護病床" in cell_text and "燒傷" not in cell_text:
            sibling = cell.find_next_sibling(["td", "th"])
            if sibling:
                nums = re.findall(r"\d+", sibling.get_text())
                if nums:
                    icu  = int(nums[0])
                    burn = _find_burn_in_soup(soup)
                    return {"icu": icu, "burn_icu": burn or 0}

    # Strategy 2: dl/dt+dd definition list  (newer MOHW layout)
    for dt in soup.find_all("dt"):
        if "加護病床" in dt.get_text() and "燒傷" not in dt.get_text():
            dd = dt.find_next_sibling("dd")
            if dd:
                nums = re.findall(r"\d+", dd.get_text())
                if nums:
                    icu  = int(nums[0])
                    burn = _find_burn_dt(soup)
                    return {"icu": icu, "burn_icu": burn or 0}

    # Strategy 3: scan full text with regex
    full_text = soup.get_text(" ", strip=True)
    icu = _first_match(_RE_ICU_NUMBER, full_text)
    if icu is not None:
        burn = _first_match(_RE_BURN_NUMBER, full_text) or 0
        return {"icu": icu, "burn_icu": burn}

    # Strategy 4: embedded JSON / script tags
    for script in soup.find_all("script"):
        content = script.string or ""
        if "加護" in content:
            icu = _first_match(_RE_ICU_NUMBER, content)
            if icu is not None:
                burn = _first_match(_RE_BURN_NUMBER, content) or 0
                return {"icu": icu, "burn_icu": burn}

    return None


def _find_burn_dt(soup: BeautifulSoup) -> Optional[int]:
    for dt in soup.find_all("dt"):
        if "燒傷" in dt.get_text():
            dd = dt.find_next_sibling("dd")
            if dd:
                nums = re.findall(r"\d+", dd.get_text())
                if nums:
                    return int(nums[0])
    return None


def _find_burn_in_soup(soup: BeautifulSoup) -> Optional[int]:
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

def discover_zones(code: str, expected_campuses: List[str]) -> Dict[str, str]:
    """
    Probe A, B, C, … until we've found all expected_campuses or run out.
    Returns {campus_name: zone_letter}.
    """
    if code in KNOWN_ZONES:
        return KNOWN_ZONES[code]

    mapping: Dict[str, str] = {}
    found_campuses: set = set()

    for letter in ZONE_LETTERS:
        if len(found_campuses) >= len(expected_campuses):
            break
        url = hospital_url(code, letter)
        try:
            resp = SESSION.get(url, timeout=15)
        except requests.RequestException:
            continue
        if resp.status_code != 200:
            if letter != "A":
                break
            continue

        soup = BeautifulSoup(resp.text, "lxml")
        full_text = soup.get_text(" ", strip=True)

        for campus in expected_campuses:
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

def _get_with_retry(url: str) -> Optional[requests.Response]:
    """GET with up to 3 retries; handles 429 rate-limiting."""
    for attempt, delay in enumerate((*RETRY_DELAYS, None), 1):
        try:
            resp = SESSION.get(url, timeout=15)
        except requests.RequestException as exc:
            if delay is None:
                print(f"      Network error: {exc}", end="")
                return None
            time.sleep(delay)
            continue

        if resp.status_code == 429:
            wait = int(resp.headers.get("Retry-After", delay or 30))
            print(f"      Rate-limited; waiting {wait}s …", end="", flush=True)
            time.sleep(wait)
            continue

        return resp

    return None


def fetch_hospital(code: str, campus: str) -> Optional[dict]:
    """
    Fetch bed data for one (code, campus) pair.
    Returns {"icu": int, "burn_icu": int} or None.
    """
    zone = "A"
    if campus and campus not in ("本院", ""):
        zones = KNOWN_ZONES.get(code, {})
        if campus in zones:
            zone = zones[campus]
        else:
            for letter in ZONE_LETTERS:
                resp = _get_with_retry(hospital_url(code, letter))
                if resp is None or resp.status_code != 200:
                    break
                soup = BeautifulSoup(resp.text, "lxml")
                full_text = soup.get_text(" ", strip=True)
                core = re.sub(r"[院區分院本]+$", "", campus).strip()
                if core and core in full_text:
                    zone = letter
                    KNOWN_ZONES.setdefault(code, {})[campus] = letter
                    break
                time.sleep(0.2)

    resp = _get_with_retry(hospital_url(code, zone))
    if resp is None:
        return None
    if resp.status_code != 200:
        print(f"      HTTP {resp.status_code}", end="")
        return None

    return parse_page(resp.text)


# ---------------------------------------------------------------------------
# Excel I/O
# ---------------------------------------------------------------------------

def read_hospitals() -> List[dict]:
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["ALL"]
    hospitals = []
    for i, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
        code   = row[3]   # col D
        name   = row[4]   # col E
        campus = row[5]   # col F
        icu    = row[10]  # col K  (baseline value)
        if code:
            hospitals.append({
                "row":      i,
                "code":     str(code),
                "name":     str(name)   if name   else "",
                "campus":   str(campus) if campus else "本院",
                "old_icu":  int(icu) if isinstance(icu, (int, float)) else None,
            })
    return hospitals


def _eval_icu(val) -> Optional[int]:
    if isinstance(val, (int, float)):
        return int(val)
    if isinstance(val, str) and val.startswith("="):
        try:
            return int(eval(val[1:]))  # noqa: S307
        except Exception:
            pass
    return None


def update_excel(results: Dict[str, dict], dry_run: bool = False) -> List[Tuple]:
    """
    Write ICU counts to ALL sheet columns K and L.
    Returns list of (row, name, campus, old_icu, new_icu, old_burn, new_burn) for changed rows.
    """
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["ALL"]

    updated = skipped = unchanged = 0
    changes = []

    for row in ws.iter_rows(min_row=3):
        code   = str(row[3].value) if row[3].value else None
        name   = str(row[4].value) if row[4].value else ""
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
                changes.append((row[0].row, name, campus, old_icu, new_icu, old_burn, new_burn))
                if not dry_run:
                    row[10].value = new_icu
                    row[11].value = new_burn
                updated += 1
            else:
                unchanged += 1
        else:
            skipped += 1

    if not dry_run:
        wb.save(EXCEL_FILE)
        print(f"\nSaved {EXCEL_FILE}")
    else:
        print("\n[dry-run] Excel NOT written.")

    print(f"  Updated:   {updated}")
    print(f"  Unchanged: {unchanged}")
    print(f"  No data:   {skipped}")
    return changes


def print_diff(changes: List[Tuple]) -> None:
    if not changes:
        print("\nNo changes detected.")
        return
    print(f"\n{'Row':>4}  {'Hospital':<28} {'Campus':<12}  {'ICU':>8}  {'Burn':>8}")
    print("-" * 72)
    for row_num, name, campus, old_icu, new_icu, old_burn, new_burn in changes:
        icu_str  = f"{old_icu or '?'} -> {new_icu}"
        burn_str = f"{old_burn or '?'} -> {new_burn}"
        print(f"{row_num:>4}  {name[:28]:<28} {campus[:12]:<12}  {icu_str:>8}  {burn_str:>8}")


def update_excel_from_cache(dry_run: bool = False) -> None:
    if not CACHE_FILE.exists():
        print(f"Cache file {CACHE_FILE} not found. Run the full scrape first.")
        sys.exit(1)
    with open(CACHE_FILE) as f:
        results = json.load(f)
    print(f"Loaded {len(results)} cached results from {CACHE_FILE}.")
    changes = update_excel(results, dry_run=dry_run)
    if "--diff" in sys.argv or dry_run:
        print_diff(changes)


# ---------------------------------------------------------------------------
# Test mode — verify a sample of hospitals before committing to full run
# ---------------------------------------------------------------------------

def run_test(hospitals: List[dict], n: int = 5) -> bool:
    """
    Fetch n random hospitals and report success/failure.
    Returns True if at least half succeed (connectivity OK).
    """
    sample = random.sample(hospitals, min(n, len(hospitals)))
    print(f"\n--- TEST MODE: sampling {len(sample)} hospitals ---\n")
    ok = fail = 0
    for hosp in sample:
        code   = hosp["code"]
        campus = hosp["campus"]
        name   = hosp["name"]
        print(f"  {code:12s} {campus:10s} {name[:25]:<25}", end="  ", flush=True)
        data = fetch_hospital(code, campus)
        if data:
            print(f"OK  ICU={data['icu']:4d}  Burn={data['burn_icu']}")
            ok += 1
        else:
            print("FAIL")
            fail += 1
        time.sleep(0.5)

    print(f"\nTest result: {ok}/{len(sample)} succeeded.")
    if ok == 0:
        print("ERROR: No hospitals could be fetched. Check network access to openinfo.mohw.gov.tw.")
        return False
    if ok < len(sample) // 2:
        print("WARNING: More than half of sampled hospitals failed.")
        return False
    print("Network access looks good. Proceeding with full run is recommended.")
    return True


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def export_csv(hospitals: List[dict], out: Path = Path("icu_export.csv")) -> None:
    """Export current Excel ICU data (cols K and L) to a CSV for manual editing."""
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb["ALL"]
    burn_by_key: Dict[str, int] = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        code   = str(row[3]) if row[3] else None
        campus = str(row[5]) if row[5] else "本院"
        burn   = row[11]
        if code:
            burn_by_key[f"{code}|{campus}"] = int(burn) if isinstance(burn, (int, float)) else 0

    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(["code", "campus", "name", "icu", "burn_icu"])
        for h in hospitals:
            key  = f"{h['code']}|{h['campus']}"
            burn = burn_by_key.get(key, 0)
            writer.writerow([h["code"], h["campus"], h["name"], h["old_icu"] or 0, burn])
    print(f"Exported {len(hospitals)} rows to {out}")
    print("Edit the 'icu' and 'burn_icu' columns, then run:")
    print(f"  python update_icu_beds.py --import-csv {out}")


def import_csv(csv_path: Path, dry_run: bool = False) -> None:
    """
    Load a CSV with columns: code, campus, icu, burn_icu
    and write the values back to the Excel file.
    """
    if not csv_path.exists():
        print(f"File not found: {csv_path}")
        sys.exit(1)

    results: Dict[str, dict] = {}
    skipped_rows = 0
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            code   = str(row.get("code", "")).strip()
            campus = str(row.get("campus", "本院")).strip() or "本院"
            try:
                icu  = int(row.get("icu") or 0)
                burn = int(row.get("burn_icu") or 0)
            except ValueError:
                skipped_rows += 1
                continue
            if code:
                results[f"{code}|{campus}"] = {"icu": icu, "burn_icu": burn}

    print(f"Loaded {len(results)} rows from {csv_path} ({skipped_rows} skipped).")
    # Also merge into cache file for future --from-cache use
    if not dry_run and CACHE_FILE.exists():
        with open(CACHE_FILE) as f:
            existing = json.load(f)
        existing.update(results)
        with open(CACHE_FILE, "w") as f:
            json.dump(existing, f, ensure_ascii=False, indent=2)
        print(f"Merged into {CACHE_FILE}.")

    changes = update_excel(results, dry_run=dry_run)
    if changes:
        print_diff(changes)


def print_stats(hospitals: List[dict]) -> None:
    """Print a distribution summary of current ICU bed counts from the Excel."""
    counts = [h["old_icu"] for h in hospitals if h["old_icu"] is not None]
    total  = sum(counts)
    mx     = max(counts)
    mn     = min(counts)
    avg    = total / len(counts) if counts else 0
    bands  = [(0, 0), (1, 20), (21, 50), (51, 100), (101, 200), (201, 9999)]
    labels = ["0 (no ICU)", "1-20", "21-50", "51-100", "101-200", "201+"]
    print(f"\nICU bed statistics ({len(counts)} hospitals with data):")
    print(f"  Total beds : {total}")
    print(f"  Average    : {avg:.1f}  |  Min: {mn}  |  Max: {mx}")
    print(f"\n  Distribution:")
    for (lo, hi), label in zip(bands, labels):
        n = sum(lo <= c <= hi for c in counts)
        bar = "█" * (n // 2)
        print(f"    {label:>10}  {n:4d}  {bar}")
    # Top 10 by ICU count
    top10 = sorted(
        [h for h in hospitals if h["old_icu"]],
        key=lambda h: h["old_icu"],
        reverse=True,
    )[:10]
    print(f"\n  Top 10 by ICU beds:")
    for h in top10:
        print(f"    {h['old_icu']:4d}  {h['name'][:30]:<30}  {h['campus']}")


def main() -> None:
    args = sys.argv[1:]
    dry_run = "--diff" in args

    if "--stats" in args:
        hospitals = read_hospitals()
        print_stats(hospitals)
        return

    if "--export-csv" in args:
        hospitals = read_hospitals()
        export_csv(hospitals)
        return

    if "--import-csv" in args:
        idx = args.index("--import-csv")
        if idx + 1 >= len(args):
            print("Usage: --import-csv <path-to-csv>")
            sys.exit(1)
        import_csv(Path(args[idx + 1]), dry_run=dry_run)
        return

    if "--from-cache" in args or dry_run:
        update_excel_from_cache(dry_run=dry_run)
        return

    hospitals = read_hospitals()
    print(f"Processing {len(hospitals)} hospital-campus rows …\n")

    # Test mode
    if "--test" in args:
        idx = args.index("--test")
        n = 5
        if idx + 1 < len(args) and args[idx + 1].isdigit():
            n = int(args[idx + 1])
        ok = run_test(hospitals, n)
        if not ok:
            sys.exit(1)
        cont = input("\nContinue with full scrape? [y/N] ").strip().lower()
        if cont != "y":
            sys.exit(0)

    # Load partial cache (resume support)
    results: Dict[str, dict] = {}
    if CACHE_FILE.exists():
        with open(CACHE_FILE) as f:
            results = json.load(f)
        print(f"Resuming from cache ({len(results)} already done).\n")

    # Pre-discover zones for multi-campus hospitals not already known
    multi_campus_codes: Dict[str, List[str]] = defaultdict(list)
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
                print(
                    f"[{i:3d}/{len(hospitals)}] {code:12s} {campus:10s} "
                    f"{name[:22]:<22}  (cached  ICU={d['icu']:4d} Burn={d['burn_icu']})"
                )
                continue

            print(
                f"[{i:3d}/{len(hospitals)}] {code:12s} {campus:10s} {name[:22]:<22}",
                end="  ",
                flush=True,
            )
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
                print(f"  [checkpoint saved at {i}]")

    finally:
        with open(CACHE_FILE, "w") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        print(f"\nResults cached to {CACHE_FILE}")

    changes = update_excel(results)
    if changes:
        print_diff(changes)


if __name__ == "__main__":
    main()
