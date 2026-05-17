#!/usr/bin/env python3
"""
funda_watcher.py

- Crawl Funda search results (pagination).
- Collect listing data (street, house nr, addition, postal code, city, url).
- Compare to local address list (CSV / Excel / Google Sheets).
- Send Telegram alerts when matches found (avoid duplicate alerts).
- Log activity and errors.
- Throttle requests / basic backoff.
- Persist seen listings to avoid duplicate notifications.

Requirements (pip):
    pip install requests beautifulsoup4 pandas gspread oauth2client python-dotenv openpyxl

Notes:
- For Google Sheets: provide a service account JSON file and set GOOGLE_SHEET_ID.
- For CSV/Excel: set ADDRESS_FILE path.
- Configure TELEGRAM_BOT_TOKEN and TELEGRAM_CHAT_ID.
"""

import os
import re
import time
import json
import random
import logging
import csv
from typing import List, Dict, Optional, Tuple
import requests
try:
    from curl_cffi.requests import Session as CurlSession
    CURL_CFFI_AVAILABLE = True
except ImportError:
    CURL_CFFI_AVAILABLE = False
from bs4 import BeautifulSoup
import pandas as pd
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv(dotenv_path=os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

# Optional Google Sheets imports
try:
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    GS_AVAILABLE = True
except Exception:
    GS_AVAILABLE = False

# -------------------------
# Configuration (env vars)
# -------------------------
NDA_BASE = "https://www.funda.nl/zoeken/koop"
# You can start from a particular view param (user gave ?search_result=2)
DEFAULT_PARAMS = {"search_result": "2"}  # will add 'page' param when paginating

# Data source: set either GOOGLE_SHEET_ID (and GOOGLE_SA_FILE) or ADDRESS_FILE path (CSV or Excel)
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_SA_FILE = os.getenv("GOOGLE_SA_FILE")
ADDRESS_FILE = os.getenv("ADDRESS_FILE", "addresses.csv")  # fallback CSV/XLSX file

# Telegram
TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")

# Files to persist state/logs
SEEN_DB_FILE = os.getenv("SEEN_DB_FILE", "seen_listings.json")
LOG_FILE = os.getenv("LOG_FILE", "funda_watcher.log")
LISTINGS_NDJSON_FILE = os.getenv("LISTINGS_NDJSON_FILE", "listings.ndjson")
PRINT_GSHEET_HEAD_ROWS = int(os.getenv("PRINT_GSHEET_HEAD_ROWS", "5"))
FOUND_MATCHES_CSV_FILE = os.getenv("FOUND_MATCHES_CSV_FILE", "found_it.csv")
MASTER_DATASET_SHEET_NAME = os.getenv("MASTER_DATASET_SHEET_NAME", "Master Dataset")
MASTER_DATASET_SEEN_IDS_FILE = os.getenv("MASTER_DATASET_SEEN_IDS_FILE", "seen_all_listings_ids.json")
MASTER_DATASET_EXPORT_ENABLED = os.getenv("MASTER_DATASET_EXPORT_ENABLED", "true").strip().lower() in ("1", "true", "yes", "y")
FOUND_DATA_SHEET_NAME = os.getenv("FOUND_DATA_SHEET_NAME", "Found Data")
FOUND_MATCH_CSV_FIELDS = [
    "ts",
    "listing_id",
    "listing_street",
    "listing_house_number",
    "listing_addition",
    "listing_postal_code",
    "listing_city",
    "listing_url",
    "matched_street",
    "matched_house_number",
    "matched_addition",
    "matched_postal_code",
    "matched_city",
]

MASTER_DATASET_EXPORT_FIELDS = [
    "ts",
    "page",
    "id",
    "street",
    "house_number",
    "addition",
    "postal_code",
    "city",
    "url",
    "title_text",
]

# Request / crawling options
REQUEST_HEADERS = {
    "User-Agent": os.getenv(
        "USER_AGENT",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
    "Accept-Language": "nl-NL,nl;q=0.9,en-US;q=0.8,en;q=0.7",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
}
CURL_IMPERSONATE = os.getenv("CURL_IMPERSONATE", "chrome131")  # browser TLS fingerprint to impersonate
PROXY_URL = os.getenv("PROXY_URL", "")  # optional: e.g. http://user:pass@proxyhost:port
REQUEST_TIMEOUT = 30
MIN_DELAY = 2.0   # seconds between page requests
MAX_DELAY = 5.0

# Matching sensitivity
POSTAL_REGEX = re.compile(r"\b(\d{4}\s?[A-Za-z]{2})\b")  # Dutch postal code pattern
LISTING_URL_PREFIX = "https://www.funda.nl"

# Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(LOG_FILE), logging.StreamHandler()]
)

# -------------------------
# Helper functions
# -------------------------
def load_seen_db(path: str) -> Dict[str, dict]:
    if os.path.exists(path):
        try:
            return json.load(open(path, "r", encoding="utf-8"))
        except Exception as e:
            logging.error("Failed to load seen db: %s", e)
            return {}
    return {}

def save_seen_db(path: str, db: Dict[str, dict]):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=2, ensure_ascii=False)

def load_seen_id_dict(path: str) -> Dict[str, bool]:
    if os.path.exists(path):
        try:
            data = json.load(open(path, "r", encoding="utf-8"))
            if isinstance(data, dict):
                return {str(k): bool(v) for k, v in data.items()}
        except Exception as e:
            logging.error("Failed to load seen id dict (%s): %s", path, e)
    return {}

def save_seen_id_dict(path: str, db: Dict[str, bool]):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(db, f, indent=2, ensure_ascii=False)

def init_gsheet_client():
    if not (GS_AVAILABLE and GOOGLE_SHEET_ID and GOOGLE_SA_FILE):
        return None
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_SA_FILE, scope)
    return gspread.authorize(creds)

def get_or_create_worksheet(sh, worksheet_title: str):
    """
    Return a gspread worksheet, creating it if missing.
    If the service account has no permission, returns None (and logs).
    """
    try:
        return sh.worksheet(worksheet_title)
    except Exception:
        try:
            return sh.add_worksheet(title=worksheet_title, rows=1000, cols=30)
        except Exception as e:
            logging.exception(
                "Cannot access/create worksheet '%s' (need Editor permission for service account): %s",
                worksheet_title,
                e,
            )
            return None

def ensure_master_dataset_headers(ws, headers: List[str]) -> None:
    """
    Ensure headers exist in row 1.
    If the sheet is empty or headers don't match, append a header row.
    """
    try:
        first_row = ws.row_values(1)
        if not first_row:
            ws.append_row(headers, value_input_option="RAW")
            return
        if len(first_row) < len(headers) or any(h not in first_row for h in headers):
            ws.append_row(headers, value_input_option="RAW")
    except Exception as e:
        logging.exception("Failed ensuring Master Dataset headers: %s", e)

def load_sheet_ids_column(ws, col_index: int, skip_header: bool = True) -> set:
    """
    Read all values from a specific column of a worksheet into a set.
    Used to pre-populate dedup sets from the actual sheet contents.
    col_index is 1-based.
    """
    ids: set = set()
    try:
        col_values = ws.col_values(col_index)
        start = 1 if skip_header else 0
        for val in col_values[start:]:
            val = str(val).strip()
            if val:
                ids.add(val)
        logging.info("Loaded %d existing IDs from sheet column %d for dedup", len(ids), col_index)
    except Exception as e:
        logging.warning("Could not load existing IDs from sheet (dedup may be incomplete): %s", e)
    return ids

def append_listings_to_master_dataset(
    ws,
    listings: List[dict],
    ts: int,
    page: int,
    export_seen_ids: Dict[str, bool],
) -> int:
    rows = []
    for listing_item in listings:
        key = str(listing_item.get("id") or "").strip()
        if not key:
            continue
        if export_seen_ids.get(key):
            continue
        export_seen_ids[key] = True
        rows.append([
            ts,
            page,
            listing_item.get("id") or "",
            listing_item.get("street") or "",
            listing_item.get("house_number") or "",
            listing_item.get("addition") or "",
            listing_item.get("postal_code") or "",
            listing_item.get("city") or "",
            listing_item.get("url") or "",
            listing_item.get("title_text") or "",
        ])

    if not rows:
        return 0

    try:
        ws.append_rows(rows, value_input_option="RAW")
        return len(rows)
    except Exception as e:
        logging.exception("Failed to append listings to Master Dataset: %s", e)
        # Roll back dedupe keys for this batch (best-effort).
        for r in rows:
            key = str(r[2] or "").strip()
            if key:
                export_seen_ids.pop(key, None)
        return 0

def append_match_to_found_data_sheet(
    ws,
    match_row: dict,
    headers: List[str],
    found_seen_ids: Optional[set] = None,
) -> bool:
    """
    Append a single matched listing row to the 'Found Data' Google Sheet.
    Skips if listing_id is already in found_seen_ids (dedup against sheet).
    Ensures headers exist in row 1 before appending.
    Returns True if row was written, False if skipped as duplicate.
    """
    dedup_key = str(match_row.get("listing_id") or "").strip()
    if found_seen_ids is not None and dedup_key and dedup_key in found_seen_ids:
        logging.debug("Found Data: skipping duplicate listing_id=%s", dedup_key)
        return False
    try:
        first_row = ws.row_values(1)
        if not first_row:
            ws.append_row(headers, value_input_option="RAW")
        row = [match_row.get(h, "") for h in headers]
        ws.append_row(row, value_input_option="RAW")
        if found_seen_ids is not None and dedup_key:
            found_seen_ids.add(dedup_key)
        return True
    except Exception as e:
        logging.exception("Failed to append match to Found Data sheet: %s", e)
        return False

def append_ndjson(path: str, obj: dict):
    # Append one JSON object per line; flush for "realtime" persistence.
    with open(path, "a", encoding="utf-8") as f:
        f.write(json.dumps(obj, ensure_ascii=False) + "\n")
        f.flush()

def append_found_match_to_csv(path: str, row: dict, fieldnames: List[str]) -> None:
    """
    Append a single match row to CSV, including header if the file is new/empty.
    """
    file_exists = os.path.exists(path)
    file_empty = (not file_exists) or os.path.getsize(path) == 0

    with open(path, "a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if file_empty:
            writer.writeheader()
        writer.writerow(row)

def read_addresses_from_csv_or_excel(path: str) -> pd.DataFrame:
    if path.lower().endswith((".xls", ".xlsx")):
        df = pd.read_excel(path, dtype=str)
    else:
        df = pd.read_csv(path, dtype=str)
    # normalize column names
    df.columns = [c.strip().lower() for c in df.columns]
    return df

def read_addresses_from_gsheet(sheet_id: str, sa_file: str, worksheet_name: Optional[str]=None) -> pd.DataFrame:
    if not GS_AVAILABLE:
        raise RuntimeError("gspread not installed or not available in environment.")
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name(sa_file, scope)
    client = gspread.authorize(creds)
    sh = client.open_by_key(sheet_id)
    ws = sh.sheet1 if worksheet_name is None else sh.worksheet(worksheet_name)
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame()
    df = pd.DataFrame(data)
    df.columns = [c.strip().lower() for c in df.columns]
    # Print a small preview so you can verify the script reads the right sheet/columns.
    try:
        preview_rows = df.head(PRINT_GSHEET_HEAD_ROWS).to_dict(orient="records")
        logging.info(
            "Google Sheet preview (first %d rows): %s",
            PRINT_GSHEET_HEAD_ROWS,
            preview_rows,
        )
    except Exception:
        # Preview is non-critical.
        pass
    return df

def normalize_addr_row(row: dict) -> dict:
    """Return normalized address dict with keys street, house_number, addition, postal_code, city"""
    s = str(row.get("street") or row.get("straat") or "").strip()
    num = str(row.get("house number") or row.get("house_number") or row.get("nummer") or row.get("house") or "").strip()
    add = str(row.get("addition") or row.get("toevoeging") or "").strip()
    pc = str(row.get("postal code") or row.get("postcode") or row.get("postal_code") or "").strip()
    city = str(row.get("city") or row.get("plaats") or "").strip()
    # Normalise postal code (no space + uppercase)
    pc = pc.replace(" ", "").upper() if pc else ""
    if pc and len(pc) == 6:
        pc = pc[:4] + " " + pc[4:]
    return {"street": s, "house_number": num, "addition": add, "postal_code": pc, "city": city}

def read_address_list() -> List[dict]:
    # Try Google Sheets first if configured
    if GOOGLE_SHEET_ID and GOOGLE_SA_FILE:
        logging.info("Reading address list from Google Sheet %s", GOOGLE_SHEET_ID)
        df = read_addresses_from_gsheet(GOOGLE_SHEET_ID, GOOGLE_SA_FILE)
    else:
        logging.info("Reading address list from file %s", ADDRESS_FILE)
        df = read_addresses_from_csv_or_excel(ADDRESS_FILE)
    rows = []
    for _, r in df.iterrows():
        rows.append(normalize_addr_row(r.to_dict()))
    logging.info("Loaded %d address rows", len(rows))
    return rows

def parse_listing_address(text: str) -> dict:
    """
    Parse address text from listing into components.
    Expect patterns like:
        "Main Road 1786" and separate "2157 ND Abbenes"
    The function is resilient: it looks for postal code, then extracts street/number from preceding part.
    """
    text = re.sub(r"\s+", " ", text).strip()
    postal_match = POSTAL_REGEX.search(text)
    postal = postal_match.group(1).replace(" ", "").upper() if postal_match else ""
    if postal and len(postal) == 6:
        postal = postal[:4] + " " + postal[4:]
    # Try to extract city (word after postal code)
    city = ""
    if postal_match:
        after = text[postal_match.end():].strip()
        if after:
            city = after.split()[0]
    # Try to extract street and number: first line or start of text before postal code
    before = text[:postal_match.start()] if postal_match else text
    # Often listings use two lines; but we get a single string. Look for trailing number.
    street = before.strip()
    house_number = ""
    addition = ""
    # Split last token that contains digits
    m = re.search(r"(.+?)\s+(\d+\w*)(?:\s*[-/\\]?\s*([A-Za-z0-9]+))?$", street)
    if m:
        street = m.group(1).strip()
        numpart = m.group(2).strip()
        # split numeric and alpha addition if any
        m2 = re.match(r"^(\d+)([A-Za-z0-9-]*)$", numpart)
        if m2:
            house_number = m2.group(1)
            addition = (m2.group(2) or "").strip()
        else:
            house_number = numpart
    else:
        # fallback: try to find first number in text
        m3 = re.search(r"(\d+)", street)
        if m3:
            house_number = m3.group(1)
            street = street[:m3.start()].strip()
    return {"street": street, "house_number": house_number, "addition": addition, "postal_code": postal, "city": city}

def send_telegram(text: str) -> Tuple[bool, str]:
    if not TELEGRAM_BOT_TOKEN or not TELEGRAM_CHAT_ID:
        logging.warning("Telegram not configured; skipping sending message.")
        return False, "not_configured"
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    payload = {"chat_id": TELEGRAM_CHAT_ID, "text": text, "disable_web_page_preview": True}
    try:
        r = requests.post(url, data=payload, timeout=15)
        if r.status_code == 200:
            return True, r.text
        else:
            logging.error("Telegram API returned status %s: %s", r.status_code, r.text)
            return False, r.text
    except Exception as e:
        logging.exception("Telegram send failed: %s", e)
        return False, str(e)

def extract_listings_from_page(html: str) -> List[dict]:
    soup = BeautifulSoup(html, "html.parser")
    results = []
    # Preferred: data-testid="listingDetailsAddress"
    for a in soup.select('a[data-testid="listingDetailsAddress"]'):
        try:
            title = " ".join(a.stripped_strings)
            # There is often an inner structure: street + next element contains "2157 ND Abbenes"
            # sometimes the postal/city is inside sibling div
            postal_city = ""
            # Attempt to get next text element
            sibling = a.find_next(string=re.compile(r"\d{4}\s?[A-Za-z]{2}"))
            if sibling:
                postal_city = sibling.strip()
            # fallback: full area text
            full_text = title + " " + (postal_city or "")
            parsed = parse_listing_address(full_text)
            href = a.get("href") or ""
            url = LISTING_URL_PREFIX + href if href.startswith("/") else href
            # listing id from url
            idm = None
            id_match = re.search(r"/(\d{6,})/?$", url)
            if not id_match:
                id_match = re.search(r"/(\d{6,})/", url)
            if id_match:
                idm = id_match.group(1)
            results.append({
                "id": idm or url,
                "street": parsed["street"],
                "house_number": parsed["house_number"],
                "addition": parsed["addition"],
                "postal_code": parsed["postal_code"],
                "city": parsed["city"],
                "url": url,
                "title_text": full_text
            })
        except Exception as e:
            logging.exception("Error parsing listing element: %s", e)
    # if none found with data-testid, fallback search for /detail/koop/ links
    if not results:
        for a in soup.find_all("a", href=True):
            if "/detail/koop/" in a["href"]:
                full = " ".join(a.stripped_strings)
                # look for address container relative to this anchor
                parent_text = a.parent.get_text(" ", strip=True) if a.parent else full
                parsed = parse_listing_address(parent_text)
                href = a["href"]
                url = LISTING_URL_PREFIX + href if href.startswith("/") else href
                idm = re.search(r"/(\d{6,})/", url)
                idv = idm.group(1) if idm else url
                results.append({
                    "id": idv,
                    "street": parsed["street"],
                    "house_number": parsed["house_number"],
                    "addition": parsed["addition"],
                    "postal_code": parsed["postal_code"],
                    "city": parsed["city"],
                    "url": url,
                    "title_text": parent_text
                })
    # deduplicate by id/url
    seen_ids = set()
    dedup = []
    for r in results:
        if r["id"] in seen_ids:
            continue
        seen_ids.add(r["id"])
        dedup.append(r)
    return dedup

def page_has_next_disabled(html: str) -> bool:
    soup = BeautifulSoup(html, "html.parser")
    # Funda pagination markup/labels vary (NL/EN, rel="next", data-testid, etc.)
    # Treat "no usable next link" as last page.
    candidates = []

    # Most reliable when present: rel="next"
    candidates.extend(soup.select('a[rel="next"]'))
    candidates.extend(soup.select('link[rel="next"]'))  # sometimes in <head>

    # Common accessibility labels
    candidates.extend(soup.select('a[aria-label="Next"]'))
    candidates.extend(soup.select('a[aria-label="Volgende"]'))

    # Data test ids seen on some builds
    candidates.extend(soup.select('a[data-testid*="pagination-next" i]'))
    candidates.extend(soup.select('button[data-testid*="pagination-next" i]'))

    # De-dupe while preserving order
    seen = set()
    dedup = []
    for el in candidates:
        key = (el.name, tuple(sorted((el.attrs or {}).items())))
        if key in seen:
            continue
        seen.add(key)
        dedup.append(el)

    if not dedup:
        # Pagination controls are sometimes not present in the HTML response (hydrated client-side).
        # Don't stop here; rely on repeat/no-listings safeguards instead.
        return False

    for el in dedup:
        aria_disabled = str(el.get("aria-disabled", "")).lower() == "true"
        disabled_attr = "disabled" in (el.attrs or {})
        cls = el.get("class")
        cls_str = " ".join(cls) if isinstance(cls, (list, tuple)) else (str(cls) if cls else "")
        looks_disabled = aria_disabled or disabled_attr or ("cursor-default" in cls_str) or ("disabled" in cls_str)

        href = el.get("href") if hasattr(el, "get") else None
        if el.name == "link":
            href = el.get("href")

        if looks_disabled:
            continue
        if href and str(href).strip():
            return False

    return True

# -------------------------
# Targeted per-address search
# -------------------------
def _slugify(text: str) -> str:
    """Convert a string to lowercase slug: remove non-alphanum chars, replace spaces with hyphens."""
    text = text.lower().strip()
    # Replace common Dutch special chars
    text = text.replace("é", "e").replace("ë", "e").replace("è", "e")
    text = text.replace("ü", "u").replace("ö", "o").replace("ä", "a")
    text = text.replace("ï", "i")
    # Replace spaces and underscores with hyphens
    text = re.sub(r"[\s_]+", "-", text)
    # Remove anything not alphanumeric or hyphen
    text = re.sub(r"[^a-z0-9\-]", "", text)
    # Collapse multiple hyphens
    text = re.sub(r"-+", "-", text).strip("-")
    return text


def build_targeted_search_url(city: str, street: str) -> Optional[str]:
    """
    Build a Funda targeted search URL using selected_area.
    Format: https://www.funda.nl/zoeken/koop?selected_area=["city/straat-street-name"]
    'straat-' is a constant prefix Funda uses for street-level searches.
    Returns None if city or street is empty.
    """
    city_slug = _slugify(city)
    street_slug = _slugify(street)
    if not city_slug or not street_slug:
        return None
    area = f"{city_slug}/straat-{street_slug}"
    return f'https://www.funda.nl/zoeken/koop?selected_area=["{area}"]'


def run_targeted_searches(
    address_list: List[dict],
    session,
    already_seen_ids: Optional[set] = None,
) -> List[dict]:
    """
    For each unique (city, street) pair in address_list, perform a targeted Funda
    search. Scrapes the results and returns a deduplicated list of *new* listings
    (those not in already_seen_ids).
    """
    new_listings: List[dict] = []
    collected_ids: set = set(already_seen_ids or set())
    searched_pairs: set = set()  # dedup: skip duplicate (city_slug, street_slug) combos

    for addr in address_list:
        city = (addr.get("city") or "").strip()
        street = (addr.get("street") or "").strip()
        if not city or not street:
            continue

        url = build_targeted_search_url(city, street)
        if not url:
            continue

        if url in searched_pairs:
            continue
        searched_pairs.add(url)

        logging.info("Targeted search: %s", url)
        try:
            r = session.get(url, timeout=REQUEST_TIMEOUT)
            if r.status_code != 200:
                logging.warning("Targeted search returned %s for %s", r.status_code, url)
                delay = random.uniform(MIN_DELAY, MAX_DELAY)
                time.sleep(delay)
                continue

            page_listings = extract_listings_from_page(r.text)
            logging.info("Targeted search for '%s / %s': found %d listing(s)", city, street, len(page_listings))

            for listing in page_listings:
                lid = listing.get("id") or listing.get("url") or ""
                if lid and lid not in collected_ids:
                    collected_ids.add(lid)
                    new_listings.append(listing)

        except Exception as e:
            logging.exception("Targeted search error for %s: %s", url, e)

        delay = random.uniform(MIN_DELAY, MAX_DELAY)
        time.sleep(delay)

    logging.info("Targeted searches complete. %d new unique listing(s) found.", len(new_listings))
    return new_listings


# -------------------------
# Matching logic
# -------------------------
def match_listing_to_address(listing: dict, address_list: List[dict]) -> Optional[dict]:
    def norm_postal(pc: str) -> str:
        return (pc or "").replace(" ", "").strip().upper()

    def norm_street(s: str) -> str:
        return (s or "").lower().replace(".", "").strip()

    def norm_city(s: str) -> str:
        return (s or "").lower().strip()

    # Primary: postal code + house number
    for a in address_list:
        lpc = norm_postal(listing.get("postal_code") or "")
        apc = norm_postal(a.get("postal_code") or "")
        if apc and lpc and apc == lpc and a.get("house_number") == listing.get("house_number"):
            return a

    # Fallback: street + house number + city
    for a in address_list:
        if not (a.get("street") and listing.get("street") and a.get("city") and listing.get("city")):
            continue
        if a.get("house_number") != listing.get("house_number"):
            continue
        if norm_street(a.get("street")) != norm_street(listing.get("street")):
            continue
        if norm_city(a.get("city")) != norm_city(listing.get("city")):
            continue
        return a

    return None

# -------------------------
# Main crawl + match flow
# -------------------------
def run_once():
    logging.info("Run started.")
    address_list = read_address_list()
    seen_db = load_seen_db(SEEN_DB_FILE)
    export_seen_ids = load_seen_id_dict(MASTER_DATASET_SEEN_IDS_FILE)

    # Use curl_cffi for browser TLS impersonation (bypasses Cloudflare bot detection)
    if CURL_CFFI_AVAILABLE:
        session = CurlSession(
            impersonate=CURL_IMPERSONATE,
            proxies={"http": PROXY_URL, "https": PROXY_URL} if PROXY_URL else None,
        )
        session.headers.update(REQUEST_HEADERS)
        logging.info("Using curl_cffi session (impersonate=%s, proxy=%s)", CURL_IMPERSONATE, PROXY_URL or "none")
    else:
        logging.warning("curl_cffi not available — falling back to requests (may be blocked by Cloudflare)")
        session = requests.Session()
        session.headers.update(REQUEST_HEADERS)
        if PROXY_URL:
            session.proxies.update({"http": PROXY_URL, "https": PROXY_URL})
    master_ws = None
    found_ws = None
    found_seen_ids: set = set()  # populated from Found Data sheet below
    if GS_AVAILABLE and GOOGLE_SHEET_ID and GOOGLE_SA_FILE:
        try:
            gclient = init_gsheet_client()
            if gclient is not None:
                sh = gclient.open_by_key(GOOGLE_SHEET_ID)
                if MASTER_DATASET_EXPORT_ENABLED:
                    master_ws = get_or_create_worksheet(sh, MASTER_DATASET_SHEET_NAME)
                    if master_ws is not None:
                        ensure_master_dataset_headers(master_ws, MASTER_DATASET_EXPORT_FIELDS)
                        logging.info("Master Dataset export enabled: %s", MASTER_DATASET_SHEET_NAME)
                        # Load existing IDs from the sheet (col 3 = 'id') to seed dedup set.
                        # MASTER_DATASET_EXPORT_FIELDS = [ts, page, id, street, ...]
                        sheet_ids = load_sheet_ids_column(master_ws, col_index=3, skip_header=True)
                        new_ids = sheet_ids - set(export_seen_ids.keys())
                        if new_ids:
                            logging.info("Seeding Master Dataset dedup with %d IDs read from sheet", len(new_ids))
                            for sid in new_ids:
                                export_seen_ids[sid] = True
                            save_seen_id_dict(MASTER_DATASET_SEEN_IDS_FILE, export_seen_ids)
                found_ws = get_or_create_worksheet(sh, FOUND_DATA_SHEET_NAME)
                if found_ws is not None:
                    logging.info("Found Data sheet ready: %s", FOUND_DATA_SHEET_NAME)
                    # Load existing listing_id values from Found Data (col 2 = 'listing_id') to seed dedup set.
                    # FOUND_MATCH_CSV_FIELDS = [ts, listing_id, ...]
                    found_seen_ids = load_sheet_ids_column(found_ws, col_index=2, skip_header=True)
                    logging.info("Found Data dedup set has %d existing entries", len(found_seen_ids))
        except Exception as e:
            logging.exception("Google Sheets setup error: %s", e)
    page = 1
    processed = 0
    matches_found = 0
    error_flag = False
    last_page_signature: Optional[Tuple[str, ...]] = None
    repeated_signature_count = 0
    repeat_stop_threshold = int(os.getenv("REPEAT_STOP_THRESHOLD", "2"))

    max_pages_env = os.getenv("MAX_PAGES")
    max_pages = int(max_pages_env) if (max_pages_env and max_pages_env.strip()) else None
    logging.info(
        "Pagination settings: MAX_PAGES=%s (None means unlimited), REPEAT_STOP_THRESHOLD=%d",
        max_pages if max_pages is not None else "None",
        repeat_stop_threshold,
    )

    while True:
        if max_pages is not None and page > max_pages:
            logging.info("Reached MAX_PAGES=%d — stopping.", max_pages)
            break
        params = DEFAULT_PARAMS.copy()
        params["page"] = str(page)
        try:
            logging.info("Fetching page %s with params=%s", page, params)
            r = session.get( NDA_BASE, params=params, timeout=REQUEST_TIMEOUT)
            if r.status_code != 200:
                logging.error("Non-200 from Funda for page %s: %s", page, r.status_code)
                send_telegram(f"[ERROR] Funda returned status {r.status_code} for page {page}")
                error_flag = True
                break
            listings = extract_listings_from_page(r.text)
            logging.info("Page %s: extracted %d listings", page, len(listings))
            processed += len(listings)
            if not listings:
                logging.info("No listings found on page %s — stopping pagination.", page)
                break

            # Stop if the *same page results* repeat N times in a row.
            # Use listing URLs (or ids) as the page signature.
            sig_items = []
            for listing_item in listings:
                sig_items.append(listing_item.get("url") or listing_item.get("id") or "")
            page_signature = tuple(sorted([s for s in sig_items if s]))
            if last_page_signature is not None and page_signature and page_signature == last_page_signature:
                repeated_signature_count += 1
                logging.info(
                    "Page %s repeated same results (%d/%d).",
                    page,
                    repeated_signature_count,
                    repeat_stop_threshold,
                )
                if repeated_signature_count >= repeat_stop_threshold:
                    logging.info(
                        "Same page results repeated %d times — stopping pagination.",
                        repeat_stop_threshold,
                    )
                    break
            else:
                repeated_signature_count = 0
            last_page_signature = page_signature or last_page_signature

            # Realtime persistence: append scraped listings immediately
            ts = int(time.time())
            for listing_item in listings:
                append_ndjson(LISTINGS_NDJSON_FILE, {"ts": ts, "page": page, "listing": listing_item})

            # Realtime export: append listings to Google Sheets "Master Dataset"
            if master_ws is not None and MASTER_DATASET_EXPORT_ENABLED:
                appended = append_listings_to_master_dataset(
                    master_ws,
                    listings,
                    ts,
                    page,
                    export_seen_ids,
                )
                if appended:
                    save_seen_id_dict(MASTER_DATASET_SEEN_IDS_FILE, export_seen_ids)

            for listing in listings:
                # match
                matched = match_listing_to_address(listing, address_list)
                if matched:
                    # only notify if not seen before (id key)
                    lid = listing["id"]
                    if lid not in seen_db:
                        matches_found += 1
                        msg = f"Match found\n{matched.get('street')} {matched.get('house_number')}{(' '+matched.get('addition')) if matched.get('addition') else ''}, {matched.get('city') or listing.get('city')}\n{listing.get('url')}"
                        logging.info("NEW MATCH: %s", msg)
                        send_telegram(msg)
                        now = int(time.time())
                        seen_db[lid] = {"listing": listing, "matched_to": matched, "first_seen": now}
                        match_row = {
                            "ts": now,
                            "listing_id": listing.get("id") or "",
                            "listing_street": listing.get("street") or "",
                            "listing_house_number": listing.get("house_number") or "",
                            "listing_addition": listing.get("addition") or "",
                            "listing_postal_code": listing.get("postal_code") or "",
                            "listing_city": listing.get("city") or "",
                            "listing_url": listing.get("url") or "",
                            "matched_street": matched.get("street") or "",
                            "matched_house_number": matched.get("house_number") or "",
                            "matched_addition": matched.get("addition") or "",
                            "matched_postal_code": matched.get("postal_code") or "",
                            "matched_city": matched.get("city") or "",
                        }
                        append_found_match_to_csv(
                            FOUND_MATCHES_CSV_FILE,
                            match_row,
                            FOUND_MATCH_CSV_FIELDS,
                        )
                        # Save match to Google Sheets "Found Data" tab
                        if found_ws is not None:
                            append_match_to_found_data_sheet(found_ws, match_row, FOUND_MATCH_CSV_FIELDS, found_seen_ids)
                        # Realtime persistence: save after every new match
                        save_seen_db(SEEN_DB_FILE, seen_db)
                    else:
                        logging.debug("Already seen listing %s - skipping notify", listing["id"])

            # Realtime persistence: checkpoint after each page
            save_seen_db(SEEN_DB_FILE, seen_db)

            # Stop when Next button becomes non-clickable (cursor-default/disabled/missing href).
            if page_has_next_disabled(r.text):
                logging.info("Next button disabled on page %s — stopping pagination.", page)
                break

            # throttle
            delay = random.uniform(MIN_DELAY, MAX_DELAY)
            logging.debug("Sleeping %.2f sec", delay)
            time.sleep(delay)
            page += 1
        except Exception as e:
            logging.exception("Error during page fetch/parsing: %s", e)
            send_telegram(f"[ERROR] Exception during crawl: {e}")
            error_flag = True
            break

    save_seen_db(SEEN_DB_FILE, seen_db)
    save_seen_id_dict(MASTER_DATASET_SEEN_IDS_FILE, export_seen_ids)

    # -------------------------------------------------------
    # STEP 2: Targeted per-address searches on Funda
    # Run *after* the general crawl + master-dataset write,
    # *before* notifications are sent.
    # -------------------------------------------------------
    logging.info("Starting targeted per-address Funda searches...")
    # Collect ids already retrieved in the main crawl so we only add genuinely new listings.
    main_crawl_ids: set = set()
    for lid in export_seen_ids:
        main_crawl_ids.add(lid)

    targeted_listings = run_targeted_searches(address_list, session, already_seen_ids=main_crawl_ids)

    if targeted_listings:
        ts_targeted = int(time.time())
        # Persist new targeted listings to ndjson
        for listing_item in targeted_listings:
            append_ndjson(LISTINGS_NDJSON_FILE, {"ts": ts_targeted, "page": "targeted", "listing": listing_item})

        # Append new targeted listings to Master Dataset sheet
        if master_ws is not None and MASTER_DATASET_EXPORT_ENABLED:
            appended = append_listings_to_master_dataset(
                master_ws,
                targeted_listings,
                ts_targeted,
                0,  # page=0 signals a targeted search result
                export_seen_ids,
            )
            if appended:
                save_seen_id_dict(MASTER_DATASET_SEEN_IDS_FILE, export_seen_ids)

        # Match targeted listings and send notifications
        for listing in targeted_listings:
            matched = match_listing_to_address(listing, address_list)
            if matched:
                lid = listing["id"]
                if lid not in seen_db:
                    matches_found += 1
                    msg = (
                        f"Match found (targeted search)\n"
                        f"{matched.get('street')} {matched.get('house_number')}"
                        f"{(' ' + matched.get('addition')) if matched.get('addition') else ''}, "
                        f"{matched.get('city') or listing.get('city')}\n"
                        f"{listing.get('url')}"
                    )
                    logging.info("NEW MATCH (targeted): %s", msg)
                    send_telegram(msg)
                    now = int(time.time())
                    seen_db[lid] = {"listing": listing, "matched_to": matched, "first_seen": now}
                    match_row = {
                        "ts": now,
                        "listing_id": listing.get("id") or "",
                        "listing_street": listing.get("street") or "",
                        "listing_house_number": listing.get("house_number") or "",
                        "listing_addition": listing.get("addition") or "",
                        "listing_postal_code": listing.get("postal_code") or "",
                        "listing_city": listing.get("city") or "",
                        "listing_url": listing.get("url") or "",
                        "matched_street": matched.get("street") or "",
                        "matched_house_number": matched.get("house_number") or "",
                        "matched_addition": matched.get("addition") or "",
                        "matched_postal_code": matched.get("postal_code") or "",
                        "matched_city": matched.get("city") or "",
                    }
                    append_found_match_to_csv(FOUND_MATCHES_CSV_FILE, match_row, FOUND_MATCH_CSV_FIELDS)
                    if found_ws is not None:
                        append_match_to_found_data_sheet(found_ws, match_row, FOUND_MATCH_CSV_FIELDS, found_seen_ids)
                    save_seen_db(SEEN_DB_FILE, seen_db)
                else:
                    logging.debug("Already seen targeted listing %s - skipping notify", listing["id"])

        save_seen_db(SEEN_DB_FILE, seen_db)
        save_seen_id_dict(MASTER_DATASET_SEEN_IDS_FILE, export_seen_ids)

    processed_targeted = len(targeted_listings)
    processed += processed_targeted
    logging.info("Run finished. Processed %d listings (%d from targeted searches). Matches found this run: %d. Error flag: %s", processed, processed_targeted, matches_found, error_flag)

    # Daily summary — always sent so you can verify the run completed.
    from datetime import datetime, timezone
    run_date = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    status_icon = "⚠️ ERRORS" if error_flag else "✅ OK"
    match_line = (
        f"🏠 Matches found: {matches_found}"
        if matches_found > 0
        else "🔍 No matches found today"
    )
    summary_msg = (
        f"📋 Funda Daily Summary — {run_date}\n"
        f"Status: {status_icon}\n"
        f"📄 Listings checked: {processed}\n"
        f"{match_line}\n"
        f"❌ Errors: {'Yes — check logs' if error_flag else 'None'}"
    )
    logging.info("Sending daily summary via Telegram.")
    send_telegram(summary_msg)
    return {"processed": processed, "matches": matches_found, "errors": error_flag}

if __name__ == "__main__":
    run_once()