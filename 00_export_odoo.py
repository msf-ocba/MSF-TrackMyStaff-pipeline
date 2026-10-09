"""
00_export_odoo.py
------------------
Logs into an Odoo instance, mirrors the browser's list-view lookup
(web_search_read on product.template), then exports the same records to
XLSX via /web/export/xlsx.

This script produces the reference files under templates/ that the
rest of the pipeline consumes (tms_article_list_path, location_list_path
and mission_code_prefixes_path in config.conf [templates]). It runs as
step 00 of run_pipeline.py, before 01_download_sftp.py, so every batch is
validated against fresh TMS article / location data.

Exit codes (see run_pipeline.py):
    0 = both reference workbooks were refreshed
    1 = any failure (login, CSRF, export, missing config/allowlist, ...)

Usage:
    python 00_export_odoo.py

Security notes:
  - Credentials are read from environment variables first (recommended),
    falling back to config.conf (which must be gitignored -- see
    config.conf.example for the template).
  - Nothing sensitive (password, csrf token, full cookie values) is ever
    printed or logged. Only non-sensitive diagnostics are shown.
  - All requests go over the URL you configure; make sure it's https.
  - Uses a fresh CSRF token scoped to the authenticated session for the
    export POST, rather than reusing the login-page token.
"""

import configparser
import json
import os
import re
import sys
from pathlib import Path

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo


class CsrfError(Exception):
    """Raised when Odoo rejects a request specifically due to an invalid/expired CSRF token."""
    pass


CONFIG_FILE = Path(__file__).with_name("config.conf")

# Same keys and same templates/ folder as the [templates] section the
# rest of the pipeline reads, so nothing needs renaming.
DEFAULT_TMS_ARTICLE_LIST_PATH = "./templates/TMS_UniDataArticles.xlsx"
DEFAULT_LOCATION_LIST_PATH = "./templates/location.xlsx"
DEFAULT_MISSION_CODE_PREFIXES_PATH = "./templates/Mission_code_prefixes.txt"

REQUEST_TIMEOUT = 30
EXPORT_TIMEOUT = 120

# Configuration

def load_config():
    """
    Resolve configuration with this precedence:
      1. Environment variables: ODOO_URL, ODOO_USERNAME, ODOO_PASSWORD,
         plus optional path overrides (see below)
      2. config.conf (local file, must be gitignored -- never commit it)

    Environment variables are preferred because they avoid ever writing
    a plaintext password to disk in the repo folder, and they're the
    standard way to inject secrets in CI/scheduled-task environments.

    The template output paths (tms_article_list_path, location_list_path,
    mission_code_prefixes_path) use the same names as the [templates]
    section in the pipeline's own config.conf, since these
    are the exact same files that the pipeline reads from templates/.
    """
    url = os.environ.get("ODOO_URL")
    username = os.environ.get("ODOO_USERNAME")
    password = os.environ.get("ODOO_PASSWORD")
    output_file = os.environ.get("TMS_ARTICLE_LIST_PATH")
    location_output_file = os.environ.get("LOCATION_LIST_PATH")
    mission_code_prefixes_path = os.environ.get("MISSION_CODE_PREFIXES_PATH")

    if url and username and password:
        print("Loaded credentials from environment variables.")
        return {
            "url": url.rstrip("/"),
            "username": username,
            "password": password,
            "output_file": output_file or DEFAULT_TMS_ARTICLE_LIST_PATH,
            "location_output_file": location_output_file or DEFAULT_LOCATION_LIST_PATH,
            "mission_code_prefixes_path": mission_code_prefixes_path or DEFAULT_MISSION_CODE_PREFIXES_PATH,
        }

    # Fall back to config.conf
    if not CONFIG_FILE.exists():
        print(f"ERROR: No environment variables set and config file not found: {CONFIG_FILE}")
        print("Either set ODOO_URL / ODOO_USERNAME / ODOO_PASSWORD as environment")
        print("variables, or copy config.conf.example to config.conf and fill it in.")
        sys.exit(1)

    config = configparser.ConfigParser()
    config.read(CONFIG_FILE)

    if "odoo" not in config:
        print("ERROR: [odoo] section missing from config.conf")
        sys.exit(1)

    for key in ("url", "username", "password"):
        if key not in config["odoo"] or not config["odoo"][key]:
            print(f"ERROR: '{key}' missing from [odoo] section in config.conf")
            sys.exit(1)

    print("Loaded credentials from config.conf.")

    templates = config["templates"] if "templates" in config else {}

    return {
        "url": config["odoo"]["url"].rstrip("/"),
        "username": config["odoo"]["username"],
        "password": config["odoo"]["password"],
        "output_file": templates.get("tms_article_list_path", DEFAULT_TMS_ARTICLE_LIST_PATH),
        "location_output_file": templates.get("location_list_path", DEFAULT_LOCATION_LIST_PATH),
        "mission_code_prefixes_path": templates.get(
            "mission_code_prefixes_path", DEFAULT_MISSION_CODE_PREFIXES_PATH
        ),
    }

# Helpers

def get_csrf_token(response):
    """
    Extract the CSRF token from an Odoo page.

    Two formats are supported:
      1. Older Odoo login pages render a hidden
         <input name="csrf_token" value="..."> field.
      2. Modern Odoo (17+) app-shell pages (e.g. /odoo) don't render
         that input -- instead the token is embedded in an inline
         <script> block as part of the session bootstrap, e.g.
         `odoo.session_info = {"csrf_token": "...", ...}`.
    """
    soup = BeautifulSoup(response.text, "html.parser")

    csrf_input = soup.find("input", {"name": "csrf_token"})

    if csrf_input and csrf_input.get("value"):
        return csrf_input["value"]

    # Fallback: pull it out of the embedded session info JS/JSON blob.
    # The key may or may not be quoted depending on Odoo version, e.g.
    # `"csrf_token": "..."` (JSON) or `csrf_token: "..."` (JS object
    # literal) -- match both.
    match = re.search(r'csrf_token"?\s*:\s*"([^"]+)"', response.text)

    if match:
        return match.group(1)

    return None


def new_session():
    session = requests.Session()

    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/153.0.0.0 Safari/537.36"
            ),
            "Accept": (
                "text/html,application/xhtml+xml,"
                "application/xml;q=0.9,"
                "image/avif,image/webp,image/apng,"
                "*/*;q=0.8"
            ),
        }
    )

    return session

# Step 1: Login

def login(session, base_url, username, password):
    print("Logging into Odoo...")

    login_url = f"{base_url}/web/login"

    response = session.get(login_url, allow_redirects=True, timeout=REQUEST_TIMEOUT)
    print(f"Login page HTTP status: {response.status_code}")
    response.raise_for_status()

    csrf_token = get_csrf_token(response)

    if not csrf_token:
        print("\nERROR: Could not find CSRF token on the Odoo login page.")
        sys.exit(1)

    print("CSRF token found.")

    login_data = {
        "csrf_token": csrf_token,
        "login": username,
        "password": password,
        "redirect": "",
    }

    response = session.post(login_url, data=login_data, allow_redirects=True, timeout=REQUEST_TIMEOUT)

    # Never log the password. login_data is not printed.
    print(f"Login response HTTP status: {response.status_code}")
    print(f"Final URL: {response.url}")
    print(f"Session established: {'session_id' in session.cookies.keys()}")

    response.raise_for_status()

    if "/web/login" in response.url:
        print("\nERROR: Odoo login appears to have failed (still on login page).")
        sys.exit(1)

    print("Login successful.")

    return response, csrf_token


# Step 2: web_search_read (mirrors opening the product list view)

def search_read_records(session, base_url, model, fields_spec, domain=None, page_size=200, context=None):
    """
    Generic paginated web_search_read caller -- the same JSON-RPC call
    the Odoo web client makes when opening a list view -- looping
    until every matching record has been retrieved.

    This is a type="json" controller, so unlike /web/export/xlsx it is
    authenticated purely via the session cookie and does not require a
    csrf_token in the payload (JSON-RPC requests can't be forged by a
    plain HTML form, which is what CSRF protection guards against).
    """
    url = f"{base_url}/web/dataset/call_kw/{model}/web_search_read"

    all_records = []
    offset = 0
    total = None

    while True:
        payload = {
            "jsonrpc": "2.0",
            "method": "call",
            "params": {
                "model": model,
                "method": "web_search_read",
                "args": [],
                "kwargs": {
                    "domain": domain or [],
                    "specification": fields_spec,
                    "offset": offset,
                    "limit": page_size,
                    "context": context
                    or {
                        "lang": "en_US",
                        "tz": "Africa/Nairobi",
                        "allowed_company_ids": [1],
                    },
                },
            },
        }

        response = session.post(
            url,
            json=payload,
            timeout=REQUEST_TIMEOUT,
            headers={"Content-Type": "application/json"},
        )

        response.raise_for_status()
        result = response.json()

        if "error" in result:
            print(f"\nERROR: Odoo returned a JSON-RPC error for web_search_read on {model}:")
            print(json.dumps(result["error"], indent=2)[:2000])
            sys.exit(1)

        page_records = result.get("result", {}).get("records", [])
        total = result.get("result", {}).get("length", total or len(page_records))

        all_records.extend(page_records)
        offset += page_size

        print(f"  Fetched {len(all_records)} of {total} records...")

        if not page_records or len(all_records) >= total:
            break

    print(f"web_search_read complete: {len(all_records)} of {total} records retrieved.")

    return all_records


def search_read_products(session, base_url):
    print("\nCalling web_search_read on product.template...")

    return search_read_records(
        session,
        base_url,
        model="product.template",
        fields_spec={
            "default_code": {},
            "name": {},
            "source_type": {},
            "categ_id": {},
        },
        domain=[],
    )


def search_read_all_locations(session, base_url):
    """
    Fetches every stock.location record (no domain filter) with the
    fields needed for a raw extraction: id (always returned by Odoo
    regardless of specification), name, complete_name, and usage
    (the location's type, e.g. internal/transit/view).
    """
    print("\nCalling web_search_read on stock.location (all locations)...")

    return search_read_records(
        session,
        base_url,
        model="stock.location",
        fields_spec={
            "name": {},
            "complete_name": {},
            "usage": {},
        },
        domain=[],
    )


def build_transit_in_rows(locations):
    """
    Replicates the Odoo "Locations" list view filtered to Stock
    Locations tagged VL_Transit_IN with "Full Location Name contains
    Coordination": complete_name must truly *end* with "VL_Transit_IN"
    (not just contain it mid-path) and contain "Coordination" somewhere
    in the path.

    Returns one row per match with both "Location Path" and "Full
    Location Name" set to complete_name, matching the Odoo view where
    both columns show the same value. This is the sheet
    the pipeline's location_list_path consumes.
    """
    rows = []

    for loc in locations:
        complete_name = loc.get("complete_name", "")

        # Case-insensitive "contains Coordination" to match Odoo's own
        # filter behavior (e.g. "NG111-Abuja coordination" -- lowercase
        # 'c' -- still matches in the Odoo UI).
        if complete_name.endswith("VL_Transit_IN") and "coordination" in complete_name.lower():
            rows.append({"location_path": complete_name, "full_location_name": complete_name})

    return rows


def load_mission_code_prefixes(path):
    """
    Reads the editable mission/project code allowlist: one 2-character
    prefix per line, '#' starts a comment, blank lines ignored.
    """
    path = Path(path)

    if not path.exists():
        print(f"ERROR: Mission code prefix file not found: {path}")
        sys.exit(1)

    prefixes = set()

    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()

        if not line or line.startswith("#"):
            continue

        prefixes.add(line.upper())

    print(f"Loaded {len(prefixes)} mission code prefix(es) from {path.name}: {sorted(prefixes)}")

    return prefixes


def _classify_location(complete_name):
    """
    Identifies whether a stock.location complete_name is one of the two
    kinds Sheet 3 cares about, and extracts its 2-character mission code.

    - ".../<CODE><digits>-<City>-Coordination"
        -> type "Coordination", code = last path segment's first 2 chars
    - ".../<CODE><digits>-<City>-Coordination/VL_Transit_IN"
        -> type "Transit In", code = second-to-last segment's first 2 chars

    Returns (location_type, mission_code) or (None, None) if the path
    doesn't match either pattern.
    """
    if complete_name.endswith("VL_Transit_IN"):
        parts = complete_name.split("/")

        if len(parts) < 2:
            return None, None

        code_segment = parts[-2]
        return "Transit In", code_segment[:2].upper()

    if complete_name.endswith("Coordination"):
        parts = complete_name.split("/")
        code_segment = parts[-1]
        return "Coordination", code_segment[:2].upper()

    return None, None


def build_deployed_locations(all_locations, allowed_prefixes):
    """
    Builds the Sheet 3 rows: every Coordination and Transit In location
    whose mission code appears in the allowlist.
    """
    deployed = []

    for loc in all_locations:
        complete_name = loc.get("complete_name", "")
        location_type, mission_code = _classify_location(complete_name)

        if location_type is None:
            continue

        if mission_code in allowed_prefixes:
            deployed.append(
                {
                    "complete_name": complete_name,
                    "type": location_type,
                    "mission_code": mission_code,
                }
            )

    return deployed

# Fresh CSRF token for the authenticated session

def _try_csrf_from_session_info(session, base_url):
    """Method A: POST /web/session/get_session_info (JSON-RPC)."""
    url = f"{base_url}/web/session/get_session_info"

    response = session.post(
        url,
        json={"jsonrpc": "2.0", "method": "call", "params": {}},
        timeout=REQUEST_TIMEOUT,
        headers={"Content-Type": "application/json"},
    )

    if response.status_code != 200:
        return None, {"method": "get_session_info", "status": response.status_code}

    try:
        result = response.json()
    except ValueError:
        return None, {"method": "get_session_info", "status": response.status_code, "note": "non-JSON response"}

    if "error" in result:
        return None, {"method": "get_session_info", "error": result["error"].get("message")}

    # Handle both a JSON-RPC-wrapped {"result": {...}} shape and a
    # plain unwrapped dict, depending on Odoo version.
    payload = result.get("result", result) if isinstance(result, dict) else {}

    if isinstance(payload, dict) and payload.get("csrf_token"):
        return payload["csrf_token"], None

    keys = list(payload.keys()) if isinstance(payload, dict) else []
    return None, {"method": "get_session_info", "status": 200, "keys_found": keys}


def _try_csrf_from_page(session, base_url, path):
    """Method B/C: GET an authenticated HTML page and scrape the token."""
    response = session.get(f"{base_url}{path}", allow_redirects=True, timeout=REQUEST_TIMEOUT)

    if response.status_code != 200:
        return None, {"method": f"scrape {path}", "status": response.status_code}

    token = get_csrf_token(response)

    if token:
        return token, None

    return None, {"method": f"scrape {path}", "status": 200, "final_url": response.url}


def _diagnose_csrf_presence(session, base_url, path):
    """
    Last-resort diagnostic: fetch a page and scan for any CSRF-related
    text so we can see the actual format this Odoo instance uses,
    without needing another guess. Token-like values are masked.
    """
    try:
        response = session.get(f"{base_url}{path}", allow_redirects=True, timeout=REQUEST_TIMEOUT)
    except requests.RequestException as exc:
        return [f"  {path}: request failed ({exc})"]

    text = response.text
    lines = []

    for match in re.finditer(r".{0,40}csrf.{0,60}", text, re.IGNORECASE):
        snippet = match.group(0)
        # Mask anything that looks like a long token value inside the snippet.
        snippet = re.sub(
            r"([A-Za-z0-9_\-]{16,})",
            lambda m: m.group(0)[:6] + "..." + m.group(0)[-4:],
            snippet,
        )
        lines.append(f"  {path}: ...{snippet}...")

    if not lines:
        lines.append(f"  {path}: no occurrence of 'csrf' found in page text at all")

    return lines[:5]  # cap output


def get_fresh_csrf_token(session, base_url):
    """
    The login-page CSRF token is scoped to that form submission and is
    not accepted afterwards. /web/export/xlsx is a type="http" POST
    controller and validates csrf_token on every request, so we need a
    token tied to the current logged-in session.

    Different Odoo versions/instances expose this differently, so we
    try a few known methods in order and use whichever succeeds first.
    """
    print("\nFetching a fresh CSRF token for the authenticated session...")

    attempts = []

    token, diag = _try_csrf_from_session_info(session, base_url)
    if token:
        print("Fresh CSRF token found (via get_session_info).")
        return token
    attempts.append(diag)

    token, diag = _try_csrf_from_page(session, base_url, "/odoo")
    if token:
        print("Fresh CSRF token found (via /odoo page).")
        return token
    attempts.append(diag)

    token, diag = _try_csrf_from_page(session, base_url, "/web")
    if token:
        print("Fresh CSRF token found (via /web page).")
        return token
    attempts.append(diag)

    print("\nERROR: Could not obtain a fresh CSRF token via any method.")
    print("Diagnostics from each attempt (no sensitive data included):")
    for a in attempts:
        print(f"  {a}")

    print("\nScanning raw page text for any CSRF-related markup (values masked):")
    for path in ("/odoo", "/web"):
        for line in _diagnose_csrf_presence(session, base_url, path):
            print(line)

    sys.exit(1)

# Step 3: Export to XLSX

def export_products(session, base_url, output_file, csrf_token):
    payload = {
        "import_compat": False,
        "context": {
            "lang": "en_US",
            "tz": "Africa/Nairobi",
            "allowed_company_ids": [1],
        },
        "domain": [],
        "fields": [
            {"name": "default_code", "label": "Article Code", "store": True, "type": "char"},
            {"name": "name", "label": "Article Description", "store": True, "type": "char"},
            {"name": "source_type", "label": "Source", "store": True, "type": "selection"},
            {"name": "categ_id", "label": "Category", "store": True, "type": "many2one"},
        ],
        "groupby": [],
        "ids": False,
        "model": "product.template",
    }

    print("Sending export request...")

    response = session.post(
        f"{base_url}/web/export/xlsx",
        files={
            "data": (None, json.dumps(payload)),
            "csrf_token": (None, csrf_token),
        },
        timeout=EXPORT_TIMEOUT,
    )

    print(f"Export HTTP status: {response.status_code}")
    print(f"Content-Type: {response.headers.get('Content-Type', 'unknown')}")
    print(f"Content-Length: {len(response.content):,} bytes")

    if response.status_code >= 400:
        body = response.text[:1000]

        if "csrf" in body.lower():
            print("\nExport rejected specifically due to CSRF token.")
            raise CsrfError(body)

        print("\nERROR: Odoo rejected the export request.")
        print("\nResponse:")
        print(body)
        response.raise_for_status()

    content_type = response.headers.get("Content-Type", "").lower()

    if "spreadsheetml" not in content_type and "application/vnd.ms-excel" not in content_type:
        print("\nWARNING: Response doesn't look like an XLSX file.")
        print(response.text[:1000])
        sys.exit(1)

    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_bytes(response.content)

    print("\nSUCCESS")
    print(f"Saved: {output_path.resolve()}")
    print(f"Size: {len(response.content):,} bytes")

# Post-processing: turn the exported range into a proper Excel Table

def add_excel_table(output_file, table_name="ProductsTable"):
    """
    Wraps the exported data range in a real Excel Table object (filter
    dropdowns on the header row, banded row styling) rather than
    leaving it as a plain data range.
    """
    print("\nAdding Excel table formatting...")

    wb = load_workbook(output_file)
    ws = wb.active

    max_row = ws.max_row
    max_col = ws.max_column

    if max_row < 2 or max_col < 1:
        print("WARNING: Sheet has no data rows; skipping table formatting.")
        return

    last_col_letter = get_column_letter(max_col)
    table_ref = f"A1:{last_col_letter}{max_row}"

    table = Table(displayName=table_name, ref=table_ref)
    table.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )

    ws.add_table(table)
    wb.save(output_file)

    print(f"Table '{table_name}' applied over range {table_ref}.")


def create_locations_workbook(output_file, all_locations, transit_in_rows, deployed_locations):
    """
    Builds a standalone workbook with three sheets:
      - "Transit In Locations": Location Path / Full Location Name,
        filtered to VL_Transit_IN locations whose path contains
        "Coordination" -- matches the Odoo "Locations" list view
        (Stock Locations + VL_Transit_IN + "contains Coordination").
        This is Sheet 1, the one the pipeline's location_list_path
        consumes (05_export.py reads the first sheet).
      - "All Locations": every stock.location record, as a plain
        (unstyled) Excel Table -- a raw, unformatted extraction kept
        for reference/audit.
      - "Deployed Locations": Coordination and Transit In locations
        whose mission code is in the Mission_code_prefixes.txt
        allowlist.
    """
    print(f"\nBuilding {output_file}...")

    wb = Workbook()

    # --- Sheet 1: Transit In Locations (Location Path / Full Location Name) ---
    ws1 = wb.active
    ws1.title = "Transit In Locations"

    headers1 = ["Location Path", "Full Location Name"]
    ws1.append(headers1)

    for row in transit_in_rows:
        ws1.append([row["location_path"], row["full_location_name"]])

    if ws1.max_row >= 2:
        table_ref1 = f"A1:{get_column_letter(len(headers1))}{ws1.max_row}"
        table1 = Table(displayName="TransitInLocationsTable", ref=table_ref1)
        table1.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws1.add_table(table1)
        ws1.column_dimensions["A"].width = 60
        ws1.column_dimensions["B"].width = 60
        print(f"  'Transit In Locations' table applied over range {table_ref1} ({len(transit_in_rows)} rows).")
    else:
        print("  WARNING: No Transit In locations matched; 'Transit In Locations' sheet has header only.")

    # --- Sheet 2: All Locations (raw, unformatted, kept for reference) ---
    ws2 = wb.create_sheet(title="All Locations")

    headers2 = ["ID", "Name", "Complete Name", "Usage"]
    ws2.append(headers2)

    for loc in all_locations:
        ws2.append(
            [
                loc.get("id"),
                loc.get("name"),
                loc.get("complete_name"),
                loc.get("usage"),
            ]
        )

    if ws2.max_row >= 2:
        table_ref2 = f"A1:{get_column_letter(len(headers2))}{ws2.max_row}"
        table2 = Table(displayName="AllLocationsTable", ref=table_ref2)
        # Deliberately unformatted: no row striping/banding, just a
        # recognized Excel Table so filter dropdowns are available.
        table2.tableStyleInfo = TableStyleInfo(
            name="TableStyleLight1",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=False,
            showColumnStripes=False,
        )
        ws2.add_table(table2)
        print(f"  'All Locations' table applied over range {table_ref2} ({len(all_locations)} rows).")
    else:
        print("  WARNING: No locations retrieved; 'All Locations' sheet has headers only.")

    # --- Sheet 3: Deployed Locations (final, filtered by mission code allowlist) ---
    ws3 = wb.create_sheet(title="Deployed Locations")
    headers3 = ["Location", "Type", "Mission Code"]
    ws3.append(headers3)

    for row in deployed_locations:
        ws3.append([row["complete_name"], row["type"], row["mission_code"]])

    if ws3.max_row >= 2:
        table_ref3 = f"A1:{get_column_letter(len(headers3))}{ws3.max_row}"
        table3 = Table(displayName="DeployedLocationsTable", ref=table_ref3)
        table3.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        ws3.add_table(table3)
        ws3.column_dimensions["A"].width = 70
        ws3.column_dimensions["B"].width = 14
        ws3.column_dimensions["C"].width = 14
        print(f"  'Deployed Locations' table applied over range {table_ref3} ({len(deployed_locations)} rows).")
    else:
        print("  WARNING: No deployed locations matched the allowlist; 'Deployed Locations' sheet has header only.")

    output_path = Path(output_file)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

    print(f"Saved: {output_path.resolve()}")

# Main

def main():
    config = load_config()

    session = new_session()

    try:
        login(session, config["url"], config["username"], config["password"])

        search_read_products(session, config["url"])

        # This Odoo instance rotates the session on login, which
        # invalidates the login-page CSRF token -- so we go straight
        # for a token scoped to the post-login session rather than
        # trying (and expectedly failing) with the login one first.
        csrf_token = get_fresh_csrf_token(session, config["url"])
        export_products(session, config["url"], config["output_file"], csrf_token)

        add_excel_table(config["output_file"])

        all_locations = search_read_all_locations(session, config["url"])
        transit_in_rows = build_transit_in_rows(all_locations)
        print(f"Matched {len(transit_in_rows)} Transit In location(s) (VL_Transit_IN, contains 'Coordination').")

        allowed_prefixes = load_mission_code_prefixes(config["mission_code_prefixes_path"])
        deployed_locations = build_deployed_locations(all_locations, allowed_prefixes)
        print(f"Matched {len(deployed_locations)} deployed location(s) against the mission code allowlist.")

        create_locations_workbook(
            config["location_output_file"],
            all_locations,
            transit_in_rows,
            deployed_locations,
        )
    finally:
        # Drop credentials from memory as soon as we're done with them.
        config["password"] = None
        session.close()


if __name__ == "__main__":
    main()
