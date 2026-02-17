import azure.functions as func
import csv
import json
from pathlib import Path
from io import StringIO
import base64
import os
import logging
from typing import Optional
import binascii
from io import BytesIO
import zipfile
from datetime import date, datetime, timedelta, timezone
import re
from pypdf import PdfReader, PdfWriter
from PIL import Image

import requests
from azure.identity import DefaultAzureCredential
from azure.storage.blob import BlobServiceClient, ContentSettings
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest

# Default app auth level can stay FUNCTION if you want Function-key auth.
# If you instead want Entra/EasyAuth (Managed Identity), set individual routes to
# ANONYMOUS so Function-key auth does not apply and enforce auth at the platform layer.
app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

APP_VERSION = "2026-01-07-receipts-upload-v1"

CSV_PATH = Path(__file__).with_name("expense_codes.csv")
_ROWS = None
_DEPT_ENTRIES = None
_DEPT_EMAIL_OVERRIDES = None


def _get_first_str(obj: dict, *keys: str, default: str = "") -> str:
    """
    Extract the first non-empty string value from a dict using multiple possible key names.
    Handles case-insensitive key matching and returns the default if no value is found.
    """
    if not obj or not isinstance(obj, dict):
        return default
    # Build a lowercase key map for case-insensitive lookup
    lower_map = {str(k).lower(): v for k, v in obj.items()}
    for key in keys:
        # Try exact match first
        if key in obj:
            val = obj[key]
            if val is not None:
                s = str(val).strip()
                if s:
                    return s
        # Try case-insensitive match
        lower_key = str(key).lower()
        if lower_key in lower_map:
            val = lower_map[lower_key]
            if val is not None:
                s = str(val).strip()
                if s:
                    return s
    return default


def _today_in_configured_tz() -> date:
    """
    Returns "today" using TRAVEL_TIMEZONE so relative phrases like "last Tuesday"
    match the user's expectation (Mountain Time), not UTC.

    Default: America/Denver (Mountain Time, DST-aware).
    If TRAVEL_TIMEZONE is set to "MST" (or UTC-7), uses a fixed offset.
    """

    tz_name = (os.getenv("TRAVEL_TIMEZONE") or "America/Denver").strip()
    tz = None

    # Prefer IANA timezones (DST-aware).
    try:
        from zoneinfo import ZoneInfo  # py3.9+

        tz = ZoneInfo(tz_name)
    except Exception:
        tz = None

    # Fallback: fixed-offset MST.
    if tz is None:
        if tz_name.upper() in ("MST", "UTC-7", "UTC-07", "UTC-07:00", "GMT-7", "GMT-07:00"):
            tz = timezone(timedelta(hours=-7))
        else:
            tz = timezone.utc

    return datetime.now(timezone.utc).astimezone(tz).date()


def _split_code_and_desc(value: str):
    """Turns '620 - INFORMATION TECHNOLOGY' into ('620', 'INFORMATION TECHNOLOGY')."""
    if not value:
        return "", ""
    parts = str(value).split(" - ", 1)
    code = parts[0].strip()
    desc = parts[1].strip() if len(parts) > 1 else ""
    return code, desc


def _load_rows():
    """
    Expected CSV headers (from your xlsx export) are likely:
      activity, account, department
    where each value looks like '700 - BUSINESS TRAVEL', '561 - LOAD DISPATCHING', etc.
    """
    rows = []
    if not CSV_PATH.exists():
        return rows

    with CSV_PATH.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for r in reader:
            dept_code, _ = _split_code_and_desc(r.get("department", ""))
            act_code, _ = _split_code_and_desc(r.get("activity", ""))
            acct_code, acct_desc = _split_code_and_desc(r.get("account", ""))

            if dept_code and act_code and acct_code:
                rows.append(
                    {
                        "departmentCode": dept_code,
                        "activityCode": act_code,
                        "accountCode": acct_code,
                        "description": acct_desc,
                    }
                )
    return rows


def _normalize_key(value: str) -> str:
    if not value:
        return ""
    text = str(value).upper().strip()
    text = text.replace("&", " AND ")
    text = re.sub(r"[^A-Z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _load_department_entries():
    """
    Loads unique department codes/names from the expense_codes.csv file.
    Each entry: {"departmentCode": "620", "departmentName": "INFORMATION TECHNOLOGY", "norm": "...", "tokens": set(...)}
    """
    global _DEPT_ENTRIES
    if _DEPT_ENTRIES is not None:
        return _DEPT_ENTRIES

    entries = []
    seen_codes = set()
    if not CSV_PATH.exists():
        _DEPT_ENTRIES = entries
        return _DEPT_ENTRIES

    with CSV_PATH.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for r in reader:
            dept_code, dept_name = _split_code_and_desc(r.get("department", ""))
            if not dept_code or not dept_name or dept_code in seen_codes:
                continue
            seen_codes.add(dept_code)
            norm = _normalize_key(dept_name)
            entries.append(
                {
                    "departmentCode": dept_code,
                    "departmentName": dept_name,
                    "norm": norm,
                    "tokens": set(norm.split()),
                }
            )

    _DEPT_ENTRIES = entries
    return _DEPT_ENTRIES


def _load_dept_email_overrides() -> dict:
    """
    Optional production-safe escape hatch for OrgChart data that does not uniquely
    identify a department code (e.g., multiple cost centers share the same orgchart
    department name).

    Configure with app setting ORGCHART_DEPT_EMAIL_OVERRIDES as JSON:
      {"user@core.coop":"175", "other@core.coop":"180"}
    """
    global _DEPT_EMAIL_OVERRIDES
    if _DEPT_EMAIL_OVERRIDES is not None:
        return _DEPT_EMAIL_OVERRIDES

    raw = (os.getenv("ORGCHART_DEPT_EMAIL_OVERRIDES") or "").strip()
    overrides = {}
    if raw:
        try:
            obj = json.loads(raw)
            if isinstance(obj, dict):
                for k, v in obj.items():
                    email = str(k or "").strip().lower()
                    code = str(v or "").strip()
                    if email and code:
                        overrides[email] = code
        except Exception:
            # If parsing fails, treat as "no overrides" rather than failing requests.
            overrides = {}

    _DEPT_EMAIL_OVERRIDES = overrides
    return _DEPT_EMAIL_OVERRIDES


def _department_name_for_code(department_code: str) -> str:
    department_code = str(department_code or "").strip()
    if not department_code:
        return ""
    for e in _load_department_entries():
        if e.get("departmentCode") == department_code:
            return str(e.get("departmentName") or "").strip()
    return ""


def _map_department_name_to_code(department_name: str):
    """
    Returns (departmentCode, canonicalDepartmentName, matchType, candidates)
    where candidates is a list of {"departmentCode","departmentName"}.

    matchType values: exact | fuzzy | none
    """
    if not department_name:
        return "", "", "none", []

    # If orgchart already returns "620 - INFORMATION TECHNOLOGY", honor it.
    code_hint, name_hint = _split_code_and_desc(department_name)
    if code_hint and code_hint.isdigit():
        return code_hint, name_hint or department_name, "exact", []

    query_norm = _normalize_key(department_name)
    if query_norm.isdigit():
        return query_norm, "", "exact", []

    entries = _load_department_entries()

    exact = [e for e in entries if e["norm"] == query_norm]
    if len(exact) == 1:
        e = exact[0]
        return e["departmentCode"], e["departmentName"], "exact", []

    query_tokens = set(query_norm.split())
    if not query_tokens:
        return "", "", "none", []

    scored = []
    for e in entries:
        common = len(query_tokens & e["tokens"])
        if common == 0:
            continue
        score = common / len(query_tokens)
        scored.append((score, e))

    scored.sort(key=lambda t: t[0], reverse=True)
    candidates = [
        {"departmentCode": e["departmentCode"], "departmentName": e["departmentName"]}
        for _, e in scored[:5]
    ]

    if not scored:
        return "", "", "none", []

    best_score, best = scored[0]
    second_score = scored[1][0] if len(scored) > 1 else 0

    # Auto-pick only when the match is strong and clearly better than the next option.
    if best_score >= 0.9 and best_score > second_score:
        return best["departmentCode"], best["departmentName"], "fuzzy", candidates

    return "", "", "none", candidates


def _orgchart_search_by_email(email: str, debug: bool = False) -> tuple[Optional[dict], Optional[str], list]:
    endpoint = (os.getenv("ORGCHART_SEARCH_ENDPOINT") or "").strip().rstrip("/")
    index_name = (os.getenv("ORGCHART_SEARCH_INDEX") or "").strip()
    api_key = (os.getenv("ORGCHART_SEARCH_API_KEY") or "").strip()
    api_version = (os.getenv("ORGCHART_SEARCH_API_VERSION") or "2023-11-01").strip()
    email_field = (os.getenv("ORGCHART_SEARCH_EMAIL_FIELD") or "email").strip()

    if endpoint and not endpoint.lower().startswith("http"):
        endpoint = f"https://{endpoint}.search.windows.net"

    if not endpoint or not index_name or not api_key:
        return None, "OrgChart search is not configured (missing ORGCHART_SEARCH_ENDPOINT/INDEX/API_KEY).", []

    url = f"{endpoint}/indexes/{index_name}/docs/search?api-version={api_version}"
    headers = {"api-key": api_key, "Content-Type": "application/json"}
    attempts: list[dict] = []

    def _truncate(s: str) -> str:
        s = (s or "").strip()
        return s if len(s) <= 500 else (s[:500] + "...")

    def _doc_email(d: dict) -> str:
        # Some indexes store the email in a title-like field.
        title = d.get("title")
        if title:
            t = str(title).strip().lower()
            if re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", t):
                return t

        # Some indexes store the email in a parent id/path (e.g. "KBOC/user@core.coop.json").
        parent_id = d.get("parent_id") or d.get("parentId") or ""
        if parent_id:
            m = re.search(r"([^/\\]+@[^/\\]+\.[^/\\]+)\.json", str(parent_id).strip(), flags=re.IGNORECASE)
            if m:
                return m.group(1).strip().lower()

        for k in [email_field, "email", "upn", "userPrincipalName", "mail", "Email", "UPN", "Mail"]:
            v = d.get(k)
            if v:
                return str(v).strip().lower()

        # Some indexes store the original user JSON as a string in a field like "chunk".
        chunk_raw = d.get("chunk")
        if isinstance(chunk_raw, str) and chunk_raw.lstrip().startswith("{"):
            try:
                chunk_obj = json.loads(chunk_raw)
            except Exception:
                chunk_obj = {}
            for k in ["UPN", "Mail", "Email", "email", "upn", "userPrincipalName", "mail"]:
                v = chunk_obj.get(k)
                if v:
                    return str(v).strip().lower()

        return ""

    # Strict (case-insensitive) match.
    email_l = (email or "").strip().lower()

    # Try filter-first (fast + precise) but many indexes don't have the email field marked filterable
    # (or they don't store email at the top-level). If filtering yields no exact match, fall back to
    # plain search and enforce strict matching in code.
    filter_body = {"search": "*", "filter": f"tolower({email_field}) eq '{email_l}'", "top": 5}
    try:
        resp = requests.post(url, headers=headers, json=filter_body, timeout=15)
    except Exception as e:
        return None, f"OrgChart search request failed: {e}", attempts

    docs = []
    if resp.status_code == 200:
        try:
            data = resp.json()
        except Exception:
            return None, "OrgChart search returned invalid JSON.", attempts

        docs = data.get("value") or []
        if debug:
            attempts.append(
                {"mode": "filter", "status": resp.status_code, "docsCount": len(docs), "exactCount": 0}
            )
        exact = [d for d in docs if isinstance(d, dict) and _doc_email(d) == email_l]
        if debug:
            attempts[-1]["exactCount"] = len(exact)
        if len(exact) == 1:
            return exact[0], None, attempts
        if len(exact) > 1:
            return None, "OrgChart search returned multiple exact matches for this email.", attempts
        # No exact match via filter; fall back to plain search.
    else:
        # If filter isn't supported, fall back to plain search.
        docs = []
        if debug:
            attempts.append(
                {"mode": "filter", "status": resp.status_code, "docsCount": 0, "exactCount": 0}
            )

    def _search_then_exact(search_text: str, top: int) -> tuple[Optional[dict], Optional[str], list]:
        # Some indexes are very picky about which fields are searchable. We try restricting searchFields
        # to the likely carriers of the email (and fall back if the service rejects searchFields).
        search_fields = "title,parent_id,chunk,chunk_id"
        body = {
            "search": search_text,
            "top": top,
            "queryType": "simple",
            "searchMode": "all",
            "searchFields": search_fields,
        }
        try:
            r = requests.post(url, headers=headers, json=body, timeout=15)
        except Exception as e:
            return None, f"OrgChart search request failed: {e}", []

        if r.status_code == 400:
            # If the index rejects searchFields (fields not searchable/unknown), retry without it.
            body.pop("searchFields", None)
            try:
                r = requests.post(url, headers=headers, json=body, timeout=15)
            except Exception as e:
                return None, f"OrgChart search request failed: {e}", []

        if r.status_code != 200:
            return (
                None,
                f"OrgChart search failed (HTTP {r.status_code}). {_truncate(r.text)}".strip(),
                [],
            )

        try:
            data = r.json()
        except Exception:
            return None, "OrgChart search returned invalid JSON.", []

        docs_local = data.get("value") or []
        if debug:
            attempts.append(
                {
                    "mode": "search",
                    "status": r.status_code,
                    "searchText": search_text,
                    "docsCount": len(docs_local),
                    "exactCount": 0,
                }
            )
        exact_local = [d for d in docs_local if isinstance(d, dict) and _doc_email(d) == email_l]
        if debug:
            attempts[-1]["exactCount"] = len(exact_local)
        if len(exact_local) == 1:
            return exact_local[0], None, docs_local
        if len(exact_local) > 1:
            return None, "OrgChart search returned multiple exact matches for this email.", docs_local
        return None, None, docs_local

    # Fallback 1: search for the full email string.
    doc, err, docs = _search_then_exact(email_l, top=10)
    if err:
        return None, err, attempts
    if doc:
        return doc, None, attempts

    # Fallback 2: search for just the local-part. Many analyzers strip punctuation like '@' which can
    # make exact-email searches return zero hits even though the doc contains the email.
    local_part = email_l.split("@", 1)[0].strip()
    if local_part:
        doc, err, docs2 = _search_then_exact(local_part, top=50)
        if err:
            return None, err, attempts
        if doc:
            return doc, None, attempts
        # Keep docs from the broader search for diagnostics.
        docs = docs2 or docs

    # Fallback 3: search for a normalized chunk id (common in SharePoint-index chunking pipelines)
    # Example: hjacobsen@core.coop -> hjacobsen_core_coop_0
    norm = re.sub(r"[^a-z0-9]+", "_", email_l).strip("_")
    if norm:
        for candidate in (f"{norm}_0", norm):
            doc, err, docs3 = _search_then_exact(candidate, top=50)
            if err:
                return None, err, attempts
            if doc:
                return doc, None, attempts
            docs = docs3 or docs

    # If the search did return some docs but none matched exactly, surface details.
    if docs:
        return (
            None,
            f"OrgChart search returned results but none matched email exactly. Configure ORGCHART_SEARCH_EMAIL_FIELD (currently '{email_field}').",
            attempts,
        )

    return None, None, attempts


def _zip_to_place(zip_code: str) -> tuple[Optional[str], Optional[str], Optional[str]]:
    """Returns (stateAbbr, city, error). Uses a public ZIP lookup service."""
    base = (os.getenv("ZIP_GEOCODE_BASE_URL") or "https://api.zippopotam.us/us").strip().rstrip("/")
    url = f"{base}/{zip_code}"
    try:
        resp = requests.get(url, timeout=10)
    except Exception as e:
        return None, None, f"ZIP geocode request failed: {e}"

    if resp.status_code != 200:
        return None, None, f"ZIP geocode failed (HTTP {resp.status_code})."

    try:
        data = resp.json()
    except Exception:
        return None, None, "ZIP geocode returned invalid JSON."

    places = data.get("places") or []
    if not places or not isinstance(places[0], dict):
        return None, None, "ZIP geocode returned no places."

    place = places[0]
    state = (place.get("state abbreviation") or place.get("state") or "").strip()
    city = (place.get("place name") or place.get("place") or "").strip()
    if not state or not city:
        return None, None, "ZIP geocode response missing state/city."

    return state, city, None


def _parse_iso_date(value: str) -> Optional[date]:
    s = str(value or "").strip()
    if not s:
        return None

    s_norm = s.strip()
    s_lower = s_norm.lower()
    today = _today_in_configured_tz()

    # Relative keywords.
    if s_lower in ("today",):
        return today
    if s_lower in ("yesterday",):
        return today - timedelta(days=1)
    if s_lower in ("tomorrow",):
        return today + timedelta(days=1)

    # Relative weekday phrases: "last tuesday", "next fri", "this monday".
    m = re.fullmatch(r"(last|next|this)\s+([a-z]+)", s_lower)
    if m:
        rel = m.group(1)
        wd_raw = m.group(2)
        weekday_map = {
            "mon": 0,
            "monday": 0,
            "tue": 1,
            "tues": 1,
            "tuesday": 1,
            "wed": 2,
            "wednesday": 2,
            "thu": 3,
            "thur": 3,
            "thurs": 3,
            "thursday": 3,
            "fri": 4,
            "friday": 4,
            "sat": 5,
            "saturday": 5,
            "sun": 6,
            "sunday": 6,
        }
        if wd_raw in weekday_map:
            target = weekday_map[wd_raw]
            current = today.weekday()
            if rel == "last":
                delta = (current - target) % 7
                if delta == 0:
                    delta = 7
                return today - timedelta(days=delta)
            if rel == "next":
                delta = (target - current) % 7
                if delta == 0:
                    delta = 7
                return today + timedelta(days=delta)
            if rel == "this":
                delta = (target - current) % 7
                return today + timedelta(days=delta)

    # Try strict ISO first.
    try:
        return date.fromisoformat(s)
    except Exception:
        pass

    # Accept common US formats often produced by chat inputs.
    for fmt in (
        "%m/%d/%Y",
        "%m-%d-%Y",
        "%m %d %Y",
        "%m/%d/%y",
        "%m-%d-%y",
        "%m %d %y",
        "%b %d %Y",
        "%B %d %Y",
        "%b %d, %Y",
        "%B %d, %Y",
        "%b %d %y",
        "%B %d %y",
        "%b %d, %y",
        "%B %d, %y",
    ):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass

    # Month name without year: "Dec 12" / "December 12"
    m = re.fullmatch(r"([a-z]+)\s+(\d{1,2})(?:st|nd|rd|th)?", s_lower)
    if m:
        month_raw = m.group(1)
        day_raw = m.group(2)
        month_map = {
            "jan": 1,
            "january": 1,
            "feb": 2,
            "february": 2,
            "mar": 3,
            "march": 3,
            "apr": 4,
            "april": 4,
            "may": 5,
            "jun": 6,
            "june": 6,
            "jul": 7,
            "july": 7,
            "aug": 8,
            "august": 8,
            "sep": 9,
            "sept": 9,
            "september": 9,
            "oct": 10,
            "october": 10,
            "nov": 11,
            "november": 11,
            "dec": 12,
            "december": 12,
        }
        if month_raw in month_map:
            month = month_map[month_raw]
            day = int(day_raw)
            year = today.year
            try:
                candidate = date(year, month, day)
                # Default to the most recent past date if the user omitted year.
                if candidate > today:
                    candidate = date(year - 1, month, day)
                return candidate
            except Exception:
                pass

    # Last resort: extract digits and attempt MMDDYYYY or YYYYMMDD.
    digits = re.sub(r"\D", "", s)
    if len(digits) >= 8:
        # Prefer YYYYMMDD if it looks like it starts with a year.
        y = int(digits[:4])
        if 1900 <= y <= 2100:
            try:
                return date(int(digits[:4]), int(digits[4:6]), int(digits[6:8]))
            except Exception:
                pass
        # Otherwise try MMDDYYYY.
        try:
            return date(int(digits[4:8]), int(digits[0:2]), int(digits[2:4]))
        except Exception:
            pass

    return None


def _extract_first_number(value) -> Optional[float]:
    try:
        if value is None:
            return None
        if isinstance(value, (int, float)):
            return float(value)
        s = str(value).strip().replace("$", "").replace(",", "")
        if s == "":
            return None
        return float(s)
    except Exception:
        return None


def _find_mie_rate(obj) -> Optional[float]:
    """
    Attempts to find a Meals & Incidental (M&IE) daily rate in a loosely-typed JSON response.
    Looks for common keys: mie, mealsAndIncidental, meals, incidental (meals+incidental).
    """

    if obj is None:
        return None

    if isinstance(obj, dict):
        # direct keys
        for key in ["mie", "MIE", "mealsAndIncidental", "meals_incidental", "m_and_ie"]:
            if key in obj:
                n = _extract_first_number(obj.get(key))
                if n is not None:
                    return n
        meals = None
        incidental = None
        for key in ["meals", "Meals"]:
            if key in obj:
                meals = _extract_first_number(obj.get(key))
        for key in ["incidental", "Incidentals", "incidentals"]:
            if key in obj:
                incidental = _extract_first_number(obj.get(key))

        # Most GSA responses use a single "meals" field for the daily M&IE rate.
        if meals is not None and incidental is None:
            return float(meals)

        if meals is not None and incidental is not None:
            return float(meals) + float(incidental)

        # recurse
        for v in obj.values():
            n = _find_mie_rate(v)
            if n is not None:
                return n

    if isinstance(obj, list):
        for item in obj:
            n = _find_mie_rate(item)
            if n is not None:
                return n

    return None


_US_STATES = {
    "AL": "ALABAMA", "AK": "ALASKA", "AZ": "ARIZONA", "AR": "ARKANSAS", "CA": "CALIFORNIA",
    "CO": "COLORADO", "CT": "CONNECTICUT", "DE": "DELAWARE", "FL": "FLORIDA", "GA": "GEORGIA",
    "HI": "HAWAII", "ID": "IDAHO", "IL": "ILLINOIS", "IN": "INDIANA", "IA": "IOWA",
    "KS": "KANSAS", "KY": "KENTUCKY", "LA": "LOUISIANA", "ME": "MAINE", "MD": "MARYLAND",
    "MA": "MASSACHUSETTS", "MI": "MICHIGAN", "MN": "MINNESOTA", "MS": "MISSISSIPPI",
    "MO": "MISSOURI", "MT": "MONTANA", "NE": "NEBRASKA", "NV": "NEVADA", "NH": "NEW HAMPSHIRE",
    "NJ": "NEW JERSEY", "NM": "NEW MEXICO", "NY": "NEW YORK", "NC": "NORTH CAROLINA",
    "ND": "NORTH DAKOTA", "OH": "OHIO", "OK": "OKLAHOMA", "OR": "OREGON", "PA": "PENNSYLVANIA",
    "RI": "RHODE ISLAND", "SC": "SOUTH CAROLINA", "SD": "SOUTH DAKOTA", "TN": "TENNESSEE",
    "TX": "TEXAS", "UT": "UTAH", "VT": "VERMONT", "VA": "VIRGINIA", "WA": "WASHINGTON",
    "WV": "WEST VIRGINIA", "WI": "WISCONSIN", "WY": "WYOMING", "DC": "DISTRICT OF COLUMBIA",
}
_STATE_NAME_TO_ABBR = {v: k for k, v in _US_STATES.items()}


def _parse_city_state(raw: str) -> tuple[Optional[str], Optional[str]]:
    """Parse 'City, ST' or 'City ST' or 'City, State Name' into (city, state_abbr)."""
    s = raw.strip()
    if not s:
        return None, None

    # Try "City, ST" or "City, State Name"
    parts = None
    if "," in s:
        parts = [p.strip() for p in s.split(",", 1)]
    else:
        # Try splitting on last whitespace-separated token as state
        tokens = s.rsplit(None, 1)
        if len(tokens) == 2:
            parts = tokens

    if not parts or len(parts) < 2 or not parts[0] or not parts[1]:
        return None, None

    city = parts[0].strip()
    state_raw = parts[1].strip().upper()

    # Direct abbreviation match
    if state_raw in _US_STATES:
        return city, state_raw

    # Full state name match
    abbr = _STATE_NAME_TO_ABBR.get(state_raw)
    if abbr:
        return city, abbr

    return None, None


def _gsa_per_diem_city_state_lookup(city: str, state: str, travel_date: Optional[date], debug: bool = False) -> tuple[Optional[dict], Optional[str]]:
    """Look up GSA per diem rate by city and state abbreviation."""
    api_key = (os.getenv("GSA_API_KEY") or "").strip()
    if not api_key:
        return None, "GSA_API_KEY is not configured."

    base_url = (os.getenv("GSA_PER_DIEM_BASE_URL") or "https://api.gsa.gov/travel/perdiem/v2").strip().rstrip("/")
    dt = travel_date or _today_in_configured_tz()
    fiscal_year = dt.year + 1 if dt.month >= 10 else dt.year

    headers = {"Accept": "application/json", "x-api-key": api_key}
    params = {"api_key": api_key}

    city_norm = city.upper().replace(".", " ").replace("'", " ").replace("-", " ")
    city_norm = " ".join(city_norm.split())
    city_enc = city_norm.replace(" ", "%20")
    city_url = f"{base_url}/rates/city/{city_enc}/state/{state.upper()}/year/{fiscal_year}"

    attempts = []
    try:
        resp = requests.get(city_url, headers=headers, params=params, timeout=15)
    except Exception as e:
        return None, f"GSA per diem request failed: {e}"

    if debug:
        attempts.append({"url": city_url, "status": resp.status_code})

    if resp.status_code == 200:
        try:
            body = resp.json()
            if debug:
                body = {"_debug": {"attempts": attempts, "city": city, "state": state}, "data": body}
            return body, None
        except Exception:
            return None, "GSA per diem returned invalid JSON."

    return None, f"GSA per diem lookup failed for {city}, {state} (HTTP {resp.status_code})."


def _gsa_per_diem_lookup(zip_code: str, travel_date: Optional[date], debug: bool = False) -> tuple[Optional[dict], Optional[str]]:
    api_key = (os.getenv("GSA_API_KEY") or "").strip()
    if not api_key:
        return None, "GSA_API_KEY is not configured."

    base_url = (os.getenv("GSA_PER_DIEM_BASE_URL") or "https://api.gsa.gov/travel/perdiem/v2").strip().rstrip("/")

    dt = travel_date or _today_in_configured_tz()
    # Federal fiscal year starts Oct 1.
    fiscal_year = dt.year + 1 if dt.month >= 10 else dt.year

    headers = {"Accept": "application/json", "x-api-key": api_key}
    params = {"api_key": api_key}

    attempts = []

    def _try(url: str):
        try:
            resp = requests.get(url, headers=headers, params=params, timeout=15)
        except Exception as e:
            return None, None, f"GSA per diem request failed: {e}"

        content_type = (resp.headers.get("content-type") or "").split(";")[0].strip().lower()
        if debug:
            attempts.append({"url": url, "status": resp.status_code, "contentType": content_type})

        if resp.status_code == 200:
            try:
                return resp.json(), resp.status_code, None
            except Exception:
                return None, resp.status_code, "GSA per diem returned invalid JSON."

        return (resp.text or ""), resp.status_code, None

    # Official v2 ZIP endpoint:
    #   GET /rates/zip/{zip}/year/{year}
    zip_url = f"{base_url}/rates/zip/{zip_code}/year/{fiscal_year}"
    body, status, err = _try(zip_url)
    if err:
        return None, err
    if status == 200:
        if debug:
            body = {"_debug": {"attempts": attempts}, "data": body}
        return body, None

    # Fallback: city/state endpoint using ZIP -> city/state.
    state, city, geo_err = _zip_to_place(zip_code)
    if geo_err is None and state and city:
        city_norm = city.upper().replace(".", " ").replace("'", " ").replace("-", " ")
        city_norm = " ".join(city_norm.split())
        city_enc = city_norm.replace(" ", "%20")
        city_url = f"{base_url}/rates/city/{city_enc}/state/{state.upper()}/year/{fiscal_year}"
        body2, status2, err2 = _try(city_url)
        if err2:
            return None, err2
        if status2 == 200:
            if debug:
                body2 = {"_debug": {"attempts": attempts, "zipCity": city, "zipState": state}, "data": body2}
            return body2, None

    # If we got here, we failed.
    if debug and attempts:
        return None, f"GSA per diem lookup failed. Attempts: {attempts}"

    # Prefer the ZIP failure status.
    return None, f"GSA per diem lookup failed (HTTP {status})."


@app.route(route="orgchart-lookup", methods=["GET"], auth_level=func.AuthLevel.FUNCTION)
def orgchart_lookup(req: func.HttpRequest) -> func.HttpResponse:
    email = (req.params.get("email") or req.params.get("upn") or "").strip()
    debug = str(req.params.get("debug") or "").strip().lower() in ("1", "true", "yes")
    if not email:
        return func.HttpResponse(
            json.dumps({"ok": False, "found": False, "error": "email is required"}),
            mimetype="application/json",
        )

    doc, err, attempts = _orgchart_search_by_email(email, debug=debug)
    if err:
        payload = {"ok": False, "found": False, "email": email.lower().strip(), "error": err}
        if debug:
            payload["debug"] = {"attempts": attempts}
        return func.HttpResponse(
            # Always return 200 so Copilot Studio connector actions don't hard-fail the topic on non-2xx.
            json.dumps(payload),
            mimetype="application/json",
        )

    if not doc:
        payload = {"ok": True, "found": False, "email": email.lower().strip()}
        if debug:
            payload["debug"] = {"attempts": attempts}
        return func.HttpResponse(
            json.dumps(payload),
            mimetype="application/json",
        )
    # Extract fields with tolerant key names. Do not return phone numbers by default.
    chunk_obj = {}
    try:
        chunk_raw = doc.get("chunk")
        if isinstance(chunk_raw, str) and chunk_raw.lstrip().startswith("{"):
            chunk_obj = json.loads(chunk_raw)
    except Exception:
        chunk_obj = {}

    resolved_email = (
        doc.get("email")
        or doc.get("upn")
        or doc.get("userPrincipalName")
        or doc.get("mail")
        or doc.get("Email")
        or doc.get("UPN")
        or doc.get("Mail")
        or chunk_obj.get("email")
        or chunk_obj.get("upn")
        or chunk_obj.get("userPrincipalName")
        or chunk_obj.get("mail")
        or chunk_obj.get("Email")
        or chunk_obj.get("UPN")
        or chunk_obj.get("Mail")
        or email
    )
    display_name = (doc.get("displayName") or doc.get("name") or doc.get("fullName") or chunk_obj.get("DisplayName") or chunk_obj.get("displayName") or "")
    job_title = (doc.get("jobTitle") or doc.get("title") or chunk_obj.get("JobTitle") or chunk_obj.get("jobTitle") or "")
    department_name = (doc.get("department") or doc.get("Department") or chunk_obj.get("Department") or chunk_obj.get("department") or "")

    # If OrgChart doesn't provide a unique department code, allow an email-level override.
    # This is intentionally optional and non-fatal if misconfigured.
    dept_override_used = False
    dept_override_code = ""
    resolved_email_lc = str(resolved_email).strip().lower()
    overrides = _load_dept_email_overrides()
    if resolved_email_lc and overrides and resolved_email_lc in overrides:
        dept_override_used = True
        dept_override_code = overrides.get(resolved_email_lc) or ""
        dept_code = str(dept_override_code).strip()
        dept_name_canonical = _department_name_for_code(dept_code) or str(department_name).strip()
        match_type = "exact"
        candidates = []
    else:
        dept_code, dept_name_canonical, match_type, candidates = _map_department_name_to_code(department_name)

    payload = {
        "ok": True,
        "found": True,
        "email": str(resolved_email).strip().lower(),
        "displayName": str(display_name).strip(),
        "jobTitle": str(job_title).strip(),
        "departmentName": str(department_name).strip(),
        "departmentCode": dept_code,
        "departmentNameMapped": dept_name_canonical,
        "departmentMatchType": match_type,
        "departmentCandidates": candidates,
    }
    if debug:
        debug_obj = {"attempts": attempts}
        if dept_override_used:
            debug_obj["deptOverride"] = {"email": resolved_email_lc, "departmentCode": dept_override_code}
        payload["debug"] = debug_obj
    return func.HttpResponse(json.dumps(payload), mimetype="application/json")


@app.route(route="orgchart-lookup-upn", methods=["GET"], auth_level=func.AuthLevel.FUNCTION)
def orgchart_lookup_upn(req: func.HttpRequest) -> func.HttpResponse:
    """
    Identical to orgchart-lookup, but requires the query parameter name `upn`.
    This exists to work around Copilot Studio tool/input binding issues where an
    existing required parameter name can get stuck as an incompatible "custom value" type.
    """
    upn = (req.params.get("upn") or req.params.get("email") or "").strip()
    debug = str(req.params.get("debug") or "").strip().lower() in ("1", "true", "yes")
    if not upn:
        return func.HttpResponse(
            json.dumps({"ok": False, "found": False, "error": "upn is required"}),
            mimetype="application/json",
        )

    doc, err, attempts = _orgchart_search_by_email(upn, debug=debug)
    if err:
        payload = {"ok": False, "found": False, "email": upn.lower().strip(), "error": err}
        if debug:
            payload["debug"] = {"attempts": attempts}
        return func.HttpResponse(json.dumps(payload), mimetype="application/json")

    if not doc:
        payload = {"ok": True, "found": False, "email": upn.lower().strip()}
        if debug:
            payload["debug"] = {"attempts": attempts}
        return func.HttpResponse(json.dumps(payload), mimetype="application/json")

    chunk_obj = {}
    try:
        chunk_raw = doc.get("chunk")
        if isinstance(chunk_raw, str) and chunk_raw.lstrip().startswith("{"):
            chunk_obj = json.loads(chunk_raw)
    except Exception:
        chunk_obj = {}

    display_name = _get_first_str(chunk_obj, "DisplayName", "displayName", default="")
    job_title = _get_first_str(chunk_obj, "JobTitle", "jobTitle", default="")
    department_name = _get_first_str(chunk_obj, "Department", "department", default="")
    dept_override_used = False
    dept_override_code = ""
    upn_lc = str(upn).strip().lower()
    overrides = _load_dept_email_overrides()
    if upn_lc and overrides and upn_lc in overrides:
        dept_override_used = True
        dept_override_code = overrides.get(upn_lc) or ""
        department_code = str(dept_override_code).strip()
        department_name_mapped = _department_name_for_code(department_code) or str(department_name).strip()
        match_type = "exact"
        candidates = []
    else:
        department_code, department_name_mapped, match_type, candidates = _map_department_name_to_code(department_name)

    payload = {
        "ok": True,
        "found": True,
        "email": upn.lower().strip(),
        "displayName": display_name,
        "jobTitle": job_title,
        "departmentName": department_name,
        "departmentCode": department_code,
        "departmentNameMapped": department_name_mapped,
        "departmentMatchType": match_type,
        "departmentCandidates": candidates,
    }
    if debug:
        debug_obj = {"attempts": attempts}
        if dept_override_used:
            debug_obj["deptOverride"] = {"email": upn_lc, "departmentCode": dept_override_code}
        payload["debug"] = debug_obj
    return func.HttpResponse(json.dumps(payload), mimetype="application/json")


@app.route(route="per-diem-lookup", methods=["GET"], auth_level=func.AuthLevel.FUNCTION)
def per_diem_lookup(req: func.HttpRequest) -> func.HttpResponse:
    params = req.params or {}

    def _param_ci(*names: str) -> str:
        if not params:
            return ""
        lower = {str(k).lower(): v for k, v in params.items()}
        for name in names:
            v = lower.get(str(name).lower())
            if v is None:
                continue
            s = str(v).strip()
            if s:
                return s
        return ""

    # Power Platform connectors sometimes change casing or send inputs in the body even for GET operations.
    body = {}
    try:
        body_json = req.get_json()
        if isinstance(body_json, dict):
            body = body_json
    except Exception:
        body = {}

    def _body_ci(*names: str) -> str:
        if not body:
            return ""
        lower = {str(k).lower(): v for k, v in body.items()}
        for name in names:
            v = lower.get(str(name).lower())
            if v is None:
                continue
            s = str(v).strip()
            if s:
                return s
        return ""

    raw_body_text = ""
    try:
        raw_body_text = (req.get_body() or b"").decode("utf-8", errors="ignore").strip()
    except Exception:
        raw_body_text = ""

    zip_code_raw = (
        _param_ci("zipCode", "zip", "zipcode", "zip_code", "location")
        or _body_ci("zipCode", "zip", "zipcode", "zip_code", "location")
        or raw_body_text
        or ""
    ).strip()

    travel_date = _parse_iso_date(
        (_param_ci("travelDate", "date", "travel_date") or _body_ci("travelDate", "date", "travel_date")).strip()
    )

    debug = str(req.params.get('debug') or req.params.get('includeDebug') or '').strip().lower() in ('1','true','yes')

    # Try to extract a 5-digit ZIP first.
    zip_digits = re.sub(r"\D", "", zip_code_raw or "")
    zip_code = (zip_digits[:5] if zip_digits else "").strip()

    raw = None
    err = None

    if re.fullmatch(r"\d{5}", zip_code or ""):
        # Standard ZIP lookup
        raw, err = _gsa_per_diem_lookup(zip_code, travel_date, debug=debug)
    else:
        # Try to parse as "City, State" or "City State"
        city_raw = zip_code_raw.strip()
        if not city_raw:
            return func.HttpResponse(
                json.dumps({"ok": False, "error": "Provide a 5-digit ZIP or City, State (e.g. Denver, CO)."}),
                mimetype="application/json",
            )
        city_name, state_abbr = _parse_city_state(city_raw)
        if not city_name or not state_abbr:
            return func.HttpResponse(
                json.dumps({"ok": False, "error": f"Could not parse location '{city_raw}'. Use a 5-digit ZIP or City, State (e.g. Denver, CO)."}),
                mimetype="application/json",
            )
        raw, err = _gsa_per_diem_city_state_lookup(city_name, state_abbr, travel_date, debug=debug)
    if err:
        return func.HttpResponse(
            json.dumps({"ok": False, "error": err}),
            mimetype="application/json",
        )

    raw_data = raw.get('data') if isinstance(raw, dict) and 'data' in raw else raw
    mie = _find_mie_rate(raw_data)
    if mie is None:
        return func.HttpResponse(
            json.dumps({"ok": False, "error": "Unable to extract an M&IE rate from the per diem response."}),
            mimetype="application/json",
        )

    return func.HttpResponse(
        json.dumps(
            {
                "ok": True,
                "zipCode": zip_code,
                "debug": (raw.get('_debug') if isinstance(raw, dict) and '_debug' in raw else None),
                "travelDate": travel_date.isoformat() if travel_date else None,
                "mieRate": mie,
            }
        ),
        mimetype="application/json",
    )


@app.route(route="health", methods=["GET"], auth_level=func.AuthLevel.FUNCTION)
def health(req: func.HttpRequest) -> func.HttpResponse:
    return func.HttpResponse(json.dumps({"ok": True}), mimetype="application/json")


@app.route(route="expense-codes", methods=["GET"], auth_level=func.AuthLevel.FUNCTION)
def expense_codes(req: func.HttpRequest) -> func.HttpResponse:
    global _ROWS
    if _ROWS is None:
        _ROWS = _load_rows()

    dept = (req.params.get("departmentCode") or "").strip()
    act = (req.params.get("activityCode") or "").strip()

    if not dept or not act:
        return func.HttpResponse(
            json.dumps({"error": "departmentCode and activityCode are required"}),
            status_code=400,
            mimetype="application/json",
        )

    matches = [
        {"accountCode": r["accountCode"], "description": r.get("description", "")}
        for r in _ROWS
        if r["departmentCode"] == dept and r["activityCode"] == act
    ]

    # de-dupe
    seen = set()
    uniq = []
    for m in matches:
        if m["accountCode"] in seen:
            continue
        seen.add(m["accountCode"])
        uniq.append(m)

    return func.HttpResponse(
        json.dumps({"departmentCode": dept, "activityCode": act, "matches": uniq}),
        mimetype="application/json",
    )


_IMPORTFORMAT_FIELDS = [
    "GL Division",  # 1
    "GL Department",  # 2
    "GL Account",  # 3
    "GL Activity",  # 4
    "Reference",  # 5
    "Amount",  # 6
    "Vendor",  # 7
    "Organization Name",  # 8
    "First Name",  # 9
    "Last Name",  # 10
    "Joint Name",  # 11
    "Address Line 1",  # 12
    "Address Line 2",  # 13
    "Address Line 3",  # 14
    "City",  # 15
    "State",  # 16
    "ZIP",  # 17
    "Bank Account",  # 18
    "Due Date",  # 19
    "Invoice",  # 20
    "Customer",  # 21
    "Payment Type",  # 22
    "1099",  # 23
    "Type",  # 24
    "Invoice Date",  # 25
    "GL Post Date",  # 26
    "Discount",  # 27
    "Sales Tax",  # 28
    "Additional Charge 1 Code",  # 29
    "Additional Charge 1 Amount",  # 30
    "Addl Chg 1 Taxable",  # 31
    "Additional Charge 2 Code",  # 32
    "Additional Charge 2 Amount",  # 33
    "Addl Chg 2 Taxable",  # 34
    "Use Tax 1 Code",  # 35
    "Use Tax 1 Amount",  # 36
    "Use Tax 2 Code",  # 37
    "Use Tax 2 Amount",  # 38
    "Use Tax 3 Code",  # 39
    "Use Tax 3 Amount",  # 40
    "Use Tax 4 Code",  # 41
    "Use Tax 4 Amount",  # 42
    "Taxable",  # 43
    "Apply Addl Chg 1",  # 44
    "Apply Addl Chg 2",  # 45
    "Distribute Tax",  # 46
    "Distribute Addl Chg 1",  # 47
    "Distribute Addl Chg 2",  # 48
    "AP GL Division",  # 49
    "AP GL Account",  # 50
    "Dispute",  # 51
    "Separate Payment",  # 52
    "Times To Post",  # 53
    "Invoice Type",  # 54
    "Extended Reference",  # 55
    "Authorization Type",  # 56
    "Credit Card",  # 57
    "Paid Vendor",  # 58
    "Charge Dt",  # 59
    "BU Project",  # 60
    "GL Distribution Reference",  # 61
]


def _fmt_amount(value) -> str:
    try:
        return f"{float(value):.2f}"
    except Exception:
        return ""


def _coalesce(*values) -> str:
    for v in values:
        if v is None:
            continue
        s = str(v)
        if s.strip() == "":
            continue
        return s
    return ""


def _make_reference(base_reference: str, original_account_code: str, gl_override: str) -> str:
    ref = (base_reference or "").strip()
    if gl_override and original_account_code:
        suffix = f" ACCT={original_account_code}".strip()
        if len(ref) + len(suffix) <= 40:
            ref = (ref + suffix).strip()
        else:
            keep = max(0, 40 - len(suffix))
            ref = (ref[:keep].rstrip() + suffix)[:40].strip()
    return ref[:40]


_INVOICE_PREFIX_MAP = {
    "receipt": "EXP",
    "boots": "EXP",
    "perdiem": "PER DIEM",
    "mileage": "MIL",
}


def _invoice_number(item_type: str, today: date) -> str:
    """Build invoice number like 'EXP 02-2026' from item type and today's date."""
    key = (item_type or "").replace(" ", "").strip().lower()
    prefix = _INVOICE_PREFIX_MAP.get(key, "EXP")
    return f"{prefix} {today.strftime('%m')}-{today.strftime('%Y')}"


def _iter_import_rows(payload: dict):
    division = _coalesce(payload.get("division"), "0000")
    vendor = _coalesce(payload.get("vendor"), "CORE")

    today = date.today()
    due = today + timedelta(days=7)
    # Use MM/DD/YYYY (common for import CSVs).
    today_s = today.strftime("%m/%d/%Y")
    due_s = due.strftime("%m/%d/%Y")

    requester = payload.get("requester") or {}
    org_name = _coalesce(requester.get("organizationName"))
    first_name = _coalesce(requester.get("firstName"))
    last_name = _coalesce(requester.get("lastName"))

    items = payload.get("items")
    if items is None and isinstance(payload.get("draftItemsJson"), str):
        items = json.loads(payload["draftItemsJson"])
    if not isinstance(items, list):
        items = []

    for item in items:
        if not isinstance(item, dict):
            continue

        dept = _coalesce(item.get("departmentCode"))
        base_reference = _coalesce(item.get("reference"))
        gl_override = _coalesce(item.get("glAccountOverride"))
        item_type = _coalesce(item.get("type"), "Receipt")
        invoice = _invoice_number(item_type, today)

        lines = item.get("lines")
        if not isinstance(lines, list) or len(lines) == 0:
            lines = [
                {
                    "amount": item.get("amountTotal") or item.get("amount"),
                    "activityCode": item.get("activityCode"),
                    "accountCode": item.get("accountCode"),
                }
            ]

        for line in lines:
            if not isinstance(line, dict):
                continue

            activity = _coalesce(line.get("activityCode"), item.get("activityCode"))
            original_account_code = _coalesce(line.get("accountCode"), item.get("accountCode"))
            gl_account = gl_override or original_account_code

            reference = _make_reference(base_reference, original_account_code, gl_override)
            amount = _fmt_amount(line.get("amount"))

            extended_ref = ""
            if gl_override and original_account_code:
                extended_ref = _coalesce(
                    item.get("description"),
                    f"OVERRIDE={gl_override}; ORIGINAL_ACCT={original_account_code}",
                )

            row = {field: "" for field in _IMPORTFORMAT_FIELDS}
            row["GL Division"] = division
            row["GL Department"] = dept
            row["GL Account"] = gl_account
            row["GL Activity"] = activity
            # Column 5 (Reference) is a category label, not the original reference.
            row["Reference"] = "TRAINING/EDUCATION" if str(activity).strip() == "770" else "EXPENSES / MILEAGE"
            row["Amount"] = amount
            row["Vendor"] = vendor
            row["Organization Name"] = org_name
            row["First Name"] = first_name
            row["Last Name"] = last_name
            row["Address Line 1"] = "."
            # Dates required by import (positions 19, 25, 26)
            row["Due Date"] = due_s
            row["Invoice"] = invoice
            row["Invoice Date"] = today_s
            row["GL Post Date"] = today_s
            # Put the prior per-line reference into notes (Extended Reference, position 55).
            # Preserve any override notes by appending.
            if extended_ref:
                row["Extended Reference"] = f"{reference} | {extended_ref}"
            else:
                row["Extended Reference"] = reference

            yield row


@app.route(route="import-csv", methods=["POST"], auth_level=func.AuthLevel.FUNCTION)
def import_csv(req: func.HttpRequest) -> func.HttpResponse:
    try:
        payload = req.get_json()
    except Exception:
        return func.HttpResponse(
            json.dumps({"error": "Invalid JSON body"}),
            status_code=400,
            mimetype="application/json",
        )

    output = StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=_IMPORTFORMAT_FIELDS, extrasaction="ignore")

    for row in _iter_import_rows(payload):
        writer.writerow(row)

    return func.HttpResponse(output.getvalue(), mimetype="text/csv")


def _graph_send_mail(
    *,
    from_user: str,
    to_email: str,
    cc_emails: Optional[list[str]] = None,
    subject: str,
    body_text: str,
    body_html: Optional[str] = None,
    csv_text: str,
    csv_filename: str = "travel-expense.csv",
    additional_attachments: Optional[list[dict]] = None,
) -> Optional[str]:
    try:
        token = _graph_access_token()
    except Exception as e:
        return f"Failed to acquire Graph token via DefaultAzureCredential: {e}"

    if (os.getenv("LOG_GRAPH_TOKEN_CLAIMS") or "").strip().lower() in {"1", "true", "yes", "y"}:
        try:
            # JWT format: header.payload.signature. We only decode payload for debugging.
            parts = token.split(".")
            if len(parts) >= 2:
                payload_b64 = parts[1]
                pad = "=" * ((4 - (len(payload_b64) % 4)) % 4)
                payload_json = base64.urlsafe_b64decode((payload_b64 + pad).encode("utf-8")).decode("utf-8")
                claims = json.loads(payload_json)
                logging.info(
                    "Graph token claims: aud=%s roles=%s scp=%s appid=%s oid=%s tid=%s",
                    claims.get("aud"),
                    claims.get("roles"),
                    claims.get("scp"),
                    claims.get("appid"),
                    claims.get("oid"),
                    claims.get("tid"),
                )
            else:
                logging.warning("Graph token is not in expected JWT format; cannot decode claims.")
        except (ValueError, binascii.Error, UnicodeDecodeError) as e:
            logging.warning("Failed to decode Graph token claims: %s", e)

    url = f"https://graph.microsoft.com/v1.0/users/{from_user}/sendMail"
    csv_b64 = base64.b64encode(csv_text.encode("utf-8")).decode("ascii")
    attachments = [
        {
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": csv_filename,
            "contentType": "text/csv",
            "contentBytes": csv_b64,
        }
    ]
    if additional_attachments:
        attachments.extend(additional_attachments)

    body_content_type = "Text"
    body_content = body_text
    if isinstance(body_html, str) and body_html.strip() != "":
        body_content_type = "HTML"
        body_content = body_html

    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": body_content_type, "content": body_content},
            "toRecipients": [{"emailAddress": {"address": to_email}}],
            "attachments": attachments,
        },
        "saveToSentItems": "true",
    }

    cc_emails = [e.strip() for e in (cc_emails or []) if (e or "").strip()]
    if len(cc_emails) > 0:
        payload["message"]["ccRecipients"] = [{"emailAddress": {"address": e}} for e in cc_emails]

    resp = requests.post(
        url,
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        data=json.dumps(payload),
        timeout=30,
    )
    logging.info("Graph sendMail request complete: status=%s", resp.status_code)
    if resp.status_code not in (202, 200):
        logging.warning("Graph sendMail failed response (truncated): %s", (resp.text or "")[:2000])
        return f"Graph sendMail failed: {resp.status_code} {resp.text}"
    return None


def _graph_access_token() -> str:
    return DefaultAzureCredential().get_token("https://graph.microsoft.com/.default").token


def _graph_request(method: str, url: str, *, headers: Optional[dict] = None, json_body=None, timeout_s: int = 60) -> requests.Response:
    hdrs = {"Authorization": f"Bearer {_graph_access_token()}"}
    if headers:
        hdrs.update(headers)
    return requests.request(method, url, headers=hdrs, json=json_body, timeout=timeout_s)


def _graph_get_json(url: str, *, timeout_s: int = 60) -> dict:
    resp = _graph_request("GET", url, headers={"Accept": "application/json"}, timeout_s=timeout_s)
    if resp.status_code not in (200, 201, 202):
        raise RuntimeError(f"Graph GET failed: url={url} status={resp.status_code} {(resp.text or '')[:500]}")
    return resp.json()


def _graph_get_bytes(url: str, *, timeout_s: int = 120) -> bytes:
    resp = _graph_request("GET", url, headers={"Accept": "application/octet-stream"}, timeout_s=timeout_s)
    if resp.status_code not in (200, 201, 202):
        raise RuntimeError(f"Graph GET (bytes) failed: url={url} status={resp.status_code} {(resp.text or '')[:500]}")
    return resp.content or b""


def _graph_delete(url: str, *, timeout_s: int = 60) -> None:
    resp = _graph_request("DELETE", url, headers={"Accept": "application/json"}, timeout_s=timeout_s)
    if resp.status_code not in (200, 202, 204):
        raise RuntimeError(f"Graph DELETE failed: url={url} status={resp.status_code} {(resp.text or '')[:500]}")


def _parse_sharepoint_item_ids(payload: dict) -> list[str]:
    return _parse_csvish(
        payload.get("sharepointItemIds")
        or payload.get("sharePointItemIds")
        or payload.get("spItemIds")
        or payload.get("receiptSharepointItemIds")
        or payload.get("receiptSharePointItemIds")
        or payload.get("receiptItemIds")
    )


def _parse_sharepoint_urls(payload: dict) -> list[str]:
    return _parse_csvish(
        payload.get("sharepointFileUrls")
        or payload.get("sharePointFileUrls")
        or payload.get("sharepointUrls")
        or payload.get("sharePointUrls")
        or payload.get("receiptSharepointUrls")
        or payload.get("receiptSharePointUrls")
        or payload.get("receiptUrls")
    )


def _graph_share_id(url: str) -> str:
    # Graph expects a base64url-encoded URL prefixed with "u!"
    raw = url.encode("utf-8")
    b64 = base64.urlsafe_b64encode(raw).decode("ascii").rstrip("=")
    return f"u!{b64}"


def _download_receipts_from_sharepoint(payload: dict) -> tuple[list[bytes], list[str], Optional[str]]:
    """
    Downloads files from SharePoint/OneDrive via Microsoft Graph.

    Supported payload forms:
      A) driveId + itemIds:
        - sharepointDriveId
        - sharepointItemIds (list or comma-separated string of driveItem ids)
      B) share URLs (best for tools that only return webUrl):
        - sharepointFileUrls (list or comma-separated)
    """
    drive_id = _coalesce(payload.get("sharepointDriveId"), payload.get("sharePointDriveId"), payload.get("spDriveId"))
    item_ids = _parse_sharepoint_item_ids(payload)
    urls = _parse_sharepoint_urls(payload)

    downloaded: list[bytes] = []
    filenames: list[str] = []
    try:
        # Form B: resolve share URLs -> driveId/itemId
        if not (drive_id and item_ids) and urls:
            for i, u in enumerate(urls):
                sid = _graph_share_id(u)
                meta = _graph_get_json(
                    f"https://graph.microsoft.com/v1.0/shares/{sid}/driveItem?$select=id,name,parentReference",
                    timeout_s=30,
                )
                item_id = meta.get("id")
                pref = meta.get("parentReference") if isinstance(meta, dict) else {}
                drive_id = (pref.get("driveId") if isinstance(pref, dict) else None) or drive_id
                if item_id and drive_id:
                    item_ids.append(str(item_id))

        if not drive_id or not item_ids:
            return [], [], "sharepointDriveId+sharepointItemIds or sharepointFileUrls are required to fetch receipts from SharePoint"

        for i, item_id in enumerate(item_ids):
            meta = _graph_get_json(
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}?$select=id,name,file",
                timeout_s=30,
            )
            name = (meta.get("name") or "").strip() or f"receipt-{i+1}"
            content = _graph_get_bytes(
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
                timeout_s=120,
            )
            if not content:
                continue
            sniff_ext, _sniff_type = _sniff_file_type(content)
            if "." not in name and sniff_ext:
                name = f"{name}.{sniff_ext}"
            downloaded.append(content)
            filenames.append(name)

        if not downloaded:
            return [], [], "SharePoint receipts were empty or could not be downloaded"
        return downloaded, filenames, None
    except Exception as e:
        return [], [], f"Failed to download receipts from SharePoint: {e}"


def _purge_sharepoint_items(payload: dict) -> Optional[str]:
    """
    Best-effort deletion of temp items after successful send.
    Supports:
      - sharepointDriveId+sharepointItemIds
      - sharepointFileUrls (resolved to driveItem and then deleted)
    """
    drive_id = _coalesce(payload.get("sharepointDriveId"), payload.get("sharePointDriveId"), payload.get("spDriveId"))
    item_ids = _parse_sharepoint_item_ids(payload)
    urls = _parse_sharepoint_urls(payload)
    try:
        if urls and not (drive_id and item_ids):
            for u in urls:
                sid = _graph_share_id(u)
                meta = _graph_get_json(
                    f"https://graph.microsoft.com/v1.0/shares/{sid}/driveItem?$select=id,parentReference",
                    timeout_s=30,
                )
                item_id = meta.get("id")
                pref = meta.get("parentReference") if isinstance(meta, dict) else {}
                drive_id = (pref.get("driveId") if isinstance(pref, dict) else None) or drive_id
                if item_id and drive_id:
                    item_ids.append(str(item_id))
        if not drive_id or not item_ids:
            return None
        for item_id in item_ids:
            _graph_delete(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}")
        return None
    except Exception as e:
        return f"Failed to purge SharePoint receipts: {e}"


def _decode_b64(data_b64: str) -> bytes:
    raw = (data_b64 or "").strip()
    if len(raw) >= 2 and raw[0] in {"'", '"'} and raw[-1] == raw[0]:
        raw = raw[1:-1].strip()
    if raw.startswith("data:") and "," in raw:
        raw = raw.split(",", 1)[1]
    # Be tolerant of missing padding and stray whitespace from clients.
    raw = "".join(raw.split())
    raw = raw.replace("-", "+").replace("_", "/")
    missing = len(raw) % 4
    if missing:
        raw += "=" * (4 - missing)
    return base64.b64decode(raw, validate=True)


def _coerce_bytes(value) -> bytes:
    if value is None:
        return b""
    if isinstance(value, (bytes, bytearray)):
        return bytes(value)
    if isinstance(value, list) and all(isinstance(x, int) for x in value):
        return bytes(value)
    if isinstance(value, dict) and isinstance(value.get("data"), list) and all(
        isinstance(x, int) for x in value.get("data", [])
    ):
        return bytes(value["data"])
    if isinstance(value, str):
        raw = value.strip()
        if raw.startswith("[") and raw.endswith("]"):
            try:
                parsed = json.loads(raw)
            except Exception:
                parsed = None
            if isinstance(parsed, list) and all(isinstance(x, int) for x in parsed):
                return bytes(parsed)
        return _decode_b64(raw)
    return b""


def _sniff_file_type(data: bytes) -> tuple[Optional[str], Optional[str]]:
    if data.startswith(b"%PDF-"):
        return "pdf", "application/pdf"
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "png", "image/png"
    if data.startswith(b"\xff\xd8\xff"):
        return "jpg", "image/jpeg"
    return None, None


def _build_summary_table_pdf(payload: dict) -> Optional[bytes]:
    """
    Renders a summary table of expense items as a PDF page using PIL.
    Returns PDF bytes, or None if no items are available.
    """
    items = payload.get("items")
    if items is None and isinstance(payload.get("draftItemsJson"), str):
        try:
            items = json.loads(payload["draftItemsJson"])
        except Exception:
            items = None
    if not isinstance(items, list) or len(items) == 0:
        return None

    try:
        from PIL import ImageDraw, ImageFont
    except ImportError:
        logging.warning("PIL ImageDraw not available for summary table")
        return None

    try:
        # Table data
        headers = ["Type", "Date", "Description", "Account", "Dept", "Activity", "Amount"]
        rows = []
        total = 0.0
        for it in items:
            if not isinstance(it, dict):
                continue
            item_type = str(it.get("type") or "").strip()
            date_val = str(it.get("travelDate") or it.get("receiptDate") or "").strip()
            ref = str(it.get("reference") or "").strip()
            dept = str(it.get("departmentCode") or "").strip()
            activity = str(it.get("activityCode") or "").strip()
            account = str(it.get("accountCode") or "").strip()
            try:
                amt = float(it.get("amount", 0))
            except (ValueError, TypeError):
                amt = 0.0
            total += amt
            rows.append([item_type, date_val, ref[:40], account, dept, activity, f"${amt:.2f}"])

        if not rows:
            return None

        requester = str(payload.get("requesterEmail") or payload.get("toEmail") or "").strip()

        # Rendering settings
        dpi = 150
        page_w = int(8.5 * dpi)
        page_h = int(11 * dpi)
        margin = int(0.5 * dpi)

        # Use default font
        try:
            font = ImageFont.truetype("arial.ttf", 14)
            font_bold = ImageFont.truetype("arialbd.ttf", 14)
            font_header = ImageFont.truetype("arialbd.ttf", 18)
        except Exception:
            font = ImageFont.load_default()
            font_bold = font
            font_header = font

        # Column widths (proportional to page width)
        usable_w = page_w - 2 * margin
        col_ratios = [0.08, 0.12, 0.28, 0.10, 0.10, 0.10, 0.12]
        total_ratio = sum(col_ratios)
        col_widths = [int(r / total_ratio * usable_w) for r in col_ratios]
        # Adjust last column to fill remaining space
        col_widths[-1] = usable_w - sum(col_widths[:-1])

        row_height = 22
        header_area = 60

        img = Image.new("RGB", (page_w, page_h), "white")
        draw = ImageDraw.Draw(img)

        y = margin

        # Title and metadata
        draw.text((margin, y), "Travel Expense Summary", fill="black", font=font_header)
        y += 28
        if requester:
            draw.text((margin, y), f"Requester: {requester}", fill="black", font=font)
            y += 20
        from datetime import date as _date_type
        draw.text((margin, y), f"Date: {_date_type.today().isoformat()}", fill="black", font=font)
        y += 20
        draw.text((margin, y), f"Items: {len(rows)}    Total: ${total:.2f}", fill="black", font=font)
        y += 30

        # Draw table header
        x = margin
        for i, h in enumerate(headers):
            draw.rectangle([x, y, x + col_widths[i], y + row_height], fill="#4472C4", outline="black")
            draw.text((x + 4, y + 3), h, fill="white", font=font_bold)
            x += col_widths[i]
        y += row_height

        # Draw rows
        for row_idx, row in enumerate(rows):
            bg = "#F2F2F2" if row_idx % 2 == 0 else "white"
            x = margin
            for i, cell in enumerate(row):
                draw.rectangle([x, y, x + col_widths[i], y + row_height], fill=bg, outline="#CCCCCC")
                # Truncate text to fit column
                text = str(cell)
                draw.text((x + 4, y + 3), text, fill="black", font=font)
                x += col_widths[i]
            y += row_height

        # Draw total row
        x = margin
        total_label_w = sum(col_widths[:-1])
        draw.rectangle([x, y, x + total_label_w, y + row_height], fill="#D9E2F3", outline="black")
        draw.text((x + total_label_w - 60, y + 3), "Total:", fill="black", font=font_bold)
        x += total_label_w
        draw.rectangle([x, y, x + col_widths[-1], y + row_height], fill="#D9E2F3", outline="black")
        draw.text((x + 4, y + 3), f"${total:.2f}", fill="black", font=font_bold)

        # Convert to PDF
        out = BytesIO()
        img.save(out, format="PDF", resolution=dpi)
        return out.getvalue()
    except Exception as e:
        logging.warning("Failed to build summary table PDF: %s", e)
        return None


def _merge_pdf_bytes(pdf_blobs: list[bytes]) -> bytes:
    writer = PdfWriter()
    for blob in pdf_blobs:
        reader = PdfReader(BytesIO(blob))
        for page in reader.pages:
            writer.add_page(page)
    out = BytesIO()
    writer.write(out)
    return out.getvalue()


def _bytes_to_pdf(blob: bytes) -> tuple[bytes, Optional[str]]:
    """
    Returns (pdf_bytes, error). Supports PDFs and common image formats.
    Resizes large images to keep PDF attachments under email size limits.
    """
    _ext, ctype = _sniff_file_type(blob)
    if ctype == "application/pdf":
        return blob, None
    if ctype in {"image/png", "image/jpeg"}:
        try:
            img = Image.open(BytesIO(blob))

            if getattr(img, "mode", None) not in {"RGB"}:
                img = img.convert("RGB")

            # Resize large images to max 2000px on longest side (keeps receipts readable)
            max_dim = 2000
            if img.width > max_dim or img.height > max_dim:
                ratio = min(max_dim / img.width, max_dim / img.height)
                new_w = int(img.width * ratio)
                new_h = int(img.height * ratio)
                img = img.resize((new_w, new_h), Image.LANCZOS)

            out = BytesIO()
            img.save(out, format="PDF", resolution=150, quality=90)
            pdf_bytes = out.getvalue()
            if not pdf_bytes.startswith(b"%PDF-"):
                return b"", "image-to-pdf conversion did not produce a PDF"
            return pdf_bytes, None
        except Exception as e:
            return b"", f"image-to-pdf conversion failed: {e}"
    return b"", "unsupported receipt type (only PDF/JPG/PNG supported)"


def _receipt_bundle_format(payload: dict) -> str:
    """
    Returns 'pdf' or 'zip'.

    Back-compat:
    - receiptBundleFormat='pdf' => pdf
    - receiptBundleFormat='zip' => zip
    """
    fmt = (payload.get("receiptBundleFormat") or payload.get("bundleFormat") or "").strip().lower()
    if fmt in {"pdf", "receipts.pdf"}:
        return "pdf"
    if fmt in {"zip", "receipts.zip"}:
        return "zip"
    # New default: PDF (single attachment).
    return "pdf"

def _foundry_get_access_token() -> str:
    """
    Azure AI Foundry project endpoints use the https://ai.azure.com/.default scope.
    """
    return DefaultAzureCredential().get_token("https://ai.azure.com/.default").token


def _blob_service_client() -> BlobServiceClient:
    """
    Uses either a connection string (preferred for simplicity) or Managed Identity.
    Configure one of:
      - RECEIPTS_STORAGE_CONNECTION_STRING
      - RECEIPTS_STORAGE_ACCOUNT_URL (e.g., https://<acct>.blob.core.windows.net)
    """
    conn = (os.getenv("RECEIPTS_STORAGE_CONNECTION_STRING") or "").strip()
    if conn:
        return BlobServiceClient.from_connection_string(conn)
    account_url = (os.getenv("RECEIPTS_STORAGE_ACCOUNT_URL") or "").strip()
    if not account_url:
        raise RuntimeError("RECEIPTS_STORAGE_ACCOUNT_URL (or RECEIPTS_STORAGE_CONNECTION_STRING) is not configured")
    return BlobServiceClient(account_url=account_url, credential=DefaultAzureCredential())


def _receipt_container_name() -> str:
    return (os.getenv("RECEIPTS_CONTAINER") or "travel-expense-receipts").strip() or "travel-expense-receipts"


def _new_upload_id() -> str:
    # URL-safe, human-pastable id
    import secrets

    return f"up_{secrets.token_urlsafe(18)}"


def _upload_prefix(upload_id: str) -> str:
    return f"uploads/{upload_id.strip().replace('..','')}/"


def _sniff_extension_for_name(data: bytes, fallback_name: str) -> str:
    ext, _ = _sniff_file_type(data)
    if ext:
        return ext
    # keep whatever the user provided
    name = (fallback_name or "").strip()
    if "." in name:
        return name.rsplit(".", 1)[-1].lower()
    return "bin"


def _download_receipts_from_blob(upload_id: str) -> tuple[list[bytes], list[str], Optional[str]]:
    """
    Returns (bytes_list, filenames, error).
    """
    if not upload_id or not str(upload_id).strip():
        return [], [], "receiptUploadId is required"
    try:
        bsc = _blob_service_client()
        container = bsc.get_container_client(_receipt_container_name())
        prefix = _upload_prefix(str(upload_id))
        blobs = list(container.list_blobs(name_starts_with=prefix))
        if not blobs:
            return [], [], f"No receipts found for upload id '{upload_id}'"

        downloaded: list[bytes] = []
        filenames: list[str] = []
        for b in blobs:
            name = str(b.name or "")
            fn = name.split("/")[-1] if "/" in name else name
            content = container.download_blob(name).readall()
            if not content:
                continue
            downloaded.append(content)
            filenames.append(fn or f"receipt-{len(filenames)+1}")

        if not downloaded:
            return [], [], f"Receipts were empty for upload id '{upload_id}'"
        return downloaded, filenames, None
    except Exception as e:
        return [], [], f"Failed to download receipts from blob storage: {e}"


def _bundle_blobs_as_attachment(*, blobs: list[bytes], filenames: list[str], payload: dict) -> tuple[list[dict], int, bool, Optional[str]]:
    """
    Returns (attachments, raw_bytes, bundled, error).
    """
    if not blobs:
        return [], 0, False, "No receipt bytes provided to bundle"
    bundle_format = _receipt_bundle_format(payload)

    if bundle_format == "pdf":
        pdfs: list[bytes] = []
        # Prepend summary table page
        summary_pdf = _build_summary_table_pdf(payload)
        if summary_pdf:
            pdfs.append(summary_pdf)
        for i, b in enumerate(blobs):
            pdf_b, err = _bytes_to_pdf(b)
            if err:
                name = filenames[i] if i < len(filenames) else f"receipt-{i+1}"
                return [], 0, False, f"receipt '{name}' {err}"
            pdfs.append(pdf_b)
        merged = _merge_pdf_bytes(pdfs)
        pdf_name = (payload.get("receiptPdfName") or "receipts.pdf").strip() or "receipts.pdf"
        if not pdf_name.lower().endswith(".pdf"):
            pdf_name = f"{pdf_name}.pdf"
        return (
            [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": pdf_name,
                    "contentType": "application/pdf",
                    "contentBytes": base64.b64encode(merged).decode("ascii"),
                }
            ],
            len(merged),
            True,
            None,
        )

    zip_name = (payload.get("receiptZipName") or "receipts.zip").strip() or "receipts.zip"
    zip_buf = BytesIO()
    with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for i, content in enumerate(blobs):
            name = filenames[i] if i < len(filenames) else f"receipt-{i+1}"
            if not name:
                name = f"receipt-{i+1}"
            zf.writestr(name, content)
    zip_bytes = zip_buf.getvalue()
    if len(zip_bytes) == 0:
        return [], 0, False, "Generated receipts.zip was empty"
    return (
        [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": zip_name,
                "contentType": "application/zip",
                "contentBytes": base64.b64encode(zip_bytes).decode("ascii"),
            }
        ],
        len(zip_bytes),
        True,
        None,
    )


# Receipt upload is a human-facing fallback page; keep it ANONYMOUS so users can access it
# without embedding a function key into client-side JavaScript.
@app.route(route="receipt-upload-init", methods=["POST"], auth_level=func.AuthLevel.ANONYMOUS)
def receipt_upload_init(req: func.HttpRequest) -> func.HttpResponse:
    """
    Returns an upload id the user can paste back into chat.
    """
    upload_id = _new_upload_id()
    return func.HttpResponse(json.dumps({"uploadId": upload_id}), mimetype="application/json")


@app.route(route="receipt-upload-file", methods=["PUT"], auth_level=func.AuthLevel.ANONYMOUS)
def receipt_upload_file(req: func.HttpRequest) -> func.HttpResponse:
    """
    Uploads a single receipt file bytes to Blob Storage.
    Query params:
      - uploadId
      - filename
    """
    upload_id = (req.params.get("uploadId") or "").strip()
    filename = (req.params.get("filename") or "").strip()
    if not upload_id:
        return func.HttpResponse(json.dumps({"error": "uploadId is required"}), status_code=400, mimetype="application/json")
    if not filename:
        filename = "receipt"
    data = req.get_body() or b""
    if len(data) == 0:
        return func.HttpResponse(json.dumps({"error": "empty body"}), status_code=400, mimetype="application/json")

    # Ensure filename has an extension when possible (helps downstream).
    ext = _sniff_extension_for_name(data, filename)
    if "." not in filename and ext:
        filename = f"{filename}.{ext}"

    content_type = (req.headers.get("Content-Type") or "").strip()
    if not content_type or content_type == "application/octet-stream":
        _ext2, sniff_type = _sniff_file_type(data)
        if sniff_type:
            content_type = sniff_type
        else:
            content_type = "application/octet-stream"

    try:
        bsc = _blob_service_client()
        container = bsc.get_container_client(_receipt_container_name())
        try:
            container.create_container()
        except Exception:
            pass

        blob_name = _upload_prefix(upload_id) + filename.replace("\\", "/").split("/")[-1]
        container.upload_blob(
            name=blob_name,
            data=data,
            overwrite=True,
            content_settings=ContentSettings(content_type=content_type),
        )
        return func.HttpResponse(json.dumps({"ok": True, "uploadId": upload_id, "blob": blob_name}), mimetype="application/json")
    except Exception as e:
        return func.HttpResponse(json.dumps({"error": f"upload failed: {e}"}), status_code=500, mimetype="application/json")


@app.route(route="receipt-upload", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
def receipt_upload_page(req: func.HttpRequest) -> func.HttpResponse:
    """
    Lightweight upload page for users who can't pass receipts through Foundry/Teams attachment APIs reliably.
    """
    html = """<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Travel Expense Receipt Upload</title>
    <style>
      body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin: 24px; }
      .box { max-width: 720px; padding: 16px; border: 1px solid #ddd; border-radius: 10px; }
      code { background: #f6f8fa; padding: 2px 6px; border-radius: 6px; }
      button { padding: 8px 12px; }
      .muted { color: #666; }
      .ok { color: #0a7; }
      .err { color: #b00; }
    </style>
  </head>
  <body>
    <div class="box">
      <h2>Upload receipts</h2>
      <p class="muted">This upload path is used when chat attachments can’t be fetched reliably. Upload PDFs or images, then paste the upload id back into chat.</p>
      <p><button id="start">Start upload</button> <span id="status" class="muted"></span></p>
      <div id="step2" style="display:none">
        <p><strong>Upload ID:</strong> <code id="uploadId"></code></p>
        <p><input id="files" type="file" multiple /></p>
        <p><button id="upload">Upload files</button></p>
        <pre id="log" class="muted" style="white-space:pre-wrap"></pre>
        <p id="done" class="ok" style="display:none">Done. Paste the upload id into chat: <code id="uploadId2"></code></p>
      </div>
    </div>
    <script>
      const statusEl = document.getElementById('status');
      const step2 = document.getElementById('step2');
      const uploadIdEl = document.getElementById('uploadId');
      const uploadIdEl2 = document.getElementById('uploadId2');
      const logEl = document.getElementById('log');
      let uploadId = null;

      function log(msg) { logEl.textContent += msg + "\\n"; }

      document.getElementById('start').addEventListener('click', async () => {
        statusEl.textContent = 'Creating upload id...';
        const resp = await fetch('./receipt-upload-init', { method: 'POST' });
        const data = await resp.json();
        uploadId = data.uploadId;
        uploadIdEl.textContent = uploadId;
        uploadIdEl2.textContent = uploadId;
        step2.style.display = '';
        statusEl.textContent = '';
      });

      document.getElementById('upload').addEventListener('click', async () => {
        const input = document.getElementById('files');
        if (!uploadId) { alert('Click Start upload first'); return; }
        if (!input.files || input.files.length === 0) { alert('Choose files'); return; }
        log('Uploading ' + input.files.length + ' file(s)...');
        for (const f of input.files) {
          log('Uploading: ' + f.name);
          const url = './receipt-upload-file?uploadId=' + encodeURIComponent(uploadId) + '&filename=' + encodeURIComponent(f.name);
          const r = await fetch(url, { method: 'PUT', headers: { 'Content-Type': f.type || 'application/octet-stream' }, body: f });
          const j = await r.json();
          if (!r.ok || !j.ok) { log('ERROR: ' + JSON.stringify(j)); statusEl.textContent='Upload failed'; statusEl.className='err'; return; }
        }
        document.getElementById('done').style.display = '';
        statusEl.textContent = 'Upload complete';
        statusEl.className = 'ok';
      });
    </script>
  </body>
</html>"""
    return func.HttpResponse(html, mimetype="text/html")


def _get_document_intelligence_client():
    """
    Returns an Azure Document Intelligence client using Managed Identity.
    Requires DOCUMENT_INTELLIGENCE_ENDPOINT env var.
    """
    endpoint = os.getenv("DOCUMENT_INTELLIGENCE_ENDPOINT", "").strip()
    if not endpoint:
        raise RuntimeError("DOCUMENT_INTELLIGENCE_ENDPOINT environment variable not set")
    credential = DefaultAzureCredential()
    return DocumentIntelligenceClient(endpoint=endpoint, credential=credential)


@app.route(route="receipt-analyze", methods=["POST"], auth_level=func.AuthLevel.FUNCTION)
def receipt_analyze(req: func.HttpRequest) -> func.HttpResponse:
    """
    Analyzes a receipt image using Azure Document Intelligence.

    Accepts either:
    - JSON body with base64 encoded image: {"imageBase64": "..."}
    - JSON body with blob reference: {"uploadId": "...", "filename": "..."}

    Returns extracted receipt data:
    - merchant: Merchant/vendor name
    - date: Transaction date
    - total: Total amount
    - subtotal: Subtotal if available
    - tax: Tax amount if available
    - items: Array of line items with description and amount
    - category: Suggested category based on merchant
    """
    try:
        body = req.get_json() if req.get_body() else {}
    except Exception:
        body = {}

    image_bytes = None

    # Option 1: Base64 encoded image in request
    image_base64 = body.get("imageBase64", "").strip()
    if image_base64:
        try:
            # Remove data URL prefix if present
            if "," in image_base64:
                image_base64 = image_base64.split(",", 1)[1]
            image_bytes = base64.b64decode(image_base64)
        except Exception as e:
            return func.HttpResponse(
                json.dumps({"ok": False, "error": f"Invalid base64 image: {e}"}),
                status_code=400,
                mimetype="application/json"
            )

    # Option 2: Fetch from blob storage
    upload_id = body.get("uploadId", "").strip()
    filename = body.get("filename", "").strip()
    if not image_bytes and upload_id:
        try:
            bsc = _blob_service_client()
            container = bsc.get_container_client(_receipt_container_name())

            # If filename provided, fetch that specific file
            if filename:
                blob_name = _upload_prefix(upload_id) + filename
                blob_client = container.get_blob_client(blob_name)
                image_bytes = blob_client.download_blob().readall()
            else:
                # Fetch the first file in the upload
                prefix = _upload_prefix(upload_id)
                blobs = list(container.list_blobs(name_starts_with=prefix))
                if blobs:
                    blob_client = container.get_blob_client(blobs[0].name)
                    image_bytes = blob_client.download_blob().readall()
        except Exception as e:
            return func.HttpResponse(
                json.dumps({"ok": False, "error": f"Failed to fetch from blob storage: {e}"}),
                status_code=400,
                mimetype="application/json"
            )

    # Option 3: Fetch from URL (for chat attachments)
    image_url = body.get("imageUrl", "").strip()
    if not image_bytes and image_url:
        try:
            logging.info("receipt-analyze fetching from URL: %s", image_url[:100])
            resp = requests.get(image_url, timeout=30)
            if resp.status_code == 200:
                image_bytes = resp.content
            else:
                return func.HttpResponse(
                    json.dumps({"ok": False, "error": f"Failed to fetch image from URL: HTTP {resp.status_code}"}),
                    status_code=400,
                    mimetype="application/json"
                )
        except Exception as e:
            return func.HttpResponse(
                json.dumps({"ok": False, "error": f"Failed to fetch image from URL: {e}"}),
                status_code=400,
                mimetype="application/json"
            )

    if not image_bytes:
        return func.HttpResponse(
            json.dumps({"ok": False, "error": "No image provided. Send imageBase64, uploadId, or imageUrl."}),
            status_code=400,
            mimetype="application/json"
        )

    # Resize large images to fit Document Intelligence limit (4 MB)
    MAX_IMAGE_BYTES = 4 * 1024 * 1024
    if len(image_bytes) > MAX_IMAGE_BYTES:
        try:
            img = Image.open(BytesIO(image_bytes))
            orig_fmt = (img.format or "").upper()
            orig_size = len(image_bytes)
            # Convert to RGB (drop alpha) so JPEG save always works
            if img.mode in ("RGBA", "LA", "P"):
                img = img.convert("RGB")
            # Progressively scale down until under limit
            scale = 0.85
            for _ in range(20):
                new_w = int(img.width * scale)
                new_h = int(img.height * scale)
                if new_w < 100 or new_h < 100:
                    break
                resized = img.resize((new_w, new_h), Image.LANCZOS)
                buf = BytesIO()
                resized.save(buf, format="JPEG", quality=85)
                if buf.tell() <= MAX_IMAGE_BYTES:
                    image_bytes = buf.getvalue()
                    logging.info("receipt-analyze resized image from %d to %d bytes (%dx%d)",
                                 orig_size, len(image_bytes), new_w, new_h)
                    break
                scale *= 0.75
                img = resized
            else:
                logging.warning("receipt-analyze could not resize image under 4MB")
        except Exception as e:
            logging.warning("receipt-analyze image resize failed: %s", e)

    # Helper to extract value from DocumentField (SDK v1.0+ removed .value)
    def _field_val(field):
        """Extract the value from a DocumentField, handling SDK version differences."""
        if field is None:
            return None
        # Newer SDK: type-specific properties
        for attr in ("value_string", "value_number", "value_integer", "value_date",
                      "value_currency", "value_array", "value_object", "content", "value"):
            v = getattr(field, attr, None)
            if v is not None:
                # value_currency is an object with .amount
                if attr == "value_currency" and hasattr(v, "amount"):
                    return v.amount
                return v
        return None

    # Analyze with Document Intelligence
    try:
        client = _get_document_intelligence_client()

        # Use prebuilt-receipt model
        poller = client.begin_analyze_document(
            "prebuilt-receipt",
            AnalyzeDocumentRequest(bytes_source=image_bytes),
        )
        result = poller.result()

        # Extract receipt data
        receipt_data = {
            "ok": True,
            "merchant": "",
            "date": "",
            "total": 0.0,
            "subtotal": 0.0,
            "tax": 0.0,
            "items": [],
            "category": "other",
            "rawText": "",
        }

        if result.documents:
            doc = result.documents[0]
            fields = doc.fields or {}

            # Merchant name
            if "MerchantName" in fields:
                val = _field_val(fields["MerchantName"])
                if val:
                    receipt_data["merchant"] = str(val)

            # Transaction date
            if "TransactionDate" in fields:
                val = _field_val(fields["TransactionDate"])
                if val:
                    if hasattr(val, "strftime"):
                        receipt_data["date"] = val.strftime("%m/%d/%Y")
                    else:
                        receipt_data["date"] = str(val)

            # Total
            if "Total" in fields:
                val = _field_val(fields["Total"])
                if val is not None:
                    receipt_data["total"] = float(val)

            # Subtotal
            if "Subtotal" in fields:
                val = _field_val(fields["Subtotal"])
                if val is not None:
                    receipt_data["subtotal"] = float(val)

            # Tax
            if "TotalTax" in fields:
                val = _field_val(fields["TotalTax"])
                if val is not None:
                    receipt_data["tax"] = float(val)

            # Line items
            if "Items" in fields:
                items_val = _field_val(fields["Items"])
                if items_val:
                    for item in items_val:
                        item_fields = _field_val(item) if not isinstance(item, dict) else item
                        if not item_fields:
                            item_fields = getattr(item, "value_object", None) or {}
                        line_item = {
                            "description": "",
                            "amount": 0.0,
                            "quantity": 0.0,
                        }
                        if "Description" in item_fields:
                            desc = _field_val(item_fields["Description"]) if not isinstance(item_fields["Description"], str) else item_fields["Description"]
                            if desc:
                                line_item["description"] = str(desc)
                        if "TotalPrice" in item_fields:
                            price = _field_val(item_fields["TotalPrice"]) if not isinstance(item_fields["TotalPrice"], (int, float)) else item_fields["TotalPrice"]
                            if price is not None:
                                line_item["amount"] = float(price)
                        if "Quantity" in item_fields:
                            qty = _field_val(item_fields["Quantity"]) if not isinstance(item_fields["Quantity"], (int, float)) else item_fields["Quantity"]
                            if qty is not None:
                                line_item["quantity"] = float(qty)
                        receipt_data["items"].append(line_item)

            # Suggest category based on merchant name
            merchant_lower = receipt_data["merchant"].lower()
            if any(kw in merchant_lower for kw in ["hotel", "marriott", "hilton", "hyatt", "inn", "suites", "lodge"]):
                receipt_data["category"] = "hotel"
            elif any(kw in merchant_lower for kw in ["airline", "delta", "united", "american", "southwest", "flight"]):
                receipt_data["category"] = "airfare"
            elif any(kw in merchant_lower for kw in ["uber", "lyft", "taxi", "cab"]):
                receipt_data["category"] = "transportation"
            elif any(kw in merchant_lower for kw in ["parking", "garage"]):
                receipt_data["category"] = "parking"
            elif any(kw in merchant_lower for kw in ["restaurant", "cafe", "coffee", "starbucks", "mcdonald", "wendy", "subway", "chipotle", "diner", "grill", "kitchen", "bistro"]):
                receipt_data["category"] = "meal"
            elif any(kw in merchant_lower for kw in ["gas", "fuel", "shell", "chevron", "exxon", "bp", "conoco", "phillips"]):
                receipt_data["category"] = "fuel"

        # Include raw text for debugging/fallback
        if result.content:
            receipt_data["rawText"] = result.content[:1000]  # Limit to first 1000 chars

        # Echo back uploadId so it can be used for receipt bundling on submit
        if upload_id:
            receipt_data["uploadId"] = upload_id

        return func.HttpResponse(
            json.dumps(receipt_data),
            mimetype="application/json"
        )

    except Exception as e:
        logging.exception("Receipt analysis failed")
        return func.HttpResponse(
            json.dumps({"ok": False, "error": f"Analysis failed: {e}"}),
            status_code=500,
            mimetype="application/json"
        )


def _foundry_get_json(project_endpoint: str, path: str) -> dict:
    base = (project_endpoint or "").rstrip("/")
    url = f"{base}{path}"
    resp = requests.get(
        url,
        headers={
            "Authorization": f"Bearer {_foundry_get_access_token()}",
            "Accept": "application/json",
        },
        timeout=30,
    )
    if resp.status_code not in (200, 202):
        raise RuntimeError(f"Foundry GET failed: url={url} status={resp.status_code} {(resp.text or '')[:500]}")
    return resp.json()


def _foundry_get_bytes(project_endpoint: str, path: str) -> bytes:
    base = (project_endpoint or "").rstrip("/")
    url = f"{base}{path}"
    resp = requests.get(
        url,
        headers={
            "Authorization": f"Bearer {_foundry_get_access_token()}",
            "Accept": "application/octet-stream",
        },
        timeout=60,
    )
    if resp.status_code not in (200, 202):
        raise RuntimeError(f"Foundry GET (bytes) failed: url={url} status={resp.status_code} {(resp.text or '')[:500]}")
    return resp.content or b""


def _foundry_list_files(project_endpoint: str) -> list[dict]:
    """
    Best-effort listing of Foundry files/assets. Returns a list of dicts.
    """
    last_err = None
    for path in (
        "/files?api-version=v1",
        "/agents/files?api-version=v1",
        "/assets?api-version=v1",
        "/agents/assets?api-version=v1",
    ):
        try:
            data = _foundry_get_json(project_endpoint, path)
            items = data.get("value") if isinstance(data, dict) else None
            if isinstance(items, list):
                logging.info("Foundry listed files via %s; count=%s", path, len(items))
                return items
        except Exception as e:
            last_err = e
            continue
    raise RuntimeError(f"Failed to list Foundry files/assets: {last_err}")


def _extract_foundry_file_id(file_obj: dict) -> Optional[str]:
    if not isinstance(file_obj, dict):
        return None
    return file_obj.get("id") or file_obj.get("file_id") or file_obj.get("fileId") or None


def _extract_foundry_file_name(file_obj: dict) -> str:
    if not isinstance(file_obj, dict):
        return ""
    return str(file_obj.get("filename") or file_obj.get("fileName") or file_obj.get("name") or "").strip()


def _parse_csvish(value) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(v).strip() for v in value if str(v).strip()]
    if isinstance(value, str):
        return [s.strip() for s in value.replace("\n", ",").split(",") if s.strip()]
    return []


def _iter_foundry_thread_file_ids(messages_json: dict) -> list[dict]:
    """
    Returns a list of {"fileId": "...", "filename": "..." } (best-effort).

    Foundry payload shapes can vary; we try common patterns:
    - messages_json.value[] with .attachments[] containing file_id / fileId / id
    """
    out: list[dict] = []
    msgs = messages_json.get("value") if isinstance(messages_json, dict) else None
    if not isinstance(msgs, list):
        return out

    for m in msgs:
        if not isinstance(m, dict):
            continue
        atts = m.get("attachments") or []
        if not isinstance(atts, list):
            continue
        for a in atts:
            if not isinstance(a, dict):
                continue
            file_id = a.get("file_id") or a.get("fileId") or a.get("id")
            if not file_id:
                continue
            filename = a.get("filename") or a.get("name") or ""
            out.append({"fileId": str(file_id), "filename": str(filename)})

    # de-dupe while preserving order
    seen = set()
    uniq = []
    for r in out:
        if r["fileId"] in seen:
            continue
        seen.add(r["fileId"])
        uniq.append(r)
    return uniq


def _deep_collect_file_ids(obj) -> list[str]:
    """
    Recursively collect values for keys that look like file ids.
    We intentionally keep this loose because Foundry message shapes vary.
    """
    found: list[str] = []
    if obj is None:
        return found
    if isinstance(obj, dict):
        for k, v in obj.items():
            lk = str(k).lower()
            if lk in {"file_id", "fileid", "file"} and isinstance(v, str):
                found.append(v)
            found.extend(_deep_collect_file_ids(v))
        return found
    if isinstance(obj, list):
        for item in obj:
            found.extend(_deep_collect_file_ids(item))
        return found
    return found


def _iter_foundry_thread_file_ids_fallback(messages_json: dict) -> list[dict]:
    """
    Fallback for Foundry message payloads that don't populate a top-level `attachments` array.

    Strategy:
    - Consider only messages whose role is `user` (avoid including knowledge/tool files).
    - Deep-scan message fields for `file_id`/`fileId` keys.
    """
    out: list[dict] = []
    msgs = messages_json.get("value") if isinstance(messages_json, dict) else None
    if not isinstance(msgs, list):
        return out

    for m in msgs:
        if not isinstance(m, dict):
            continue
        role = str(m.get("role") or "").strip().lower()
        if role and role != "user":
            continue
        file_ids = [s for s in _deep_collect_file_ids(m) if isinstance(s, str) and s.strip()]
        for fid in file_ids:
            out.append({"fileId": fid.strip(), "filename": ""})

    # de-dupe
    seen = set()
    uniq = []
    for r in out:
        if r["fileId"] in seen:
            continue
        seen.add(r["fileId"])
        uniq.append(r)
    return uniq


def _build_receipts_zip_from_foundry(payload: dict) -> tuple[list[dict], int, bool, Optional[str]]:
    """
    Fetches receipt files from Foundry and returns a single bundle attachment (Graph-compatible).

    This avoids passing large base64 blobs through the LLM/workflow path, which is error-prone.
    """
    project_endpoint = (
        (payload.get("foundryProjectEndpoint") or "").strip()
        or (os.getenv("FOUNDRY_PROJECT_ENDPOINT") or "").strip()
        or (os.getenv("PROJECT_ENDPOINT") or "").strip()
    )
    if not project_endpoint:
        return [], 0, False, "FOUNDRY_PROJECT_ENDPOINT (or PROJECT_ENDPOINT) is not configured"

    bundle_format = _receipt_bundle_format(payload)

    # If the caller can provide explicit Foundry file ids, use those directly.
    file_ids = _parse_csvish(payload.get("foundryFileIds") or payload.get("fileIds") or payload.get("file_ids"))

    if len(file_ids) > 0:
        logging.info("Foundry receipts via explicit fileIds: count=%s", len(file_ids))
        downloaded: list[bytes] = []
        for file_id in file_ids:
            content = None
            last_bytes_err = None
            for content_path in (
                f"/files/{file_id}/content?api-version=v1",
                f"/agents/files/{file_id}/content?api-version=v1",
            ):
                try:
                    content = _foundry_get_bytes(project_endpoint, content_path)
                    break
                except Exception as e:
                    last_bytes_err = e
                    continue
            if content is None or len(content) == 0:
                raise RuntimeError(f"Failed to download file {file_id}: {last_bytes_err}")
            downloaded.append(content)

        if bundle_format == "pdf":
            pdfs: list[bytes] = []
            # Prepend summary table page
            summary_pdf = _build_summary_table_pdf(payload)
            if summary_pdf:
                pdfs.append(summary_pdf)
            for i, b in enumerate(downloaded):
                pdf_b, err = _bytes_to_pdf(b)
                if err:
                    return [], 0, False, f"foundryFileIds[{i}] {err}"
                pdfs.append(pdf_b)
            merged = _merge_pdf_bytes(pdfs)
            pdf_name = (payload.get("receiptPdfName") or "receipts.pdf").strip() or "receipts.pdf"
            if not pdf_name.lower().endswith(".pdf"):
                pdf_name = f"{pdf_name}.pdf"
            return (
                [
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": pdf_name,
                        "contentType": "application/pdf",
                        "contentBytes": base64.b64encode(merged).decode("ascii"),
                    }
                ],
                len(merged),
                True,
                None,
            )

        zip_name = (payload.get("receiptZipName") or "receipts.zip").strip() or "receipts.zip"
        zip_buf = BytesIO()
        with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for i, content in enumerate(downloaded):
                sniff_ext, _sniff_type = _sniff_file_type(content)
                name = f"receipt-{i+1}"
                if sniff_ext:
                    name = f"{name}.{sniff_ext}"
                zf.writestr(name, content)
        zip_bytes = zip_buf.getvalue()
        return (
            [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": zip_name,
                    "contentType": "application/zip",
                    "contentBytes": base64.b64encode(zip_bytes).decode("ascii"),
                }
            ],
            len(zip_bytes),
            True,
            None,
        )

    # If we can't list thread messages (conv_* often 404), fall back to listing project files/assets and filtering by name.
    filename_hints = _parse_csvish(payload.get("receiptFilenameHints") or payload.get("filenameHints") or payload.get("receiptNames"))
    if filename_hints:
        try:
            all_files = _foundry_list_files(project_endpoint)
            # Match in hint order; prefer exact filename match (case-insensitive), then substring match.
            upper_hints = [h.upper() for h in filename_hints]
            by_name: list[tuple[str, dict]] = []
            for f in all_files:
                name = _extract_foundry_file_name(f)
                if not name:
                    continue
                by_name.append((name.upper(), f))

            selected: list[dict] = []
            for h in upper_hints:
                exact = next((f for (n, f) in by_name if n == h), None)
                if exact is not None:
                    selected.append(exact)
                    continue
                contains = next((f for (n, f) in by_name if h in n), None)
                if contains is not None:
                    selected.append(contains)
                    continue
                base = (
                    h.replace(".PDF", "")
                    .replace(".JPG", "")
                    .replace(".JPEG", "")
                    .replace(".PNG", "")
                )
                if base:
                    base_match = next((f for (n, f) in by_name if base in n), None)
                    if base_match is not None:
                        selected.append(base_match)

            file_ids2: list[str] = []
            seen = set()
            for f in selected[:10]:
                fid = _extract_foundry_file_id(f)
                if fid and fid not in seen:
                    seen.add(fid)
                    file_ids2.append(fid)
            if file_ids2:
                logging.info("Foundry receipts via filename hints; hints=%s fileIds=%s", filename_hints, file_ids2)
                payload2 = dict(payload)
                payload2["foundryFileIds"] = file_ids2
                return _build_receipts_zip_from_foundry(payload2)
        except Exception as e:
            logging.warning("Foundry filename-hints fallback failed: %s", e)

    conversation_id = (
        payload.get("conversationId")
        or payload.get("ConversationId")
        or payload.get("threadId")
        or payload.get("ThreadId")
        or ""
    ).strip()
    if not conversation_id:
        return (
            [],
            0,
            False,
            "Receipt files missing: no attachments were provided and no Foundry identifiers were provided. "
            "Pass foundryFileIds=<id1,id2> or receiptFilenameHints=<name1,name2>, or pass conversationId/threadId "
            "and fetchReceiptsFromThread=true (if you have an agents thread id).",
        )

    # Try the most likely thread message endpoints; Foundry API surface can differ slightly by SDK/version.
    messages_json = None
    last_err = None
    used_path = None
    tried: list[str] = []
    for path in (
        f"/threads/{conversation_id}/messages?api-version=v1",
        f"/agents/threads/{conversation_id}/messages?api-version=v1",
    ):
        try:
            tried.append(path)
            messages_json = _foundry_get_json(project_endpoint, path)
            used_path = path
            break
        except Exception as e:
            last_err = e
            continue

    if messages_json is None:
        return [], 0, False, (
            f"Failed to list thread messages for id '{conversation_id}'. "
            f"Tried: {', '.join(tried)}. Last error: {last_err}. "
            "If you pasted a Foundry trace id (conv_*), it may not equal the agents thread id. "
            "In Foundry Traces, open the message where you uploaded the PDFs and copy the file_id(s), "
            "then retry with foundryFileIds=<file_id1,file_id2>."
        )

    try:
        msg_count = len(messages_json.get("value") or []) if isinstance(messages_json, dict) else 0
        logging.info("Foundry messages listed via %s; count=%s", used_path, msg_count)
    except Exception:
        pass

    file_refs = _iter_foundry_thread_file_ids(messages_json)
    if len(file_refs) == 0:
        file_refs = _iter_foundry_thread_file_ids_fallback(messages_json)
    if len(file_refs) == 0:
        # Help troubleshoot mismatched identifiers (conv_* vs thread_*).
        return [], 0, False, (
            "No file attachments found on Foundry thread. "
            "If you pasted a Foundry trace id (conv_*), it may not equal the agents thread id."
        )

    downloaded: list[bytes] = []
    filenames: list[str] = []
    for i, ref in enumerate(file_refs):
        file_id = ref["fileId"]
        filename = (ref.get("filename") or "").strip() or f"receipt-{i+1}"

        content = None
        last_bytes_err = None
        for content_path in (
            f"/files/{file_id}/content?api-version=v1",
            f"/agents/files/{file_id}/content?api-version=v1",
        ):
            try:
                content = _foundry_get_bytes(project_endpoint, content_path)
                break
            except Exception as e:
                last_bytes_err = e
                continue

        if content is None or len(content) == 0:
            raise RuntimeError(f"Failed to download file {file_id}: {last_bytes_err}")

        sniff_ext, _sniff_type = _sniff_file_type(content)
        if "." not in filename and sniff_ext:
            filename = f"{filename}.{sniff_ext}"
        filenames.append(filename)
        downloaded.append(content)

    if bundle_format == "pdf":
        pdfs: list[bytes] = []
        # Prepend summary table page
        summary_pdf = _build_summary_table_pdf(payload)
        if summary_pdf:
            pdfs.append(summary_pdf)
        for i, b in enumerate(downloaded):
            pdf_b, err = _bytes_to_pdf(b)
            if err:
                return [], 0, False, f"thread attachment '{filenames[i]}' {err}"
            pdfs.append(pdf_b)
        merged = _merge_pdf_bytes(pdfs)
        pdf_name = (payload.get("receiptPdfName") or "receipts.pdf").strip() or "receipts.pdf"
        if not pdf_name.lower().endswith(".pdf"):
            pdf_name = f"{pdf_name}.pdf"
        return (
            [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": pdf_name,
                    "contentType": "application/pdf",
                    "contentBytes": base64.b64encode(merged).decode("ascii"),
                }
            ],
            len(merged),
            True,
            None,
        )

    zip_name = (payload.get("receiptZipName") or "receipts.zip").strip() or "receipts.zip"
    zip_buf = BytesIO()
    with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for filename, content in zip(filenames, downloaded):
            zf.writestr(filename, content)

    zip_bytes = zip_buf.getvalue()
    if len(zip_bytes) == 0:
        return [], 0, False, "Generated receipts.zip was empty"

    return (
        [
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": zip_name,
                "contentType": "application/zip",
                "contentBytes": base64.b64encode(zip_bytes).decode("ascii"),
            }
        ],
        len(zip_bytes),
        True,
        None,
    )


def _build_receipt_attachments(payload: dict) -> tuple[list[dict], int, bool, Optional[str]]:
    """
    Returns (attachments, total_raw_bytes, bundled, error).

    Expected payload.attachments items:
      { "name": "...", "contentType": "image/jpeg", "contentBytes": "<base64>" }

    Bundle selection:
    - receiptBundleFormat='pdf' (default): convert images to PDF and merge into a single receipts.pdf
    - receiptBundleFormat='zip': zip original files into receipts.zip
    """
    items = payload.get("attachments")
    if not isinstance(items, list) or len(items) == 0:
        return [], 0, False, None

    bundle_format = _receipt_bundle_format(payload)
    zip_name = (payload.get("receiptZipName") or "receipts.zip").strip() or "receipts.zip"
    pdf_name = (payload.get("receiptPdfName") or "receipts.pdf").strip() or "receipts.pdf"
    if not pdf_name.lower().endswith(".pdf"):
        pdf_name = f"{pdf_name}.pdf"

    decoded = []
    total_bytes = 0
    for i, att in enumerate(items):
        if not isinstance(att, dict):
            return [], 0, False, f"attachments[{i}] must be an object"
        name = (att.get("name") or "").strip() or f"receipt-{i+1}"
        content_type = (att.get("contentType") or "application/octet-stream").strip()
        content_b64 = (
            att.get("contentBytes")
            or att.get("contentBase64")
            or att.get("base64")
            or att.get("content")
            or att.get("data")
            or ""
        )
        try:
            data = _coerce_bytes(content_b64)
        except (binascii.Error, ValueError) as e:
            return [], 0, False, f"attachments[{i}] invalid base64: {e}"
        if len(data) == 0:
            return [], 0, False, f"attachments[{i}] is empty"
        sniff_ext, sniff_type = _sniff_file_type(data)
        if content_type in {"application/pdf", "image/png", "image/jpeg"} and sniff_type is None:
            return [], 0, False, f"attachments[{i}] content does not match {content_type}"
        if content_type == "application/octet-stream" and sniff_type is None:
            return [], 0, False, f"attachments[{i}] content type is unknown"
        if sniff_type:
            if content_type == "application/octet-stream":
                content_type = sniff_type
            if "." not in name and sniff_ext:
                name = f"{name}.{sniff_ext}"
            elif sniff_ext and name.lower().endswith(f".{sniff_ext}") is False and content_type == sniff_type:
                name = f"{name}.{sniff_ext}"
        total_bytes += len(data)
        decoded.append((name, content_type, data))

    if bundle_format == "pdf":
        pdfs: list[bytes] = []
        # Prepend summary table page
        summary_pdf = _build_summary_table_pdf(payload)
        if summary_pdf:
            pdfs.append(summary_pdf)
        for i, (name, _content_type, data) in enumerate(decoded):
            pdf_b, err = _bytes_to_pdf(data)
            if err:
                return [], 0, False, f"attachments[{i}] '{name}' {err}"
            pdfs.append(pdf_b)
        merged = _merge_pdf_bytes(pdfs)
        return (
            [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": pdf_name,
                    "contentType": "application/pdf",
                    "contentBytes": base64.b64encode(merged).decode("ascii"),
                }
            ],
            len(merged),
            True,
            None,
        )

    if len(decoded) >= 2 and bool(payload.get("zipReceipts", True)):
        mem = BytesIO()
        with zipfile.ZipFile(mem, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
            for name, _, data in decoded:
                zf.writestr(name, data)
        zip_bytes = mem.getvalue()
        return (
            [
                {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": zip_name,
                    "contentType": "application/zip",
                    "contentBytes": base64.b64encode(zip_bytes).decode("ascii"),
                }
            ],
            len(zip_bytes),
            True,
            None,
        )

    graph_attachments = []
    for name, content_type, data in decoded:
        graph_attachments.append(
            {
                "@odata.type": "#microsoft.graph.fileAttachment",
                "name": name,
                "contentType": content_type,
                "contentBytes": base64.b64encode(data).decode("ascii"),
            }
        )
    return graph_attachments, total_bytes, False, None


@app.route(route="submit-report", methods=["POST"], auth_level=func.AuthLevel.FUNCTION)
def submit_report(req: func.HttpRequest) -> func.HttpResponse:
    payload: dict = {}
    try:
        parsed = req.get_json()
        if isinstance(parsed, dict):
            payload = parsed
        else:
            # Allow callers to send the draft items as a bare JSON array.
            payload = {"draftItemsJson": json.dumps(parsed)}
    except Exception:
        # Copilot Studio "tools" sometimes struggle to send a typed JSON object body.
        # Accept a raw string body as draftItemsJson and read the remaining fields from query params.
        try:
            raw_body = (req.get_body() or b"").decode("utf-8", errors="ignore").strip()
        except Exception:
            raw_body = ""
        if raw_body:
            payload = {"draftItemsJson": raw_body}
        else:
            payload = {}

    # Merge any query-string overrides (useful for custom connector tools).
    try:
        params = req.params or {}

        def _q(name: str) -> str:
            return (params.get(name) or "").strip()

        def _q_bool(name: str):
            v = _q(name).lower()
            if v == "":
                return None
            if v in {"1", "true", "yes", "y"}:
                return True
            if v in {"0", "false", "no", "n"}:
                return False
            return None

        for k in [
            "toEmail",
            "fromUser",
            "requesterEmail",
            "ccEmails",
            "subject",
            "bodyText",
            "bodyHtml",
            "receiptBundleFormat",
            "sharepointFileUrls",
            "receiptUploadId",
            "draftItemsJson",
        ]:
            v = _q(k)
            if v and (k not in payload or payload.get(k) in {None, ""}):
                payload[k] = v

        for k in ["sendEmail", "ccRequester", "purgeSharepointReceipts", "allowMissingReceipts"]:
            vb = _q_bool(k)
            if vb is not None and (k not in payload or payload.get(k) is None):
                payload[k] = vb
    except Exception as e:
        logging.warning("submit-report query param merge failed: %s", e)

    try:
        logging.info("submit-report version=%s payload_keys=%s", APP_VERSION, sorted(list(payload.keys())))
    except Exception:
        pass

    # Best-effort: try to recover conversation/thread id from request headers if the caller didn't pass it.
    # This is important when the agent chat is the only interaction surface and no workflow can inject System.ConversationId.
    header_conversation_id = ""
    try:
        # Do NOT log Authorization/cookies; only log a small allowlist + any conversation-like headers.
        interesting = {}
        for k, v in (req.headers or {}).items():
            lk = (k or "").lower()
            if lk in {"authorization", "cookie", "set-cookie", "x-functions-key"}:
                continue
            if "conversation" in lk or "thread" in lk:
                interesting[k] = v
            elif lk in {"traceparent", "tracestate", "request-id", "x-request-id", "x-ms-request-id", "x-ms-client-request-id"}:
                interesting[k] = v
        if interesting:
            logging.info("submit-report request headers (filtered): %s", interesting)
        # Common patterns (unknown until we observe them in your environment).
        header_conversation_id = (
            (req.headers.get("x-ms-conversation-id") if req.headers else None)
            or (req.headers.get("x-conversation-id") if req.headers else None)
            or (req.headers.get("conversation-id") if req.headers else None)
            or (req.headers.get("x-ms-thread-id") if req.headers else None)
            or (req.headers.get("x-thread-id") if req.headers else None)
            or ""
        )
        header_conversation_id = (header_conversation_id or "").strip()
        if header_conversation_id and "conversationId" not in payload and "ConversationId" not in payload:
            payload["conversationId"] = header_conversation_id
    except Exception as e:
        logging.warning("submit-report header inspection failed: %s", e)

    output = StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=_IMPORTFORMAT_FIELDS, extrasaction="ignore")
    line_count = 0
    amount_total = 0.0
    missing_gl_count = 0
    for row in _iter_import_rows(payload):
        writer.writerow(row)
        line_count += 1
        if not (row.get("GL Account") or "").strip():
            missing_gl_count += 1
        try:
            amount_total += float(row.get("Amount") or 0)
        except Exception:
            pass

    csv_text = output.getvalue()

    def _payload_bool(name: str, default: Optional[bool] = None) -> Optional[bool]:
        v = payload.get(name)
        if v is None:
            return default
        if isinstance(v, bool):
            return v
        if isinstance(v, (int, float)):
            return bool(v)
        if isinstance(v, str):
            s = v.strip().lower()
            if s in {"1", "true", "yes", "y"}:
                return True
            if s in {"0", "false", "no", "n"}:
                return False
        return default

    enable_email_default = (os.getenv("ENABLE_EMAIL_SEND") or "").strip().lower() in {"1", "true", "yes", "y"}
    # Default behavior for submit-report is to send email unless explicitly disabled.
    requested_send_email = bool(_payload_bool("sendEmail", True))
    # Safety: require both the request AND the env flag to be true.
    send_email = requested_send_email and enable_email_default

    to_email = (payload.get("toEmail") or os.getenv("MAIL_TO_DEFAULT") or "").strip() or "sjelinski@core.coop"
    from_user = (payload.get("fromUser") or os.getenv("MAIL_FROM_USER") or "").strip()

    requester_email = (payload.get("requesterEmail") or "").strip().lower()
    cc_requester = _payload_bool("ccRequester", None)
    if cc_requester is None:
        cc_requester = bool(requester_email)
    cc_emails = []
    if cc_requester and requester_email and requester_email != to_email.lower():
        cc_emails.append(requester_email)

    # Optional: allow callers (topics/tools) to CC additional recipients.
    # Accept comma/semicolon-separated string or a list of strings.
    cc_extra = payload.get("ccEmails") or payload.get("ccEmail") or payload.get("cc")
    extra_list: list[str] = []
    if isinstance(cc_extra, str):
        for part in re.split(r"[;,]", cc_extra):
            p = (part or "").strip().lower()
            if p:
                extra_list.append(p)
    elif isinstance(cc_extra, list):
        for v in cc_extra:
            if isinstance(v, str):
                p = v.strip().lower()
                if p:
                    extra_list.append(p)

    for e in extra_list:
        if not e or e == to_email.lower():
            continue
        if e not in cc_emails:
            cc_emails.append(e)

    subject = (payload.get("subject") or "").strip() or "Travel expense submission"
    body_text = (payload.get("bodyText") or "").strip() or (
        f"Travel expense submission generated by the bot.\nLines: {line_count}\nTotal: {amount_total:.2f}"
    )
    body_html = payload.get("bodyHtml")

    mail_error = None
    attachments = []
    attachments_zipped = False
    attachment_bytes = 0
    fetch_from_thread = bool(_payload_bool("fetchReceiptsFromThread", False))

    # Determine if the payload includes any receipt items and collect uploadIds from items.
    items = payload.get("items")
    if items is None and isinstance(payload.get("draftItemsJson"), str):
        try:
            items = json.loads(payload["draftItemsJson"])
        except Exception:
            items = None
    has_receipts = False
    item_upload_ids: list[str] = []
    if isinstance(items, list):
        for it in items:
            if isinstance(it, dict):
                item_type = str(it.get("type") or it.get("mode") or "").strip().lower()
                if item_type in ("receipt", "boots"):
                    has_receipts = True
                    # Collect receiptUploadId from individual items
                    item_upload_id = (it.get("receiptUploadId") or it.get("uploadId") or "").strip()
                    if item_upload_id and item_upload_id not in item_upload_ids:
                        item_upload_ids.append(item_upload_id)
                    # Boots flow can include a required authorization form attachment.
                    auth_upload_id = (
                        it.get("bootAuthorizationUploadId")
                        or it.get("bootsAuthorizationUploadId")
                        or it.get("authorizationUploadId")
                        or ""
                    ).strip()
                    if auth_upload_id and auth_upload_id not in item_upload_ids:
                        item_upload_ids.append(auth_upload_id)

    if payload.get("attachments") is not None:
        attachments, attachment_bytes, attachments_zipped, att_error = _build_receipt_attachments(payload)
        if att_error:
            mail_error = att_error

    # Preferred: SharePoint temp drop. Agent uploads receipts to SharePoint, then passes driveId+itemIds (or URLs) here.
    sharepoint_drive_id = _coalesce(payload.get("sharepointDriveId"), payload.get("sharePointDriveId"), payload.get("spDriveId"))
    sharepoint_item_ids = _parse_sharepoint_item_ids(payload)
    sharepoint_urls = _parse_sharepoint_urls(payload)
    if (has_receipts or fetch_from_thread) and not mail_error and len(attachments) == 0 and (
        (sharepoint_drive_id and len(sharepoint_item_ids) > 0) or len(sharepoint_urls) > 0
    ):
        logging.info(
            "submit-report receipt-bundle: source=sharepoint driveId=%s itemCount=%s urlCount=%s",
            sharepoint_drive_id,
            len(sharepoint_item_ids),
            len(sharepoint_urls),
        )
        sp_bytes, sp_names, sp_err = _download_receipts_from_sharepoint(payload)
        if sp_err:
            logging.warning("submit-report receipt-bundle (sharepoint) failed: %s", sp_err)
            mail_error = sp_err
        else:
            sp_atts, sp_count, sp_bundled, bundle_err = _bundle_blobs_as_attachment(blobs=sp_bytes, filenames=sp_names, payload=payload)
            if bundle_err:
                logging.warning("submit-report receipt-bundle (sharepoint) failed: %s", bundle_err)
                mail_error = bundle_err
            else:
                attachments = sp_atts
                attachment_bytes = sp_count
                attachments_zipped = sp_bundled

    # Fallback: If receipts were uploaded via the receipt-upload page, fetch them from Blob Storage and bundle them.
    # Check both top-level receiptUploadId and item-level uploadIds collected above.
    receipt_upload_id = _coalesce(payload.get("receiptUploadId"), payload.get("ReceiptUploadId"), payload.get("uploadId"), payload.get("UploadId"))
    all_upload_ids = item_upload_ids.copy()
    if receipt_upload_id and receipt_upload_id not in all_upload_ids:
        all_upload_ids.insert(0, receipt_upload_id)

    if (has_receipts or fetch_from_thread) and not mail_error and len(attachments) == 0 and len(all_upload_ids) > 0:
        logging.info("submit-report receipt-bundle: source=blob uploadIds=%s", all_upload_ids)
        all_blob_bytes: list[bytes] = []
        all_blob_names: list[str] = []
        blob_err = None
        for uid in all_upload_ids:
            uid_bytes, uid_names, uid_err = _download_receipts_from_blob(uid)
            if uid_err:
                logging.warning("submit-report receipt-bundle (blob) uploadId=%s failed: %s", uid, uid_err)
                # Continue to try other uploadIds, but track the error
                if blob_err is None:
                    blob_err = uid_err
            else:
                all_blob_bytes.extend(uid_bytes)
                all_blob_names.extend(uid_names)

        if len(all_blob_bytes) > 0:
            blob_err = None  # Clear error if we got some receipts

        if blob_err:
            logging.warning("submit-report receipt-bundle (blob) failed: %s", blob_err)
            mail_error = blob_err
        elif len(all_blob_bytes) > 0:
            blob_atts, blob_count, blob_bundled, bundle_err = _bundle_blobs_as_attachment(
                blobs=all_blob_bytes,
                filenames=all_blob_names,
                payload=payload,
            )
            if bundle_err:
                logging.warning("submit-report receipt-bundle (blob) failed: %s", bundle_err)
                mail_error = bundle_err
            else:
                attachments = blob_atts
                attachment_bytes = blob_count
                attachments_zipped = blob_bundled

    # If we have receipt items but no usable attachments were provided, fetch original files from Foundry
    # and build a single receipts bundle (default: receipts.pdf) server-side.
    conv_for_fetch = (
        payload.get("conversationId")
        or payload.get("ConversationId")
        or payload.get("threadId")
        or payload.get("ThreadId")
        or ""
    )
    has_file_ids = bool(payload.get("foundryFileIds") or payload.get("fileIds") or payload.get("file_ids"))
    has_filename_hints = bool(payload.get("receiptFilenameHints") or payload.get("filenameHints") or payload.get("receiptNames"))
    if (has_receipts or fetch_from_thread) and not mail_error and len(attachments) == 0 and (conv_for_fetch or has_file_ids or has_filename_hints or fetch_from_thread):
        logging.info(
            "submit-report receipt-bundle: has_receipts=%s fetchReceiptsFromThread=%s conversationId=%s",
            has_receipts,
            fetch_from_thread,
            (conv_for_fetch or "").strip(),
        )
        foundry_attachments, foundry_bytes, foundry_zipped, foundry_err = _build_receipts_zip_from_foundry(payload)
        if foundry_err:
            # Hard stop: we don't want "sent=true" emails without receipts.pdf when receipts exist.
            logging.warning("submit-report receipt-bundle failed: %s", foundry_err)
            mail_error = foundry_err
        else:
            logging.info("submit-report receipt-bundle success: bytes=%s", foundry_bytes)
            attachments = foundry_attachments
            attachment_bytes = foundry_bytes
            attachments_zipped = foundry_zipped

    # Final safety: if receipt items are present and we're sending email, never send without receipts unless explicitly allowed.
    allow_missing_receipts = bool(_payload_bool("allowMissingReceipts", False))
    if requested_send_email and has_receipts and len(attachments) == 0 and not allow_missing_receipts:
        upload_url = (os.getenv("RECEIPTS_UPLOAD_PAGE_URL") or "").strip()
        upload_hint = f" Upload receipts at: {upload_url}" if upload_url else ""
        mail_error = (
            "Receipt files missing: no attachments were provided and no Foundry identifiers were provided to fetch receipts.pdf. "
            "Pass sharepointDriveId+sharepointItemIds (preferred) or sharepointFileUrls, or pass foundryFileIds/receiptFilenameHints, "
            "or pass conversationId/threadId and fetchReceiptsFromThread=true, "
            "or set allowMissingReceipts=true to send without receipts." + upload_hint
        )

    if requested_send_email and not enable_email_default:
        mail_error = "Email sending is disabled (ENABLE_EMAIL_SEND is not true)."
    elif requested_send_email and not from_user:
        mail_error = "MAIL_FROM_USER is required to send email via Graph."
    elif requested_send_email and missing_gl_count > 0:
        mail_error = f"Missing GL Account on {missing_gl_count} line(s)."
    elif send_email:
        # Graph sendMail simple attachments are limited to ~3 MB raw (4 MB base64).
        # Images are resized in _bytes_to_pdf so combined PDF is typically well under this.
        max_raw = int(os.getenv("GRAPH_MAX_ATTACHMENT_BYTES") or "7000000")
        if attachment_bytes > max_raw:
            mail_error = f"Attachments too large ({attachment_bytes} bytes). Reduce size or set GRAPH_MAX_ATTACHMENT_BYTES higher."

        if mail_error is None:
            logging.info(
                "submit-report attachments: count=%s zipped=%s bytes=%s",
                len(attachments),
                attachments_zipped,
                attachment_bytes,
            )
            mail_error = _graph_send_mail(
                from_user=from_user,
                to_email=to_email,
                cc_emails=cc_emails,
                subject=subject,
                body_text=body_text,
                body_html=body_html,
                csv_text=csv_text,
                additional_attachments=attachments,
            )
            # Best-effort purge of temp SharePoint receipts after successful send.
            if mail_error is None and bool(_payload_bool("purgeSharepointReceipts", False)):
                purge_err = _purge_sharepoint_items(payload)
                if purge_err:
                    logging.warning(purge_err)

    if requested_send_email and mail_error:
        logging.warning("submit-report not sent: %s", mail_error)

    return func.HttpResponse(
        json.dumps(
            {
                "ok": (not requested_send_email) or (send_email and mail_error is None),
                "sent": send_email and mail_error is None,
                "toEmail": to_email,
                "lineCount": line_count,
                "amountTotal": round(amount_total, 2),
                "csvFilename": "travel-expense.csv",
                "emailError": mail_error,
                "attachmentCount": len(attachments),
                "attachmentsZipped": attachments_zipped,
                "hasReceiptItems": has_receipts,
                "fetchReceiptsFromThread": fetch_from_thread,
                "allowMissingReceipts": allow_missing_receipts,
                "conversationId": _coalesce(
                    payload.get("conversationId"),
                    payload.get("ConversationId"),
                    payload.get("threadId"),
                    payload.get("ThreadId"),
                ),
                "conversationIdFromHeaders": header_conversation_id or "",
            }
        ),
        mimetype="application/json",
    )
