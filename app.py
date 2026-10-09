import io
import os
import re
import time
import threading
import datetime
import uuid
import json as _json
import requests
import openpyxl
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from flask import Flask, render_template, request, jsonify, send_file, Response, stream_with_context
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)

_token_cache = {"token": None, "expires_at": 0}

TOKEN_URL      = "https://id.cisco.com/oauth2/default/v1/token"
EOX_BASE       = "https://apix.cisco.com/supporttools/eox/rest/5"
SW_SUGGEST     = "https://apix.cisco.com/software/suggestion/v2/suggestions/software/productIds"
SW_SUGGEST_MDF = "https://apix.cisco.com/software/suggestion/v2/suggestions/software/mdfIds"
PRODUCT_INFO   = "https://apix.cisco.com/product/v1/information/product_ids"
PRODUCT_MDF    = "https://apix.cisco.com/product/v1/information/product_ids_mdf"

# Job state (in-memory; this is a single-user dev tool)
_jobs:       dict[str, dict] = {}
_job_events: dict[str, list] = {}
_job_lock = threading.Lock()

# Persistent data store (server-side JSON files)
_DATA_DIR  = os.path.join(os.path.dirname(__file__), "data")
_PRESETS_F   = os.path.join(_DATA_DIR, "presets.json")
_HISTORY_F   = os.path.join(_DATA_DIR, "history.json")
_OVERRIDES_F = os.path.join(_DATA_DIR, "overrides.json")
_DOWNLOADS_DIR = os.path.join(_DATA_DIR, "downloads")
os.makedirs(_DOWNLOADS_DIR, exist_ok=True)
_data_lock = threading.Lock()


def _read_json(path: str, default):
    try:
        with open(path, encoding="utf-8") as f:
            return _json.load(f)
    except Exception:
        return default


def _write_json(path: str, data):
    os.makedirs(_DATA_DIR, exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        _json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp, path)


# ── Auth ───────────────────────────────────────────────────────────────────────

def get_access_token(client_id: str, client_secret: str) -> str:
    now = time.time()
    if _token_cache["token"] and now < _token_cache["expires_at"]:
        return _token_cache["token"]
    resp = requests.post(
        TOKEN_URL,
        data={"grant_type": "client_credentials", "client_id": client_id, "client_secret": client_secret},
        timeout=15,
    )
    resp.raise_for_status()
    data = resp.json()
    _token_cache["token"]      = data["access_token"]
    _token_cache["expires_at"] = now + data.get("expires_in", 3600) - 60
    return _token_cache["token"]


# ── API fetchers ───────────────────────────────────────────────────────────────

def _parse_eox_records(records: list[dict], result: dict[str, dict]):
    # When the API returns multiple records for the same PID, keep only the first.
    for rec in records:
        canonical = (rec.get("EOLProductID") or "").strip().upper()
        queried   = (rec.get("EOXInputValue") or "").strip().upper()

        if canonical:
            # Store under canonical key (used for chain resolution)
            if canonical not in result:
                result[canonical] = rec
            # Also store under the originally queried PID when it differs
            # (e.g. "WS-C2960X-24PS-BR" → EOLProductID "WS-C2960X-24PS-L"):
            # this lets description/info lookups find the record via the variant key.
            if queried and queried != canonical:
                if queried not in result:
                    result[queried] = rec
            continue

        error = rec.get("EOXError") or {}
        desc  = error.get("ErrorDescription", "")
        m = re.search(r'\[([^\]]+)\]', desc)
        if m:
            key = m.group(1).upper()
            if key not in result:
                result[key] = rec
            continue
        m = re.search(r'product ID\(?s?\)?: (.+)$', desc, re.IGNORECASE)
        if m:
            for p in m.group(1).split(","):
                key = p.strip().upper()
                if key not in result:
                    result[key] = rec


def _fetch_one_eox(token: str, headers: dict, pid: str) -> dict[str, dict]:
    try:
        url = f"{EOX_BASE}/EOXByProductID/1/{requests.utils.quote(pid)}"
        r = requests.get(url, params={"responseencoding": "json"},
                         headers=headers, timeout=30)
        if r.ok:
            out: dict[str, dict] = {}
            for rec in r.json().get("EOXRecord", []):
                pid_in_rec = (rec.get("EOLProductID") or "").strip().upper()
                # Only store under the canonical key returned by the API.
                # Never alias the queried PID → a different record: if the API
                # returns EOLProductID="WS-C2960G-8TC-L" when we queried
                # "WS-C2960-8TC-L", storing the G-variant record under the
                # queried key would corrupt chain resolution with wrong data.
                # When the API returns multiple records for the same PID, keep only the first.
                key = pid_in_rec or pid.upper()
                if key not in out:
                    out[key] = rec
            return out
    except Exception:
        pass
    return {}


def fetch_eox(token: str, pids: list[str], on_progress=None) -> dict[str, dict]:
    result: dict[str, dict] = {}
    headers = {"Authorization": f"Bearer {token}"}
    for i in range(0, len(pids), 20):
        batch = pids[i:i + 20]
        pid_str = ",".join(batch)
        url = f"{EOX_BASE}/EOXByProductID/1/{requests.utils.quote(pid_str)}"
        r = _get_with_retry(url, headers, params={"responseencoding": "json"})
        if r and r.ok:
            _parse_eox_records(r.json().get("EOXRecord", []), result)
        if on_progress:
            on_progress(min(i + 20, len(pids)))

    for p in pids:
        if p.upper() not in result:
            result.update(_fetch_one_eox(token, headers, p))

    if on_progress:
        on_progress(len(pids))
    return result


def _get_with_retry(url: str, headers: dict, params=None, retries: int = 3) -> "requests.Response | None":
    for attempt in range(retries):
        try:
            r = requests.get(url, headers=headers, params=params, timeout=30)
            if r.ok:
                return r
            # Retry on transient server/rate-limit errors only
            if r.status_code not in (429, 500, 502, 503, 504):
                return r
        except requests.exceptions.Timeout:
            pass
        except Exception:
            return None
        if attempt < retries - 1:
            time.sleep(0.4 * (attempt + 1))
    return None


def _fetch_one_product(token: str, headers: dict, pid: str) -> dict[str, dict]:
    r = _get_with_retry(f"{PRODUCT_INFO}/{requests.utils.quote(pid)}", headers)
    if r and r.ok:
        for item in r.json().get("product_list", []):
            key = (item.get("product_id") or "").upper()
            if key:
                return {key: item}
    return {}


def fetch_product_info(token: str, pids: list[str], on_progress=None) -> dict[str, dict]:
    result: dict[str, dict] = {}
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    for i in range(0, len(pids), 20):
        batch = pids[i:i + 20]
        pid_str = ",".join(batch)
        r = _get_with_retry(f"{PRODUCT_INFO}/{requests.utils.quote(pid_str)}", headers)
        if r and r.ok:
            for item in r.json().get("product_list", []):
                pid = (item.get("product_id") or "").upper()
                if pid:
                    result[pid] = item
        if on_progress:
            on_progress(min(i + 20, len(pids)))

    for p in pids:
        if p.upper() not in result:
            result.update(_fetch_one_product(token, headers, p))

    if on_progress:
        on_progress(len(pids))
    return result


def _parse_sw_suggestions(product_list: list, result: dict, pid_key: str = "basePID", mdf_to_pid: dict = None):
    for item in product_list:
        prod = item.get("product", {})
        if mdf_to_pid is not None:
            pid = mdf_to_pid.get(prod.get("mdfId", ""), "")
        else:
            pid = (prod.get(pid_key) or "").upper()
        suggestions = item.get("suggestions", [])
        if not pid or not suggestions:
            continue
        err = suggestions[0].get("errorDetailsResponse")
        if err and err.get("errorCode"):
            continue
        best = next((s for s in suggestions if str(s.get("isSuggested", "")).upper() == "Y"), suggestions[0])
        best["softwareType"] = prod.get("softwareType", "")
        result[pid] = best


def fetch_sw_suggestions(token: str, pids: list[str], on_progress=None) -> dict[str, dict] | None:
    result: dict[str, dict] = {}
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    # Non-slash PIDs: batch query (max 10 per request)
    queryable  = [p for p in pids if "/" not in p]
    slash_pids = [p for p in pids if "/" in p]

    # Track every PID the batch API explicitly responded to (success OR error).
    # PIDs in this set already got their definitive answer from the batch — no
    # point retrying them individually if they simply have no SW suggestion.
    _batch_seen: set[str] = set()

    for i in range(0, len(queryable), 10):
        pid_str = ",".join(queryable[i:i + 10])
        r = _get_with_retry(f"{SW_SUGGEST}/{pid_str}", headers)
        if r is None:
            continue
        if r.status_code == 403:
            return None
        if not r.ok:
            continue
        body = r.json()
        for item in body.get("productList", []):
            p = (item.get("product", {}).get("basePID") or "").upper()
            if p:
                _batch_seen.add(p)
        _parse_sw_suggestions(body.get("productList", []), result)
        if on_progress:
            on_progress(min(i + 10, len(queryable)))

    # Sequential fallback — only for PIDs the batch API completely omitted
    # (not those it returned with an error like S3_NO_REC_FOUND, which are
    # already definitive and would waste one round-trip each if retried).
    _fallback = [p for p in queryable if p.upper() not in result and p.upper() not in _batch_seen]

    # Total "remaining" work after the batch: fallback PIDs + slash PIDs.
    # Progress is reported linearly across this remaining count.
    _remaining_total = len(_fallback) + len(slash_pids)
    _remaining_done  = 0

    def _remaining_progress():
        if on_progress and _remaining_total:
            on_progress(min(len(queryable) + _remaining_done, len(pids) - 1))

    for pid in _fallback:
        r = _get_with_retry(f"{SW_SUGGEST}/{requests.utils.quote(pid)}", headers)
        if r is None:
            pass
        elif r.status_code == 403:
            return None
        elif r.ok:
            before = set(result.keys())
            _parse_sw_suggestions(r.json().get("productList", []), result)
            new_keys = set(result.keys()) - before
            if new_keys and pid.upper() not in result:
                result[pid.upper()] = result[next(iter(new_keys))]
        _remaining_done += 1
        _remaining_progress()

    # Slash PIDs: MDF lookup → SW suggestion by MDF ID
    if slash_pids:
        mdf_map: dict[str, str] = {}  # {pid_upper: mdf_id}
        for i in range(0, len(slash_pids), 5):
            pid_str = ",".join(slash_pids[i:i + 5])
            r = _get_with_retry(
                f"{PRODUCT_MDF}/{requests.utils.quote(pid_str, safe='/')}",
                headers,
            )
            if r and r.ok:
                for item in r.json().get("product_list", []):
                    pid_    = (item.get("product_id") or "").upper()
                    mdf_id  = item.get("product_name_mdf", "")
                    if pid_ and mdf_id:
                        mdf_map[pid_] = mdf_id
            _remaining_done += len(slash_pids[i:i + 5])
            _remaining_progress()

        if mdf_map:
            mdf_to_pid = {v: k for k, v in mdf_map.items()}  # {mdf_id: pid}
            mdf_ids = list(mdf_to_pid.keys())
            for i in range(0, len(mdf_ids), 10):
                mdf_str = ",".join(mdf_ids[i:i + 10])
                r = _get_with_retry(f"{SW_SUGGEST_MDF}/{mdf_str}", headers)
                if r is None:
                    continue
                if r.status_code == 403:
                    return None
                if not r.ok:
                    continue
                _parse_sw_suggestions(r.json().get("productList", []), result, mdf_to_pid=mdf_to_pid)

    if on_progress:
        on_progress(len(pids))

    return result


_SW_SUFFIXES = ["-A", "-E"]


def _sw_suffix_fallback(token: str, pids: list[str], info_map: dict, sw_map: dict, on_progress=None) -> int:
    """For valid-HW PIDs with no SW suggestion, retry with -A / -E suffixes.

    Results are stored back into sw_map keyed by the *original* PID.
    Returns the number of PIDs resolved via fallback.
    """
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    candidates = [p for p in pids if p.upper() not in sw_map and info_map.get(p.upper(), {}).get("product_name")]
    if not candidates:
        return 0

    variants: list[tuple[str, str]] = []
    for pid in candidates:
        for suffix in _SW_SUFFIXES:
            variants.append((pid.upper(), pid.upper() + suffix))

    suffixed_pids = [v for _, v in variants]
    suffix_result: dict[str, dict] = {}
    total_batches = max((len(suffixed_pids) + 9) // 10, 1)
    for batch_idx, i in enumerate(range(0, len(suffixed_pids), 10)):
        batch = suffixed_pids[i:i + 10]
        resp = _get_with_retry(f"{SW_SUGGEST}/{','.join(batch)}", headers)
        if resp is None:
            pass
        elif resp.status_code == 403:
            break
        elif resp.ok:
            _parse_sw_suggestions(resp.json().get("productList", []), suffix_result)
        if on_progress:
            on_progress(batch_idx + 1, total_batches)

    resolved = 0
    for orig_pid, suffixed_pid in variants:
        if orig_pid not in sw_map and suffixed_pid in suffix_result:
            sw_map[orig_pid] = suffix_result[suffixed_pid]
            resolved += 1

    return resolved


# Maps softwareType values returned by the SW Suggestion API to the OSType
# identifiers expected by EOXBySWReleaseString.
_SW_TYPE_TO_OSTYPE: dict[str, str] = {
    "IOS XE SOFTWARE":  "IOS-XE",
    "IOS-XE SOFTWARE":  "IOS-XE",
    "IOS XE":           "IOS-XE",
    "IOS-XE":           "IOS-XE",
    "IOS XR SOFTWARE":  "IOS XR",
    "IOS XR":           "IOS XR",
    "IOS SOFTWARE":     "IOS",
    "IOS":              "IOS",
    "NX-OS SOFTWARE":   "NX-OS",
    "NX-OS":            "NX-OS",
    "ASA SOFTWARE":     "ASA",
    "ASA":              "ASA",
    "WLC SOFTWARE":     "WLC",
    "WLC":              "WLC",
    "SAN-OS SOFTWARE":  "SAN-OS",
    "SAN-OS":           "SAN-OS",
}


# OSTypes that use dot notation (e.g. "17.15.5") — work as-is with the EoX API.
# All other types use classic IOS parenthesis notation (e.g. "15.2(7)E13").
_OSTYPE_DOT_NOTATION = {"IOS-XE", "IOS XR", "NX-OS"}


def _to_eox_release_fmt(release: str, ostype: str) -> str:
    """Convert SW Suggestion releaseFormat1 to the format EOXBySWReleaseString expects.

    IOS-XE / IOS XR / NX-OS return dot notation (17.15.5) which the EoX API accepts.
    Classic IOS uses parenthesis notation but SW Suggestion returns dots/hyphens:
      15.2.7E13   →  15.2(7)E13
      15.5.1-SY17 →  15.5(1)SY17   (hyphen before train is stripped)
      12.4.15T    →  12.4(15)T
    """
    if ostype in _OSTYPE_DOT_NOTATION:
        return release
    # Major.Minor.Patch[-]Train  →  Major.Minor(Patch)Train
    m = re.match(r'^(\d+\.\d+)\.(\d+)-?([A-Za-z].*)?$', release)
    if m:
        return f"{m.group(1)}({m.group(2)}){m.group(3) or ''}"
    return release


_SW_MONTHS = {m: n for n, m in enumerate(
    ("JAN", "FEB", "MAR", "APR", "MAY", "JUN",
     "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"), start=1)}


def _parse_sw_release_date(value: str) -> "datetime.date | None":
    """Parse the SW Suggestion releaseDate format ('19-Sep-2025')."""
    try:
        d, m, y = (value or "").strip().split("-")
        return datetime.date(int(y), _SW_MONTHS[m.upper()[:3]], int(d))
    except Exception:
        return None


def _iso_date(value: str) -> "datetime.date | None":
    try:
        return datetime.date.fromisoformat((value or "").strip())
    except Exception:
        return None


def _sw_eox_applies(eovss: str, ldos: str, release_date: "datetime.date | None") -> bool:
    """Reject EOXBySWReleaseString records that cannot describe this release.

    EOXBySWReleaseString answers with the *orderable software licence SKUs* that
    carry a train (SGRWILK9-15207E, SASR9031NPEK9-35SA…), across every platform
    that ever shipped it — not with the support lifecycle of the exact maintenance
    release we asked about.  When those SKUs were retired long before the release
    shipped, their dates are meaningless here: 15.2(7)E13 shipped 19-Sep-2025 yet
    the matching records claim EoVSS 2024-03-30, which is how a C1000-16T-2G-L
    ended up advertising software support ending *before* its own hardware dates.
    A release cannot lose support before it exists, so drop those records and
    leave the SW columns empty rather than print a date that undercuts the
    hardware lifecycle.
    """
    if not release_date:
        return True
    for value in (eovss, ldos):
        d = _iso_date(value)
        if d and d < release_date:
            return False
    return True


def enrich_sw_eox(token: str, sw_map: dict, on_progress=None):
    """Enrich sw_map entries with EoVSS and LDoS from EOXBySWReleaseString.

    Collects unique (releaseFormat1, softwareType) pairs, normalises softwareType
    to the OSType identifier the EoX API expects (e.g. "IOS XE Software" → "IOS-XE"),
    queries once per unique pair, then injects _sw_eovss and _sw_ldos into each
    matching sw_map entry. Skips entries where sw_map value is None (no SW subscription).
    """
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    # Group PIDs by (release, normalised_ostype) — query each pair only once
    pair_to_pids: dict[tuple[str, str], list[str]] = {}
    pair_release_date: dict[tuple[str, str], "datetime.date | None"] = {}
    for pid, rec in sw_map.items():
        if not rec:
            continue
        release = (rec.get("releaseFormat1") or "").strip()
        raw_os  = (rec.get("softwareType")   or "").strip()
        ostype  = _SW_TYPE_TO_OSTYPE.get(raw_os.upper(), raw_os)
        if not release:
            continue
        pair_to_pids.setdefault((release, ostype), []).append(pid)
        pair_release_date.setdefault(
            (release, ostype), _parse_sw_release_date(rec.get("releaseDate") or ""))

    pairs = list(pair_to_pids.keys())
    cache: dict[tuple[str, str], dict] = {}

    for idx, (release, ostype) in enumerate(pairs):
        eox_release = _to_eox_release_fmt(release, ostype)
        input1 = f"{eox_release},{ostype}" if ostype else eox_release
        r = _get_with_retry(
            f"{EOX_BASE}/EOXBySWReleaseString/1/",
            headers,
            params={"input1": input1, "responseencoding": "json"},
        )
        if r and r.ok:
            for rec in r.json().get("EOXRecord", []):
                if rec.get("EOXError"):
                    continue
                eovss = rec.get("EndOfSecurityVulSupportDate", {}).get("value") or ""
                ldos  = rec.get("LastDateOfSupport", {}).get("value") or ""
                if not (eovss or ldos):
                    continue
                if not _sw_eox_applies(eovss, ldos, pair_release_date.get((release, ostype))):
                    continue
                cache[(release, ostype)] = {"_sw_eovss": eovss, "_sw_ldos": ldos}
                break
        if on_progress:
            on_progress(idx + 1, len(pairs))

    for (release, ostype), extra in cache.items():
        for pid in pair_to_pids[(release, ostype)]:
            if sw_map.get(pid):
                sw_map[pid].update(extra)


# ── Text helpers ───────────────────────────────────────────────────────────────

def _spaced(s: str) -> str:
    """Insert spaces into concatenated product strings (e.g. CiscoCatalyst1000SeriesSwitches)."""
    if not s or ' ' in s:
        return s
    s = re.sub(r'([a-z])([A-Z])',   r'\1 \2', s)   # camelCase boundary
    s = re.sub(r'([A-Za-z])(\d)',   r'\1 \2', s)   # letter → digit
    s = re.sub(r'(\d)([A-Za-z])',   r'\1 \2', s)   # digit → letter
    return s


# ── EoX helpers ────────────────────────────────────────────────────────────────

def _date_val(record: dict, key: str) -> str:
    return record.get(key, {}).get("value", "") or ""


def _eox_status(record: dict) -> str:
    error = record.get("EOXError") or {}
    eid   = error.get("ErrorID", "")
    if eid == "SSA_ERR_021_Pid":
        return "Invalid SKU"
    if eid == "SSA_ERR_026" or not record.get("EndOfSaleDate", {}).get("value"):
        return "Active"
    now = time.time()
    def ts(key):
        v = record.get(key, {}).get("value", "")
        try:
            return datetime.datetime.fromisoformat(v).timestamp() if v else None
        except Exception:
            return None
    if (t := ts("LastDateOfSupport")) and t < now:
        return "End of Life"
    return "End of Sale"


def _resolved_status(eox_rec: dict, info_rec: dict) -> str:
    """Determine EoX status, cross-checked against Product Info.

    The EoX API returns SSA_ERR_021_Pid ('Invalid SKU') or no record at all for
    products it doesn't cover (e.g. SMB/SG series).  If Product Info confirms
    the product exists, those cases are promoted to 'Active'.
    """
    has_info = bool(info_rec and info_rec.get("product_name"))
    if not eox_rec:
        return "Active" if has_info else "No Record"
    status = _eox_status(eox_rec)
    if status in ("Invalid SKU", "No Record") and has_info:
        return "Active"
    return status


# ── PID normalisation fallback ────────────────────────────────────────────────

# Suffixes Cisco appends to a base PID without changing the hardware itself:
# licence level (-A/-E/-M), region (-EU/-NA/…), packaging (-RF/-WS).  The EoX and
# Product Info APIs only know the base PID, so an inventory full of real, orderable
# SKUs like C9300-48P-M or SG250X-48P-K9-EU comes back as "Invalid SKU".
_PID_TRIM_SUFFIXES = {
    "A", "E", "M", "K9", "RF", "WS",
    "EU", "NA", "UK", "CN", "IN", "JP", "AU", "BR", "AR", "KR", "G5",
}

# Suffixes the APIs *require* but inventories often drop (WS-C3750-48TS is stored
# as such although only WS-C3750-48TS-S / -E exist).  Ordered by likelihood.
_PID_APPEND_SUFFIXES = ("S", "E", "L", "A", "M", "K9")


def _eox_wildcard(token: str, prefix: str) -> dict[str, dict]:
    """Query EOXByProductID with a trailing wildcard: dict[EOLProductID, record]."""
    headers = {"Authorization": f"Bearer {token}"}
    r = _get_with_retry(f"{EOX_BASE}/EOXByProductID/1/{requests.utils.quote(prefix + '*')}",
                        headers, params={"responseencoding": "json"})
    out: dict[str, dict] = {}
    if r and r.ok:
        try:
            for rec in r.json().get("EOXRecord", []):
                key = (rec.get("EOLProductID") or "").strip().upper()
                if key and key not in out:
                    out[key] = rec
        except Exception:
            pass
    return out


def _eox_by_wildcard_exact(token: str, pid: str, max_trims: int = 2) -> dict:
    """Find `pid`'s own EoX record through a wildcard query, or {}.

    The exact-match path of EOXByProductID rejects PIDs the same API returns in
    full under a wildcard: SG250X-48P-K9-EU answers SSA_ERR_021_Pid directly, yet
    "SG250X-48P*" lists it with its real dates.  Progressively shorter prefixes
    are tried because the wildcard has to start *before* the part the exact
    lookup chokes on.

    Only a record whose EOLProductID equals `pid` is accepted.  Near-matches are
    deliberately refused — "C9500-24Y4C-E*" also returns the promo bundle
    C9500-24Y4C-EP-3, whose End-of-Life dates have nothing to do with the switch.
    """
    prefix = (pid or "").strip().upper()
    for _ in range(max_trims + 1):
        if not prefix:
            break
        found = _eox_wildcard(token, prefix)
        if pid in found:
            return found[pid]
        head, sep, _tail = prefix.rpartition("-")
        if not sep or not head:
            break
        prefix = head
    return {}


def _pid_variants(pid: str, max_trims: int = 2) -> list[str]:
    """Ordered candidate PIDs to try when the APIs don't know `pid`.

    Trims come first (same hardware, different licence/region) and only then
    appends (a truncated PID completed with its licence suffix).
    """
    pid = (pid or "").strip().upper()
    if not pid:
        return []
    out: list[str] = []
    seen = {pid}

    cur = pid
    for _ in range(max_trims):
        head, sep, tail = cur.rpartition("-")
        if not sep or not head or tail not in _PID_TRIM_SUFFIXES:
            break
        cur = head
        if cur not in seen:
            seen.add(cur)
            out.append(cur)

    for suf in _PID_APPEND_SUFFIXES:
        cand = f"{pid}-{suf}"
        if cand not in seen:
            seen.add(cand)
            out.append(cand)

    return out


def resolve_pid_aliases(token: str, pids: list[str], eox_map: dict, info_map: dict,
                        on_progress=None) -> dict[str, str]:
    """Second-chance lookup for PIDs both APIs rejected.

    Queries _pid_variants() for every unresolved PID and, on the first candidate
    the APIs do know, copies its records under the queried PID with an `_alias_of`
    marker (surfaced in the "PID résolu" column, so nothing is substituted
    silently).  Records are *copied*, never shared: enrich_replacement_chains
    writes _final_replacement into eox_map[queried_pid] and would otherwise
    corrupt the alias target's record when both are in the same batch.

    Returns dict[queried_pid_upper, resolved_pid].
    """
    unknown = [
        p for p in pids
        if _resolved_status(eox_map.get(p.upper()), info_map.get(p.upper(), {}))
        in ("Invalid SKU", "No Record")
    ]
    if not unknown:
        return {}

    cand_map = {p.upper(): _pid_variants(p) for p in unknown}

    to_query: list[str] = []
    seen: set[str] = set()
    for cands in cand_map.values():
        for c in cands:
            if c not in seen and c not in eox_map and c not in info_map:
                seen.add(c)
                to_query.append(c)

    cand_eox:  dict = fetch_eox(token, to_query)          if to_query else {}
    cand_info: dict = fetch_product_info(token, to_query) if to_query else {}

    aliases: dict[str, str] = {}
    for idx, p in enumerate(unknown):
        u = p.upper()

        # 1. Wildcard lookup — recovers the PID's own record whenever only the
        #    exact-match endpoint failed.  No alias: this really is the PID.
        if (rec := _eox_by_wildcard_exact(token, u)):
            eox_map[u] = dict(rec)
            if on_progress:
                on_progress(idx + 1, len(unknown))
            continue

        # 2. Licence/region trimming and suffix completion.
        for c in cand_map[u]:
            e = cand_eox.get(c)  or eox_map.get(c)
            i = cand_info.get(c) or info_map.get(c, {})
            if _resolved_status(e, i) in ("Invalid SKU", "No Record"):
                continue
            if e:
                rec = dict(e)
                rec["_alias_of"] = c
                eox_map[u] = rec
            if i:
                rec = dict(i)
                rec["_alias_of"] = c
                info_map[u] = rec
            aliases[u] = c
            break
        if on_progress:
            on_progress(idx + 1, len(unknown))

    return aliases


# ── Replacement chain resolution ──────────────────────────────────────────────

def _migration_next(rec: dict) -> str:
    """Return the next-hop PID from an EoX record's MigrationDetails.

    Prefers MigrationProductId (a clean PID); falls back to MigrationProductName
    when MigrationProductId is absent (Cisco sometimes puts the PID there instead,
    e.g. WS-C3850-12S-S → MigrationProductName="C9300-24S").
    """
    md = rec.get("EOXMigrationDetails") or {}
    return ((md.get("MigrationProductId") or md.get("MigrationProductName") or "").strip().upper())


def _resolve_replacement_chain(token: str, eox_map: dict, start_pid: str, max_depth: int = 8) -> str:
    """Follow MigrationProductId links until reaching a product with no replacement.

    The BFS in enrich_replacement_chains pre-fetches most nodes, but this
    function keeps an on-demand fallback so stale/missing cache entries don't
    truncate the chain prematurely.
    """
    seen = {start_pid.upper()}
    current = start_pid.upper()

    for _ in range(max_depth):
        if current not in eox_map:
            fetched = fetch_eox(token, [current])
            eox_map.update(fetched)
            # If the API returned a record under a different canonical key
            # (e.g. queried "WS-C2960-8TC-L", got EOLProductID "WS-C2960C-8TC-L"),
            # do NOT alias the wrong record under `current`. Just stop here —
            # the chain cannot be followed further without risking wrong data.
            if current not in eox_map:
                break
        e = eox_map.get(current, {})
        nxt = _migration_next(e)
        if not nxt or nxt in seen:
            break
        seen.add(nxt)
        current = nxt

    return current


def enrich_replacement_chains(token: str, pids: list[str], eox_map: dict, on_progress=None):
    """Resolve the full replacement chain for every PID that has a MigrationProductId.

    Uses BFS to batch-fetch all replacement chain nodes first, so the per-PID
    resolution phase is entirely cache-based (no individual network calls).
    Stores the resolved final PID in eox_map[pid]['_final_replacement'].
    """
    candidates = [
        (pid.upper(), _migration_next(eox_map.get(pid.upper(), {})))
        for pid in pids
    ]
    candidates = [(p, r) for p, r in candidates if r]

    if not candidates:
        return

    # BFS: pre-fetch EoX records for every node in the replacement chains.
    # `visited` tracks PIDs whose MigrationProductId has already been examined.
    # We intentionally do NOT pre-seed visited from eox_map — PIDs already cached
    # (e.g. because they were in the original input) still need their downstream
    # nodes fetched and visited.
    visited: set[str] = set()
    frontier = {r for _, r in candidates}
    for _ in range(8):
        if not frontier:
            break
        to_fetch = [p for p in frontier if p not in eox_map]
        if to_fetch:
            eox_map.update(fetch_eox(token, to_fetch))
        visited.update(frontier)
        next_frontier: set[str] = set()
        for p in frontier:
            nxt = _migration_next(eox_map.get(p, {}))
            if nxt and nxt not in visited:
                next_frontier.add(nxt)
        frontier = next_frontier

    # Resolve chains — now cache-only
    total = len(candidates)
    for idx, (pid, imm) in enumerate(candidates):
        final = _resolve_replacement_chain(token, eox_map, imm)
        eox_map[pid]["_final_replacement"] = final
        if on_progress:
            on_progress(idx + 1, total)


# ── Excel column catalogue ─────────────────────────────────────────────────────

ALL_COLUMNS: list[tuple[str, any]] = [
    ("Product Name",          lambda p, e, i, s: _spaced(i.get("product_name") or "")                                  if i else ""),
    ("Product Description",   lambda p, e, i, s: (e.get("ProductIDDescription") or "")                                 if e else ""),
    ("Catégorie",             lambda p, e, i, s: (i.get("product_category") or "")                                     if i else ""),
    ("Product Series",        lambda p, e, i, s: _spaced(i.get("product_series") or "")                                if i else ""),
    ("Release Date",          lambda p, e, i, s: (i.get("release_date") or "")                                         if i else ""),
    ("EoX Status",            lambda p, e, i, s: _resolved_status(e, i)),
    ("PID résolu",            lambda p, e, i, s: (e.get("_alias_of") if e else "") or (i.get("_alias_of") if i else "") or ""),
    ("End of Sale",           lambda p, e, i, s: _date_val(e, "EndOfSaleDate")                                         if e else ""),
    ("End of SW Maintenance", lambda p, e, i, s: _date_val(e, "EndOfSWMaintenanceReleases")                            if e else ""),
    ("Last Date of Support",  lambda p, e, i, s: "N/A" if _resolved_status(e, i) == "Active" else (_date_val(e, "LastDateOfSupport")                                                                                                                    if e else "")),
    ("End of Vuln Support",   lambda p, e, i, s: "N/A" if _resolved_status(e, i) == "Active" else (_date_val(e, "EndOfSecurityVulSupportDate")                                                                                                           if e else "")),
    ("Replacement PID",       lambda p, e, i, s: "N/A" if _resolved_status(e, i) == "Active" else ((e.get("_final_replacement") or e.get("EOXMigrationDetails", {}).get("MigrationProductId") or e.get("EOXMigrationDetails", {}).get("MigrationProductName") or "") if e else "")),
    ("Bulletin URL",          lambda p, e, i, s: (e.get("LinkToProductBulletinURL") or "")                             if e else ""),
    ("SW Type",               lambda p, e, i, s: (s.get("softwareType") or "")                                         if s else ""),
    ("Suggested Release",     lambda p, e, i, s: (s.get("releaseFormat1") or "")                                       if s else ""),
    ("Train",                 lambda p, e, i, s: (s.get("relDispName") or "")                                          if s else ""),
    ("Lifecycle",             lambda p, e, i, s: (s.get("releaseLifeCycle") or "")                                     if s else ""),
    ("SW Release Date",       lambda p, e, i, s: (s.get("releaseDate") or "")                                          if s else ""),
    ("SW End of Vuln Support",lambda p, e, i, s: (s.get("_sw_eovss") or "")                                            if s else ""),
    ("SW Last Date of Support",lambda p, e, i, s: (s.get("_sw_ldos") or "")                                            if s else ""),
    ("Manual Hardware Recommendation", lambda p, e, i, s: ""),  # placeholder — replaced in build_excel with mappings closure
]

_ALL_COLUMNS_MAP  = {name: fn for name, fn in ALL_COLUMNS}
ALL_COLUMN_NAMES  = [name for name, _ in ALL_COLUMNS]

_COL_WIDTHS = {
    "Product Name": 36, "Product Description": 52, "Catégorie": 16, "Product Series": 28,
    "Release Date": 14, "EoX Status": 14, "PID résolu": 20, "End of Sale": 14,
    "End of SW Maintenance": 22, "Last Date of Support": 20,
    "Replacement PID": 22, "Bulletin URL": 50, "End of Vuln Support": 20,
    "SW Type": 20, "Suggested Release": 18, "Train": 22,
    "Lifecycle": 12, "SW Release Date": 16,
    "SW End of Vuln Support": 22, "SW Last Date of Support": 24,
    "Manual Hardware Recommendation": 28,
}

_CENTER_COLS = {
    "EoX Status", "PID résolu", "Release Date", "End of Sale",
    "End of SW Maintenance", "Last Date of Support", "End of Vuln Support", "SW Release Date",
}

# ── Excel builder ──────────────────────────────────────────────────────────────

_BLUE   = "0071E3"
_WHITE  = "FFFFFF"
_LIGHT  = "F5F5F7"

_STATUS_STYLE = {
    "End of Life":  (PatternFill("solid", fgColor="FFEBEB"), Font(color="C0392B", bold=True)),
    "End of Sale":  (PatternFill("solid", fgColor="FFF3E0"), Font(color="C96000", bold=True)),
    "Active":       (PatternFill("solid", fgColor="EAFAF1"), Font(color="1A8A3A", bold=True)),
    "Invalid SKU":  (PatternFill("solid", fgColor="F5F5F7"), Font(color="C0392B", italic=True)),
    "No Record":    (PatternFill("solid", fgColor="F5F5F7"), Font(color="86868B")),
}

_THIN_BORDER = Border(
    left=Side(style="thin", color="E5E5EA"),
    right=Side(style="thin", color="E5E5EA"),
    top=Side(style="thin", color="E5E5EA"),
    bottom=Side(style="thin", color="E5E5EA"),
)


def _header_cell(ws, row, col, value):
    c = ws.cell(row=row, column=col, value=value)
    c.fill   = PatternFill("solid", fgColor=_BLUE)
    c.font   = Font(color=_WHITE, bold=True, size=11)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = _THIN_BORDER
    return c


def _data_cell(ws, row, col, value, center=False):
    c = ws.cell(row=row, column=col, value=value)
    c.alignment = Alignment(horizontal="center" if center else "left", vertical="center")
    c.border = _THIN_BORDER
    c.font   = Font(size=11)
    return c


def build_excel(pids: list[str], eox: dict, info: dict, sw: dict | None,
                source_wb: openpyxl.Workbook | None = None,
                col_spec: list[tuple[str, str]] | None = None,
                mappings: dict | None = None) -> io.BytesIO:
    """
    col_spec: ordered list of (canonical_name, display_label) or None for all defaults.
    canonical_name is used for data extraction and styling; display_label for headers.
    mappings: optional dict[str, str] of upper-cased base PID → override PID for
              the "Manual Hardware Recommendation" column.
    """
    _mappings: dict = mappings or {}

    def _manual_hw_fn(p, e, i, s):
        status = _resolved_status(e, i)
        if status == "Active":
            base = p
        else:
            base = ((e.get("_final_replacement") or
                     e.get("EOXMigrationDetails", {}).get("MigrationProductId") or
                     e.get("EOXMigrationDetails", {}).get("MigrationProductName") or
                     "") if e else "")
        if not base:
            return ""
        return _mappings.get(base.upper(), base)

    if col_spec:
        COLUMNS = []
        for cn, lbl in col_spec:
            if cn == "__blank__":
                COLUMNS.append(("__blank__", lbl, lambda p, e, i, s: ""))
            elif cn == "Manual Hardware Recommendation":
                COLUMNS.append((cn, lbl, _manual_hw_fn))
            elif cn in _ALL_COLUMNS_MAP:
                COLUMNS.append((cn, lbl, _ALL_COLUMNS_MAP[cn]))
    else:
        COLUMNS = []
        for cn, fn in ALL_COLUMNS:
            if cn == "Manual Hardware Recommendation":
                COLUMNS.append((cn, cn, _manual_hw_fn))
            else:
                COLUMNS.append((cn, cn, fn))

    if source_wb:
        wb = source_wb
        ws = wb.active
        first_val  = str(ws.cell(row=1, column=1).value or "").strip().lower()
        has_header = first_val in {"sku", "pid", "product id", "part number", "part no", "model"}
        start_col  = ws.max_column + 1

        if has_header:
            for offset, (cn, lbl, _) in enumerate(COLUMNS):
                _header_cell(ws, 1, start_col + offset, lbl)

        for row_idx, pid in enumerate(pids, start=2 if has_header else 1):
            e = eox.get(pid.upper(), {})
            i = info.get(pid.upper(), {})
            s = (sw or {}).get(pid.upper(), {})
            for offset, (cn, lbl, extractor) in enumerate(COLUMNS):
                val = extractor(pid, e, i, s)
                c   = _data_cell(ws, row_idx, start_col + offset, val, center=cn in _CENTER_COLS)
                if cn == "EoX Status":
                    fill, font = _STATUS_STYLE.get(val, _STATUS_STYLE["No Record"])
                    c.fill = fill; c.font = Font(color=font.color, bold=font.bold, italic=font.italic, size=11)

    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "EoX Lookup"

        all_cols = [("Product ID", "Product ID", None)] + COLUMNS
        ws.row_dimensions[1].height = 32
        for col_idx, (cn, lbl, _) in enumerate(all_cols, start=1):
            _header_cell(ws, 1, col_idx, lbl)

        for row_idx, pid in enumerate(pids, start=2):
            ws.row_dimensions[row_idx].height = 18
            e = eox.get(pid.upper(), {})
            i = info.get(pid.upper(), {})
            s = (sw or {}).get(pid.upper(), {})

            _data_cell(ws, row_idx, 1, pid)
            for offset, (cn, lbl, extractor) in enumerate(COLUMNS, start=1):
                val = extractor(pid, e, i, s)
                c   = _data_cell(ws, row_idx, 1 + offset, val, center=cn in _CENTER_COLS)
                if cn == "EoX Status":
                    fill, font = _STATUS_STYLE.get(val, _STATUS_STYLE["No Record"])
                    c.fill = fill; c.font = Font(color=font.color, bold=font.bold, italic=font.italic, size=11)

        widths = {"Product ID": 18, **_COL_WIDTHS}
        for col_idx, (cn, lbl, _) in enumerate(all_cols, start=1):
            ws.column_dimensions[ws.cell(1, col_idx).column_letter].width = widths.get(cn, 18)

        ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Excel input parser ─────────────────────────────────────────────────────────

_SKU_HEADERS = {"sku", "pid", "product id", "part number", "part no", "partnumber", "model"}


# ── Excel comparison ───────────────────────────────────────────────────────────

def _parse_excel_for_compare(file) -> tuple[dict[str, dict], list[str]]:
    """Parse an EoL-generated Excel file into {pid_upper: {col: val}} + header list."""
    wb = openpyxl.load_workbook(file, data_only=True)
    ws = wb.active
    raw = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
    headers = [str(c or "").strip() for c in raw]
    while headers and not headers[-1]:
        headers.pop()

    data: dict[str, dict] = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row:
            continue
        pid = str(row[0] or "").strip().upper()
        if not pid:
            continue
        row_data: dict[str, str] = {}
        for i, col in enumerate(headers):
            if i == 0:
                continue  # Product ID column itself
            val = row[i] if i < len(row) else None
            row_data[col] = str(val or "").strip()
        data[pid] = row_data

    return data, headers


def read_pids_from_excel(file) -> tuple[list[str], openpyxl.Workbook]:
    wb = openpyxl.load_workbook(file)
    ws = wb.active
    header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
    sku_col, start_row = 1, 1
    for col_idx, val in enumerate(header_row, start=1):
        if str(val or "").strip().lower() in _SKU_HEADERS:
            sku_col, start_row = col_idx, 2
            break
    pids = []
    for row in ws.iter_rows(min_row=start_row, min_col=sku_col, max_col=sku_col, values_only=True):
        val = str(row[0] or "").strip()
        if val:
            pids.append(val)
    return pids, wb


# ── Job stats ──────────────────────────────────────────────────────────────────

def compute_stats(pids: list[str], eox_map: dict, info_map: dict, sw_map: dict) -> dict:
    eox_counts:         dict[str, int]  = {}
    eox_pids:           dict[str, list] = {}
    replacement_counts: dict[str, int]  = {}
    info_count = sw_count = 0
    col_counts: dict[str, int] = {}

    for pid in pids:
        u = pid.upper()
        e = eox_map.get(u)
        i = info_map.get(u, {})
        s = sw_map.get(u, {})

        status = _resolved_status(e, i)
        eox_counts[status] = eox_counts.get(status, 0) + 1
        eox_pids.setdefault(status, []).append(pid)

        if e:
            # Mirror the "Replacement PID" column logic exactly so the count
            # matches what the Excel contains (not just _final_replacement).
            _md = e.get("EOXMigrationDetails") or {}
            final = (
                e.get("_final_replacement") or
                _md.get("MigrationProductId") or
                _md.get("MigrationProductName") or
                ""
            ).strip()
            if final:
                key = final.upper()  # normalise case to avoid dup unique entries
                replacement_counts[key] = replacement_counts.get(key, 0) + 1

        if i.get("product_name"):
            info_count += 1
        if s.get("releaseFormat1"):
            sw_count += 1

        for cn, fn in ALL_COLUMNS:
            if cn == "__blank__":
                continue
            val = fn(pid, e or {}, i, s)
            if val:
                col_counts[cn] = col_counts.get(cn, 0) + 1

    return {
        "total":         len(pids),
        "eox":           eox_counts,
        "eox_pids":      eox_pids,
        "replacements":  replacement_counts,
        "product_info":  info_count,
        "sw":            sw_count,
        "cols":          col_counts,
    }


# ── Background job runner ──────────────────────────────────────────────────────

def _emit(job_id: str, event: str, data: str):
    with _job_lock:
        _job_events[job_id].append((event, data))


def run_job(job_id: str, client_id: str, client_secret: str,
            pids: list[str], source_wb, col_spec, download_name: str,
            hw_mappings: dict | None = None):
    try:
        _emit(job_id, "progress", _json.dumps({"msg": "Authenticating…", "pct": 5}))

        try:
            token = get_access_token(client_id, client_secret)
        except Exception as e:
            _emit(job_id, "error", _json.dumps({"error": f"Authentication failed: {e}"}))
            with _job_lock:
                _jobs[job_id]["status"] = "error"
            return

        _emit(job_id, "progress", _json.dumps({"msg": "Querying Cisco APIs…", "pct": 5}))

        eox_map: dict  = {}
        info_map: dict = {}
        sw_map: dict   = {}
        total = len(pids)

        # ── Per-API granular progress tracker ─────────────────────────────
        _api_done_count = {"eox": 0, "info": 0, "sw": 0}
        _api_complete   = set()
        _prog_lock      = threading.Lock()

        def _make_reporter(api: str):
            def report(done: int):
                with _prog_lock:
                    _api_done_count[api] = done
                    if done >= total:
                        _api_complete.add(api)
                    fraction = sum(_api_done_count[k] for k in ("eox", "info", "sw")) / (total * 3)
                    pct = int(5 + fraction * 80)
                    parts = []
                    for k, label in (("eox", "EoX"), ("info", "Info"), ("sw", "SW")):
                        n = _api_done_count[k]
                        parts.append(f"{label} {n}/{total}" if k not in _api_complete else f"{label} ✓")
                    _emit(job_id, "progress", _json.dumps({
                        "msg": "  ·  ".join(parts),
                        "pct": pct,
                        "completed": list(_api_complete),
                    }))
            return report

        with ThreadPoolExecutor(max_workers=3) as pool:
            futures = {
                pool.submit(fetch_eox,            token, pids, _make_reporter("eox")):  "eox",
                pool.submit(fetch_product_info,   token, pids, _make_reporter("info")): "info",
                pool.submit(fetch_sw_suggestions, token, pids, _make_reporter("sw")):   "sw",
            }
            for fut in as_completed(futures):
                name = futures[fut]
                try:
                    result = fut.result()
                    if name == "eox":
                        eox_map  = result
                    elif name == "info":
                        info_map = result or {}
                    elif name == "sw":
                        if result is None:
                            with _prog_lock:
                                _api_complete.add("sw")
                                _api_done_count["sw"] = total
                            _emit(job_id, "progress", _json.dumps({
                                "msg": "SW subscription inactive",
                                "pct": int(5 + sum(_api_done_count[k] for k in ("eox","info","sw")) / (total*3) * 80),
                                "completed": list(_api_complete),
                            }))
                        else:
                            sw_map = result
                except Exception as exc:
                    _emit(job_id, "progress", _json.dumps({
                        "msg": f"{name} unavailable: {exc}",
                        "pct": int(5 + sum(_api_done_count[k] for k in ("eox","info","sw")) / (total*3) * 80),
                        "completed": list(_api_complete),
                    }))

        # ── PID normalisation fallback ────────────────────────────────────
        # Licence/region variants (C9300-48P-M, SG250X-48P-K9-EU) and truncated
        # PIDs (WS-C3750-48TS) are unknown to both APIs; retry them on their base
        # PID before anything downstream treats them as invalid.
        _alias_candidates = sum(
            1 for p in pids
            if _resolved_status(eox_map.get(p.upper()), info_map.get(p.upper(), {}))
            in ("Invalid SKU", "No Record")
        )
        aliases: dict = {}
        if _alias_candidates:
            def alias_reporter(done: int, total: int):
                _emit(job_id, "progress", _json.dumps({
                    "msg": f"Résolution des PID inconnus… {done}/{total}",
                    "pct": 85,
                    "completed": list(_api_complete),
                }))
            _emit(job_id, "progress", _json.dumps({
                "msg": f"Résolution des PID inconnus ({_alias_candidates})…",
                "pct": 85,
                "completed": list(_api_complete),
            }))
            aliases = resolve_pid_aliases(token, pids, eox_map, info_map,
                                          on_progress=alias_reporter)
            if aliases:
                _emit(job_id, "progress", _json.dumps({
                    "msg": f"PID résolus : +{len(aliases)}",
                    "pct": 86,
                    "completed": list(_api_complete),
                }))

        # Pull the SW suggestion of the resolved PID for every alias.
        if aliases and sw_map is not None:
            _alias_sw = fetch_sw_suggestions(token, sorted(set(aliases.values()))) or {}
            for _orig, _target in aliases.items():
                if _orig not in sw_map and _alias_sw.get(_target):
                    sw_map[_orig] = dict(_alias_sw[_target])

        if sw_map is not None:
            _suffix_candidates = sum(
                1 for p in pids
                if p.upper() not in sw_map and info_map.get(p.upper(), {}).get("product_name")
            )
            def suffix_reporter(done: int, total: int):
                pct = 86 + int(done / max(total, 1) * 3)
                _emit(job_id, "progress", _json.dumps({
                    "msg": f"SW suffix fallback… {done}/{total}",
                    "pct": pct,
                    "completed": list(_api_complete),
                }))
            if _suffix_candidates:
                _emit(job_id, "progress", _json.dumps({
                    "msg": f"SW suffix fallback ({_suffix_candidates} candidates)…",
                    "pct": 86,
                    "completed": list(_api_complete),
                }))
            n = _sw_suffix_fallback(token, pids, info_map, sw_map,
                                    on_progress=suffix_reporter if _suffix_candidates else None)
            if n:
                _emit(job_id, "progress", _json.dumps({
                    "msg": f"SW suffix: +{n} resolved",
                    "pct": 89,
                    "completed": list(_api_complete),
                }))

        if sw_map:
            _sw_eox_pairs = sum(
                1 for rec in sw_map.values()
                if rec and (rec.get("releaseFormat1") or "").strip()
            )
            _sw_eox_unique = len({
                (rec.get("releaseFormat1","").strip(), rec.get("softwareType","").strip())
                for rec in sw_map.values() if rec and rec.get("releaseFormat1","").strip()
            })
            if _sw_eox_unique:
                _emit(job_id, "progress", _json.dumps({
                    "msg": f"Fetching SW EoX dates… 0/{_sw_eox_unique}",
                    "pct": 89,
                    "completed": list(_api_complete),
                }))
                def sw_eox_reporter(done: int, total: int):
                    pct = 89 + int(done / max(total, 1) * 3)
                    _emit(job_id, "progress", _json.dumps({
                        "msg": f"Fetching SW EoX dates… {done}/{total}",
                        "pct": pct,
                        "completed": list(_api_complete),
                    }))
                enrich_sw_eox(token, sw_map, on_progress=sw_eox_reporter)

        chain_candidates = [
            p for p in pids
            if (eox_map.get(p.upper(), {}).get("EOXMigrationDetails", {}).get("MigrationProductId") or "").strip()
        ]
        chain_total = len(chain_candidates)
        _emit(job_id, "progress", _json.dumps({
            "msg": f"Resolving replacement chains… 0/{chain_total}" if chain_total else "Resolving replacement chains…",
            "pct": 92,
            "completed": list(_api_complete),
        }))

        def chain_reporter(done: int, total: int):
            pct = 92 + int(done / max(total, 1) * 3)
            _emit(job_id, "progress", _json.dumps({
                "msg": f"Resolving replacement chains… {done}/{total}",
                "pct": pct,
                "completed": list(_api_complete),
            }))

        enrich_replacement_chains(token, pids, eox_map, on_progress=chain_reporter)

        _emit(job_id, "progress", _json.dumps({"msg": "Building Excel file…", "pct": 92}))
        buf   = build_excel(pids, eox_map, info_map, sw_map or None, source_wb, col_spec, mappings=hw_mappings)
        stats = compute_stats(pids, eox_map, info_map, sw_map or {})

        # Persist the Excel file so it can be re-downloaded from history
        dl_path = os.path.join(_DOWNLOADS_DIR, f"{job_id}.xlsx")
        buf.seek(0)
        with open(dl_path, "wb") as fh:
            fh.write(buf.read())
        buf.seek(0)

        with _job_lock:
            _jobs[job_id].update({"status": "done", "buf": buf, "download_name": download_name})

        _emit(job_id, "done", _json.dumps({
            "download_url": f"/api/download/{job_id}",
            "label": download_name,
            "stats": stats,
            "job_id": job_id,
        }))

    except Exception as e:
        _emit(job_id, "error", _json.dumps({"error": str(e)}))
        with _job_lock:
            _jobs[job_id]["status"] = "error"


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html",
                           default_client_id=os.getenv("CLIENT_ID", ""),
                           default_client_secret=os.getenv("CLIENT_SECRET", ""),
                           all_columns=ALL_COLUMN_NAMES)


@app.route("/api/start", methods=["POST"])
def start_job_route():
    client_id     = (request.form.get("client_id") or "").strip()
    client_secret = (request.form.get("client_secret") or "").strip()
    skus_text     = (request.form.get("skus") or "").strip()
    file          = request.files.get("file")

    if not client_id or not client_secret:
        return jsonify({"error": "Client ID and Client Secret are required."}), 400
    if not skus_text and not file:
        return jsonify({"error": "Provide SKUs as text or upload an Excel file."}), 400

    source_wb = None
    if file and file.filename:
        if not file.filename.lower().endswith((".xlsx", ".xlsm")):
            return jsonify({"error": "Please upload a .xlsx or .xlsm file."}), 400
        try:
            pids, source_wb = read_pids_from_excel(file)
        except Exception as e:
            return jsonify({"error": f"Could not read Excel file: {e}"}), 400
        _ts = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        download_name = f"{os.path.splitext(file.filename)[0]}_eox_{_ts}.xlsx"
    else:
        pids = [s.strip() for s in skus_text.replace("\n", ",").split(",") if s.strip()]
        _ts = datetime.datetime.now().strftime("%Y%m%d_%H%M")
        download_name = f"eox_results_{_ts}.xlsx"

    if not pids:
        return jsonify({"error": "No SKUs found."}), 400

    col_spec = None
    raw_cols = (request.form.get("columns") or "").strip()
    if raw_cols:
        try:
            data = _json.loads(raw_cols)
            if isinstance(data, list) and data:
                if isinstance(data[0], dict):
                    # New format: [{col: "...", label: "..."}, ...]
                    col_spec = [(c["col"], c.get("label") or c["col"]) for c in data if "col" in c]
                else:
                    # Legacy format: ["col_name", ...]
                    col_spec = [(c, c) for c in data if isinstance(c, str)]
        except Exception:
            pass

    # Load overrides for the active preset from the server-side file
    hw_mappings = None
    try:
        raw_preset_id = (request.form.get("preset_id") or "").strip() or "default"
        all_overrides = _read_json(_OVERRIDES_F, {})
        preset_overrides = all_overrides.get(raw_preset_id, {})
        if isinstance(preset_overrides, dict) and preset_overrides:
            hw_mappings = {k.upper(): v for k, v in preset_overrides.items()
                           if isinstance(k, str) and isinstance(v, str)}
    except Exception:
        pass

    job_id = uuid.uuid4().hex[:12]
    with _job_lock:
        _jobs[job_id]       = {"status": "running"}
        _job_events[job_id] = []

    t = threading.Thread(
        target=run_job,
        args=(job_id, client_id, client_secret, pids, source_wb, col_spec, download_name),
        kwargs={"hw_mappings": hw_mappings},
        daemon=True,
    )
    t.start()

    return jsonify({"job_id": job_id})


@app.route("/api/progress/<job_id>")
def job_progress(job_id):
    if job_id not in _jobs:
        return jsonify({"error": "Job not found"}), 404

    def event_stream():
        seen = 0
        last_ping = time.time()
        while True:
            with _job_lock:
                events = list(_job_events.get(job_id, []))
                status = _jobs.get(job_id, {}).get("status", "running")

            while seen < len(events):
                event, data = events[seen]
                yield f"event: {event}\ndata: {data}\n\n"
                seen += 1
                last_ping = time.time()

            if status in ("done", "error"):
                break

            # Heartbeat every 20s so browsers/proxies don't close the connection
            if time.time() - last_ping > 20:
                yield ": heartbeat\n\n"
                last_ping = time.time()

            time.sleep(0.2)

    return Response(
        stream_with_context(event_stream()),
        mimetype="text/event-stream",
        headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"},
    )


@app.route("/api/download/<job_id>")
def download_job(job_id):
    with _job_lock:
        job = _jobs.get(job_id)

    if not job or job.get("status") != "done":
        return jsonify({"error": "Job not ready or not found"}), 404

    buf           = job["buf"]
    download_name = job["download_name"]

    # Clean up after serving
    with _job_lock:
        _jobs.pop(job_id, None)
        _job_events.pop(job_id, None)

    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=download_name,
    )


@app.route("/api/inspect", methods=["POST"])
def inspect_sku_route():
    client_id     = (request.form.get("client_id") or "").strip()
    client_secret = (request.form.get("client_secret") or "").strip()
    sku           = (request.form.get("sku") or "").strip().upper()

    if not client_id or not client_secret:
        return jsonify({"error": "Credentials required"}), 400
    if not sku:
        return jsonify({"error": "SKU required"}), 400

    try:
        token = get_access_token(client_id, client_secret)
    except Exception as e:
        return jsonify({"error": f"Authentication failed: {e}"}), 401

    with ThreadPoolExecutor(max_workers=3) as pool:
        fut_eox  = pool.submit(fetch_eox,            token, [sku])
        fut_info = pool.submit(fetch_product_info,   token, [sku])
        fut_sw   = pool.submit(fetch_sw_suggestions, token, [sku])
        eox_data  = fut_eox.result()
        info_data = fut_info.result()
        sw_data   = fut_sw.result()

    e = eox_data.get(sku, {})
    i = info_data.get(sku, {})
    sw_map = dict(sw_data) if sw_data else {}
    if sw_data is not None:
        _sw_suffix_fallback(token, [sku], info_data, sw_map)
    s = sw_map.get(sku, {})

    alias = resolve_pid_aliases(token, [sku], eox_data, info_data)
    if alias:
        e = eox_data.get(sku, {})
        i = info_data.get(sku, {})
        target = alias[sku]
        if not s:
            alias_sw = fetch_sw_suggestions(token, [target]) or {}
            if alias_sw.get(target):
                s = dict(alias_sw[target])
                sw_map[sku] = s

    if sw_map.get(sku):
        enrich_sw_eox(token, {sku: sw_map[sku]})
        s = sw_map[sku]

    enrich_replacement_chains(token, [sku], eox_data)
    e = eox_data.get(sku, {})

    columns = {name: fn(sku, e, i, s) for name, fn in ALL_COLUMNS}

    return jsonify({
        "sku": sku,
        "columns": columns,
        "raw": {
            "eox": e,
            "product_info": i,
            "sw_suggestion": s,
        },
    })


@app.route("/api/debug/eox/<path:sku>")
def debug_eox(sku: str):
    client_id     = os.getenv("CLIENT_ID", "")
    client_secret = os.getenv("CLIENT_SECRET", "")
    try:
        token = get_access_token(client_id, client_secret)
    except Exception as e:
        return jsonify({"error": str(e)}), 401
    headers = {"Authorization": f"Bearer {token}"}
    url = f"{EOX_BASE}/EOXByProductID/1/{requests.utils.quote(sku)}"
    resp = requests.get(url, params={"responseencoding": "json"}, headers=headers, timeout=30)
    return jsonify({"status": resp.status_code, "body": resp.json() if resp.ok else resp.text})


@app.route("/api/debug/hw/<path:sku>")
def debug_hw(sku: str):
    client_id     = os.getenv("CLIENT_ID", "")
    client_secret = os.getenv("CLIENT_SECRET", "")
    try:
        token = get_access_token(client_id, client_secret)
    except Exception as e:
        return jsonify({"error": str(e)}), 401
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    resp = requests.get(f"{PRODUCT_INFO}/{requests.utils.quote(sku)}", headers=headers, timeout=30)
    return jsonify({"status": resp.status_code, "body": resp.json() if resp.ok else resp.text})


@app.route("/api/debug/sw/<path:sku>")
def debug_sw(sku: str):
    client_id     = os.getenv("CLIENT_ID", "")
    client_secret = os.getenv("CLIENT_SECRET", "")
    try:
        token = get_access_token(client_id, client_secret)
    except Exception as e:
        return jsonify({"error": str(e)}), 401
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    resp = requests.get(f"{SW_SUGGEST}/{requests.utils.quote(sku)}", headers=headers, timeout=30)
    return jsonify({"status": resp.status_code, "body": resp.json() if resp.ok else resp.text})


@app.route("/api/presets", methods=["GET"])
def get_presets():
    with _data_lock:
        d = _read_json(_PRESETS_F, {})
    return jsonify({"presets": d.get("presets", []), "active": d.get("active", "default")})


@app.route("/api/presets", methods=["POST"])
def save_presets():
    data = request.get_json(silent=True) or {}
    with _data_lock:
        _write_json(_PRESETS_F, {"presets": data.get("presets", []), "active": data.get("active", "default")})
    return jsonify({"ok": True})


@app.route("/api/overrides", methods=["GET"])
def get_overrides():
    with _data_lock:
        d = _read_json(_OVERRIDES_F, {})
    return jsonify({"overrides": d})


@app.route("/api/overrides", methods=["POST"])
def save_overrides():
    data = request.get_json(silent=True) or {}
    with _data_lock:
        _write_json(_OVERRIDES_F, data.get("overrides", {}))
    return jsonify({"ok": True})


@app.route("/api/history", methods=["GET"])
def get_history():
    with _data_lock:
        d = _read_json(_HISTORY_F, {})
    return jsonify({"history": d.get("history", [])})


@app.route("/api/history", methods=["POST"])
def save_history():
    data = request.get_json(silent=True) or {}
    new_history = data.get("history", [])
    with _data_lock:
        old = _read_json(_HISTORY_F, {}).get("history", [])
        old_ids = {e.get("file_id") for e in old if e.get("file_id")}
        new_ids = {e.get("file_id") for e in new_history if e.get("file_id")}
        for fid in old_ids - new_ids:
            try:
                os.remove(os.path.join(_DOWNLOADS_DIR, f"{fid}.xlsx"))
            except OSError:
                pass
        _write_json(_HISTORY_F, {"history": new_history})
    return jsonify({"ok": True})


@app.route("/api/history", methods=["DELETE"])
def clear_history():
    with _data_lock:
        _write_json(_HISTORY_F, {"history": []})
    for f in os.listdir(_DOWNLOADS_DIR):
        try:
            os.remove(os.path.join(_DOWNLOADS_DIR, f))
        except OSError:
            pass
    return jsonify({"ok": True})


@app.route("/api/history/<entry_id>/download")
def download_history_entry(entry_id: str):
    with _data_lock:
        history = _read_json(_HISTORY_F, {}).get("history", [])
    entry = next((e for e in history if e.get("id") == entry_id), None)
    if not entry:
        return jsonify({"error": "Entry not found"}), 404
    file_id = entry.get("file_id")
    if not file_id:
        return jsonify({"error": "No file associated with this entry"}), 404
    dl_path = os.path.join(_DOWNLOADS_DIR, f"{file_id}.xlsx")
    if not os.path.exists(dl_path):
        return jsonify({"error": "File no longer available"}), 404
    return send_file(
        dl_path,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=entry.get("label", f"{entry_id}.xlsx"),
    )


@app.route("/api/history/<entry_id>/repair", methods=["POST"])
def repair_history_entry(entry_id: str):
    """Re-compute the replacements dict for a history entry from its Excel file.

    Collects every preset label that maps to the 'Replacement PID' canonical
    column (including the canonical name itself), then searches the file headers
    for any of those labels to find the right column.
    """
    file = request.files.get("file")
    if not file:
        return jsonify({"error": "Excel file required"}), 400

    # Gather every label (across all presets) that maps to "Replacement PID"
    repl_labels: set[str] = {"replacement pid"}
    with _data_lock:
        presets_data = _read_json(_PRESETS_F, {})
    for preset in presets_data.get("presets", []):
        for col_entry in preset.get("cols", []):
            if col_entry.get("col") == "Replacement PID":
                lbl = (col_entry.get("label") or "").strip().lower()
                if lbl:
                    repl_labels.add(lbl)

    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        raw_hdrs = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
        headers = [str(c or "").strip() for c in raw_hdrs]

        repl_col = next(
            (i for i, h in enumerate(headers) if h.lower() in repl_labels),
            None,
        )
        if repl_col is None:
            found = [h for h in headers if h]
            return jsonify({
                "error": f"No 'Replacement PID' column found. "
                         f"Headers in file: {found}. "
                         f"Expected one of: {sorted(repl_labels)}"
            }), 400

        replacement_counts: dict[str, int] = {}
        for row in ws.iter_rows(min_row=2, values_only=True):
            val = str(row[repl_col] if repl_col < len(row) else "").strip()
            if val:
                key = val.upper()
                replacement_counts[key] = replacement_counts.get(key, 0) + 1
    except Exception as e:
        return jsonify({"error": f"Could not read file: {e}"}), 400

    with _data_lock:
        data = _read_json(_HISTORY_F, {})
        history = data.get("history", [])
        entry = next((e for e in history if e.get("id") == entry_id), None)
        if entry is None:
            return jsonify({"error": "History entry not found"}), 404
        entry.setdefault("stats", {})["replacements"] = replacement_counts
        _write_json(_HISTORY_F, {"history": history})

    return jsonify({"ok": True, "count": len(replacement_counts)})


@app.route("/api/compare", methods=["POST"])
def compare_route():
    file_a = request.files.get("file_a")
    file_b = request.files.get("file_b")

    if not file_a or not file_b:
        return jsonify({"error": "Two Excel files are required."}), 400

    try:
        data_a, cols_a = _parse_excel_for_compare(file_a)
        data_b, cols_b = _parse_excel_for_compare(file_b)
    except Exception as e:
        return jsonify({"error": f"Could not read file: {e}"}), 400

    # Use file A's column order as reference; append any extra columns from B
    all_cols = list(cols_a[1:]) if len(cols_a) > 1 else []
    for c in (cols_b[1:] if len(cols_b) > 1 else []):
        if c not in all_cols:
            all_cols.append(c)

    pids_a = set(data_a.keys())
    pids_b = set(data_b.keys())

    added   = sorted(pids_b - pids_a)
    removed = sorted(pids_a - pids_b)

    changed: list[dict] = []
    unchanged_count = 0

    for pid in sorted(pids_a & pids_b):
        row_a = data_a[pid]
        row_b = data_b[pid]
        diffs: list[dict] = []
        for col in all_cols:
            v_a = (row_a.get(col) or "").strip()
            v_b = (row_b.get(col) or "").strip()
            if v_a != v_b:
                diffs.append({"col": col, "old": v_a, "new": v_b})
        if diffs:
            changed.append({"pid": pid, "changes": diffs})
        else:
            unchanged_count += 1

    return jsonify({
        "added":     [{"pid": p, "values": data_b[p]} for p in added],
        "removed":   [{"pid": p, "values": data_a[p]} for p in removed],
        "changed":   changed,
        "unchanged": unchanged_count,
        "cols":      all_cols,
    })


if __name__ == "__main__":
    app.run(debug=True, port=5001, host="0.0.0.0")
