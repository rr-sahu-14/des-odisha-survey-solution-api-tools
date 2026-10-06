#!/usr/bin/env python3
"""
Approve / reject Survey Solutions interviews in bulk, using an API user's
login and password (HTTP Basic auth).

HOW IT WORKS
------------
Two separate input files, different roles and different formats:

  KEYS_EXCEL_FILE  -- an ordinary .xlsx from the District, with just an
                       interview__key column (no id).
  LOOKUP_TAB_FILE  -- a Survey Solutions .tab export (tab-delimited, not
                       comma-delimited), which carries both interview__key
                       and interview__id together on every row.

The script reads the keys you actually need from the first file, reads the
key->id lookup from the second, and matches them locally in memory -- no
network call involved for the lookup itself, and no repeated typing of keys.

No status pre-check is done here -- that's on you to confirm before running.
For each key, the script does exactly ONE network call:
  PATCH /api/v1/interviews/{id}/<action>

If an interview isn't actually in a status that allows the action, the
server itself says so (HTTP 406), exactly like the web interface would --
the script just reports that and moves on to the next key. So the work done
is always exactly proportional to how many keys you listed, however large
the whole survey is.

HOW TO USE
----------
1. Fill in the CONFIG block just below (server, workspace, the two input
   files, and which action to run).
2. KEYS_EXCEL_FILE needs an "interview__key" column (or just one column of
   keys with any header). LOOKUP_TAB_FILE needs BOTH "interview__key" and
   "interview__id" columns -- any Survey Solutions tabular data export has
   both, tab-separated.
3. Run:  python ss_interview_actions.py
   It asks for your API user id and password on screen (plain input(), not
   getpass -- some terminals don't support getpass's hidden-input prompt;
   note the password WILL be visible on screen as you type it), then starts
   in DRY_RUN mode (see CONFIG) so you can check the resolved ids look right
   before it changes anything.

Actions (Survey Solutions API v1, all PATCH):
  approve     -> .../interviews/{id}/approve     (Supervisor)   needs status Completed
  reject      -> .../interviews/{id}/reject      (Supervisor)   needs status Completed
  hq-approve  -> .../interviews/{id}/hqapprove   (Headquarters) needs Completed or ApprovedBySupervisor
  hq-reject   -> .../interviews/{id}/hqreject    (Headquarters) needs Completed or ApprovedBySupervisor
  unapprove   -> .../interviews/{id}/unapprove   (Headquarters) needs ApprovedByHeadquarters
"""
import csv
import os
import sys
import time
import uuid
from datetime import datetime
from urllib.parse import quote

import pandas as pd
import requests
from requests.auth import HTTPBasicAuth

# =====================================================================
#  CONFIG -- edit these, then just run:  python ss_interview_actions.py
# =====================================================================
SERVER = "https://your-server.example.org"   # base URL of your Survey Solutions server (real value: local_config.py)
WORKSPACE = "primary"                    # workspace name (real value: local_config.py)

# Looked up next to this script, wherever you run it from; both files are git-ignored.
_HERE = os.path.dirname(os.path.abspath(__file__))
KEYS_EXCEL_FILE = os.path.join(_HERE, "interview__keys.xlsx")   # from District -- just interview__key values
LOOKUP_TAB_FILE = os.path.join(_HERE, "out.tab")                # Survey Solutions .tab export -- has BOTH
                                                     # interview__key and interview__id

ACTION = "hq-reject"       # one of: approve | reject | hq-approve | hq-reject | unapprove
COMMENT = "Price outlier, please recheck"          # sent to the enumerator/supervisor

DRY_RUN = False              # True = only resolve keys + show what WOULD be sent, change nothing
                             # set to False once the dry run looks right
# =====================================================================

# Private values (real server, workspace) live in local_config.py in the parent
# folder, which is git-ignored.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from local_config import SERVER, WORKSPACE  # noqa: F811
except ImportError:
    pass


ACTION_ENDPOINT = {
    "approve": "approve",
    "reject": "reject",
    "hq-approve": "hqapprove",
    "hq-reject": "hqreject",
    "unapprove": "unapprove",
}

HTTP_STATUS_MEANING = {
    200: "OK",
    204: "OK",
    400: "Bad request",
    401: "Login/password rejected",
    403: "User not allowed to do this action",
    404: "Interview id not found",
    406: "Interview is in a status that does not allow this action",
}

TIMEOUT = 60      # seconds per request
DELAY = 0.2       # seconds between interviews, so the server isn't hammered


# --------------------------------------------------------------- key input --
def read_table_smart(path):
    """Read a table into a DataFrame no matter what its extension claims.
    Many 'Excel' exports are actually CSV, tab-delimited text, or an old
    .xls file wearing an .xlsx name -- pandas' pd.read_excel then fails with
    'Excel file format cannot be determined'. This looks at the file's real
    signature bytes instead of trusting the extension:
      - PK..   (zip)          -> real .xlsx/.xlsm  -> openpyxl
      - D0CF11E0 (OLE2)       -> real old .xls      -> xlrd
      - anything else         -> plain text         -> try comma, then tab,
                                                         then semicolon
    """
    with open(path, "rb") as f:
        head = f.read(8)

    if head[:2] == b"PK":
        return pd.read_excel(path, dtype=str, engine="openpyxl")
    if head[:4] == b"\xd0\xcf\x11\xe0":
        try:
            return pd.read_excel(path, dtype=str, engine="xlrd")
        except ImportError:
            sys.exit(f"'{path}' is an old-style .xls file, which needs the 'xlrd' package. "
                      f"Run:  pip install xlrd   then try again.")

    # Not a real Excel binary -- treat it as delimited text. Try each likely
    # delimiter and keep whichever one actually splits it into more than one
    # column (a wrong separator "succeeds" too, just as a single column, so
    # comma must not win by default over the real delimiter).
    candidates = []
    for sep in (",", "\t", ";"):
        try:
            df = pd.read_csv(path, dtype=str, sep=sep)
            candidates.append(df)
        except Exception:
            continue
    multi_col = [df for df in candidates if df.shape[1] > 1]
    if multi_col:
        return max(multi_col, key=lambda df: df.shape[1])
    if candidates:
        return candidates[0]
    sys.exit(f"Could not read '{path}' as Excel or as delimited text -- "
              f"open it and check what format it's really saved as.")


def read_keys_from_excel(path):
    """Read just interview__key values from the District's Excel file.
    Uses the 'interview__key' column if there's a header with that name,
    otherwise falls back to the first column. Returns a de-duplicated,
    order-preserving list."""
    if not os.path.exists(path):
        sys.exit(f"Excel file not found: {path}")

    df = read_table_smart(path)
    cols = {c.strip().lower(): c for c in df.columns}
    key_col = cols.get("interview__key", df.columns[0])

    keys, seen = [], set()
    for raw in df[key_col].dropna().astype(str):
        key = raw.strip()
        if key and key not in seen:
            seen.add(key)
            keys.append(key)
    return keys


def normalize_guid(raw):
    """Survey Solutions' .tab exports store interview__id in the compact
    32-hex-character form (no hyphens), but the REST API's URLs need the
    standard 8-4-4-4-12 hyphenated form -- otherwise every lookup silently
    fails to match. uuid.UUID() parses either form (and is case-insensitive)
    and str() on it always gives back the canonical hyphenated lowercase
    form, so this normalizes whichever format shows up."""
    raw = str(raw).strip()
    try:
        return str(uuid.UUID(raw))
    except (ValueError, AttributeError, TypeError):
        return raw   # not a recognizable GUID; leave as-is and let the API reject it


def read_keymap_from_tab(path):
    """Read interview__key -> interview__id from a Survey Solutions .tab
    export (tab-delimited, both columns required). Returns a dict, with
    every id normalized to the standard hyphenated GUID form."""
    if not os.path.exists(path):
        sys.exit(f"Lookup .tab file not found: {path}")

    df = pd.read_csv(path, sep="\t", dtype=str)
    cols = {c.strip().lower(): c for c in df.columns}
    key_col = cols.get("interview__key")
    id_col = cols.get("interview__id")
    missing = [name for name, col in (("interview__key", key_col), ("interview__id", id_col)) if not col]
    if missing:
        sys.exit(f"'{path}' must have columns {missing}. Found columns: {list(df.columns)}")

    keymap = {}
    for _, row in df.iterrows():
        key = str(row[key_col]).strip() if pd.notna(row[key_col]) else ""
        iid = str(row[id_col]).strip() if pd.notna(row[id_col]) else ""
        if key and iid:
            keymap[key] = normalize_guid(iid)
    return keymap


# ------------------------------------------------------------------ action --
def do_action(session, base_url, interview_id, action, comment):
    url = f"{base_url}/api/v1/interviews/{quote(interview_id)}/{ACTION_ENDPOINT[action]}"
    params = {"comment": comment} if comment else {}
    return session.patch(url, params=params, timeout=TIMEOUT)


def main():
    if ACTION not in ACTION_ENDPOINT:
        sys.exit(f"ACTION must be one of {list(ACTION_ENDPOINT)}, got '{ACTION}'")

    keys = read_keys_from_excel(KEYS_EXCEL_FILE)
    if not keys:
        sys.exit(f"No interview__key values found in {KEYS_EXCEL_FILE}.")

    keymap = read_keymap_from_tab(LOOKUP_TAB_FILE)

    no_id = [k for k in keys if k not in keymap]
    if no_id:
        print(f"WARNING: {len(no_id)} key(s) from {KEYS_EXCEL_FILE} were not found in "
              f"{LOOKUP_TAB_FILE} and will be skipped: {no_id[:5]}{' ...' if len(no_id) > 5 else ''}")

    # Ask for the API user id and password on screen with plain input().
    # (getpass's hidden-input prompt doesn't work in every terminal/IDE, so
    # this uses ordinary input() instead -- the password WILL be visible on
    # screen as you type it.) Env vars still override, if set.
    api_user = os.environ.get("SS_USER") or input("API user id: ").strip()
    password = os.environ.get("SS_PASSWORD") or input("Password: ").strip()

    base_url = f"{SERVER.rstrip('/')}/{quote(WORKSPACE)}"
    out = f"interview_{ACTION}_{datetime.now():%Y%m%d_%H%M%S}.csv"

    print(f"Server  : {base_url}")
    print(f"Action  : {ACTION}")
    print(f"Keys    : {len(keys)}  (from {KEYS_EXCEL_FILE}, matched against {LOOKUP_TAB_FILE})")
    if COMMENT:
        print(f"Comment : {COMMENT}")
    print(f"Mode    : {'DRY RUN -- nothing will be sent' if DRY_RUN else 'LIVE'}")
    print()

    session = requests.Session()
    session.auth = HTTPBasicAuth(api_user, password)
    session.headers.update({"Accept": "application/json"})

    results, ok, fail, skipped = [], 0, 0, 0
    for n, key in enumerate(keys, 1):
        interview_id = keymap.get(key)
        if not interview_id:
            print(f"[{n}/{len(keys)}] key {key}: not found in {LOOKUP_TAB_FILE} - skipped")
            results.append((key, "", "SKIPPED", f"not found in {LOOKUP_TAB_FILE}", ""))
            skipped += 1
            continue

        if DRY_RUN:
            print(f"[{n}/{len(keys)}] key {key} -> id {interview_id}  WOULD {ACTION.upper()}")
            results.append((key, interview_id, "DRY_RUN", "", ""))
            continue

        try:
            r = do_action(session, base_url, interview_id, ACTION, COMMENT)
            success = r.status_code in (200, 204)
            msg = HTTP_STATUS_MEANING.get(r.status_code, r.reason)
            detail = "" if success else (r.text or "").strip()[:300]
            code = r.status_code
        except requests.RequestException as e:
            success, msg, detail, code = False, "Network error", str(e)[:300], ""

        results.append((key, interview_id, "OK" if success else "FAILED", msg, detail))
        ok += success
        fail += (not success)
        print(f"[{n}/{len(keys)}] key {key} -> id {interview_id}  "
              f"HTTP {code}  {'OK' if success else 'FAILED - ' + msg}")

        if code == 401:
            print("Stopping: login/password rejected.")
            break
        time.sleep(DELAY)

    if not DRY_RUN:
        with open(out, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            w.writerow(["interview__key", "interview__id", "result", "message", "server_detail"])
            w.writerows(results)
        print(f"\nDone. Success: {ok}  Failed: {fail}  Skipped: {skipped}")
        print(f"Results saved to {out}")
    else:
        would_act = sum(1 for r in results if r[2] == "DRY_RUN")
        print(f"\nDry run complete. Would act on {would_act}, skip {skipped}.")
        print("Set DRY_RUN = False in the CONFIG block once this looks right.")


if __name__ == "__main__":
    main()