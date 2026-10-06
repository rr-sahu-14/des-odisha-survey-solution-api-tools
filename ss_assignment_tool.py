"""
ss_assignment_tool.py - Survey Solutions assignments through the HTTP API.

Pick what to run in ACTIONS (they run in the order listed):

    "export"    save every assignment to an Excel file
    "create"    create new assignments from an Excel/CSV file (one row each)
    "close"     close open assignments      (all active ones, or ids from a file)
    "archive"   archive assignments         (all active ones, or ids from a file)
    "rollover"  monthly top-up: quantity = interviews created + target per schedule

Everything is set in the CONFIG block; DRY_RUN = True (default) only prints the plan.
Requires: pip install requests pandas openpyxl
"""

import csv
import json
import os
import sys
from collections import Counter
from datetime import datetime

import pandas as pd
import requests
from requests.auth import HTTPBasicAuth

# =============================================================================
# CONFIG
# =============================================================================
ACTIONS = ["export"]       # any of: "export" "create" "close" "archive" "rollover"
DRY_RUN = True             # True = print the plan only, send nothing

SERVER = "https://your-server.example.org"      # no trailing slash (real value: local_config.py)
WORKSPACE = "primary"                           # (real value: local_config.py)
MAX_CHANGES = 5000         # a live action refuses to run on more items than this

# --- close / archive: which assignments --------------------------------------
SCOPE = "all"              # "all" = every active one | "selected" = ids in IDS_FILE
IDS_FILE = "assignment_ids.xlsx"    # xlsx / xls / csv / tab text; first column unless:
ID_COLUMN = None
FILTER_QUESTIONNAIRE = ""  # only this questionnaire GUID (applies to export/close/archive/rollover)
MIN_CREATED_UTC = ""        # e.g. "2026-08-01": ignore assignments created before

# --- create ------------------------------------------------------------------
CREATE_FILE = "new_assignments.xlsx"   # one row per assignment
CREATE_QUESTIONNAIRE = ""              # "guid$version", e.g. "00000000000000000000000000000000$2"
CREATE_RESPONSIBLE_COL = "Responsible" # column holding the login name of the enumerator
CREATE_QUANTITY_COL = "Quantity"       # optional column; quantity 1 when the column is absent
CREATE_FIRST_N = 0                     # 0 = all rows; e.g. 1 = create only the first row (test)
# every other column in CREATE_FILE is sent as identifying data (column name = variable name)

# --- rollover ----------------------------------------------------------------
CYCLE = datetime.now().strftime("%Y-%m")   # label of this month's run
SCHEDULES = {   # questionnaire GUID (32 hex, no hyphens, no "$version") : (label, interviews per cycle)
    # "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx": ("Schedule 1A", 2),
    # "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx": ("Schedule 1B", 2),
    # "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx": ("Schedule 2",  1),
}
ALLOW_DECREASE = True      # True: last month's unfinished backlog is dropped
STATE_FILE = "rollover_state.csv"

PAGE_SIZE = 100
TIMEOUT = 60
# =============================================================================

# Private values (real server, workspace, GUIDs) live in local_config.py, which is git-ignored.
try:
    from local_config import *  # noqa: F401,F403
except ImportError:
    pass

BASE = f"{SERVER.rstrip('/')}/{WORKSPACE}/api/v1"
S = None       # requests session, created in main()
LOG = []       # one row per item touched, written to a CSV at the end


# ---------------------------------------------------------------- helpers -----
def getv(d, name, default=None):
    """Case-insensitive dict lookup (API field casing varies between versions)."""
    for k, v in d.items():
        if k.lower() == name.lower():
            return v
    return default


def guid_of(qid):
    """'00000000000000000000000000000000$2' -> '00000000000000000000000000000000'"""
    return str(qid).split("$")[0].replace("-", "").lower()


def _py(o):
    """json helper: numpy numbers -> plain Python."""
    return o.item() if hasattr(o, "item") else str(o)


def call(method, path, body=None, params=None):
    kw = {"params": params, "timeout": TIMEOUT}
    if body is not None:
        kw["data"] = json.dumps(body, default=_py)
        kw["headers"] = {"Content-Type": "application/json"}
    r = S.request(method, f"{BASE}{path}", **kw)
    if r.status_code == 401:
        sys.exit("401 Unauthorized - bad API user/password. Stopping.")
    return r


def read_table(path):
    """Real xlsx / old xls / csv / tab text - whatever the file really is."""
    with open(path, "rb") as f:
        head = f.read(8)
    if head[:2] == b"PK":
        return pd.read_excel(path, engine="openpyxl")
    if head[:4] == bytes.fromhex("D0CF11E0"):
        return pd.read_excel(path, engine="xlrd")
    for sep in (",", "\t", ";"):
        df = pd.read_csv(path, sep=sep)
        if df.shape[1] > 1:
            return df
    return pd.read_csv(path)


def fetch_all():
    """Every assignment on the server (paged), narrowed by the config filters."""
    seen, rows, offset, total = set(), [], 0, None
    while True:
        r = call("GET", "/assignments", params={"Limit": PAGE_SIZE, "Offset": offset})
        if r.status_code != 200:
            sys.exit(f"Listing assignments failed: HTTP {r.status_code} {r.text[:200]}")
        data = r.json()
        page = getv(data, "Assignments", []) or []
        total = getv(data, "TotalCount", total)
        new = [a for a in page if getv(a, "Id") not in seen]
        if not new:
            break
        seen.update(getv(a, "Id") for a in new)
        rows += new
        offset += len(page)                  # the server may cap Limit below PAGE_SIZE
        if total is not None and offset >= total:
            break
    if total is not None and len(rows) != total:
        print(f"WARNING: listed {len(rows)} of {total} assignments - check paging.")
    return [a for a in rows if keep(a)]


def keep(a):
    if FILTER_QUESTIONNAIRE and FILTER_QUESTIONNAIRE.lower() not in guid_of(getv(a, "QuestionnaireId", "")):
        return False
    if MIN_CREATED_UTC and str(getv(a, "CreatedAtUtc", ""))[:10] < MIN_CREATED_UTC:
        return False
    return True


def status(a):
    if getv(a, "Archived"):
        return "Archived"
    q, done = getv(a, "Quantity"), getv(a, "InterviewsCount", 0) or 0
    if q is None or q < 0:
        return "Open (unlimited)"
    return "Open" if done < q else "Closed"


def lab(i):
    return f"#{i.get('id', i.get('row'))} {i.get('who', '')} {i.get('note', '')}".strip()


def run(action, items, send, on_ok=None):
    """Shared engine: show plan -> (dry run stops) -> confirm -> send one by one -> log."""
    print(f"\n[{action}] {len(items)} item(s)")
    for i in items[:10]:
        print("  ", lab(i))
    if len(items) > 10:
        print(f"   ... and {len(items) - 10} more (all in the log file)")
    key = lambda i: i.get("id", i.get("row"))
    if DRY_RUN:
        LOG.extend({"action": action, "key": key(i), "outcome": "WOULD_DO", "http": "",
                    "detail": lab(i)} for i in items)
        print("   DRY RUN - nothing sent.")
        return
    if not items:
        return
    if len(items) > MAX_CHANGES:
        sys.exit(f"{len(items)} items is more than MAX_CHANGES={MAX_CHANGES}.")
    if input(f"Type YES to {action} {len(items)} item(s): ").strip() != "YES":
        sys.exit("Cancelled.")
    for n, i in enumerate(items, 1):
        r = send(i)
        ok = r.status_code in (200, 204)
        LOG.append({"action": action, "key": key(i), "outcome": "OK" if ok else f"FAILED_{r.status_code}",
                    "http": r.status_code, "detail": (lab(i) + " | " + r.text[:100]).replace("\n", " ")})
        if ok and on_ok:
            on_ok(i)
        print(f"[{n}/{len(items)}] {lab(i)}: {'OK' if ok else 'FAILED ' + str(r.status_code) + ' ' + r.text[:80]}")


# ---------------------------------------------------------------- actions -----
def export():
    rows = fetch_all()
    df = pd.DataFrame(rows).drop(columns=["Password"], errors="ignore")
    df["Status"] = [status(a) for a in rows]
    df["Outstanding"] = [max((getv(a, "Quantity") or 0) - (getv(a, "InterviewsCount", 0) or 0), 0)
                         if status(a) == "Open" else 0 for a in rows]
    name = f"assignments_export_{datetime.now():%Y%m%d_%H%M}.xlsx"
    df.to_excel(name, index=False)
    print(f"\n[export] {len(df)} assignments -> {name}")
    print("  ", dict(Counter(df["Status"])))


def _targets(open_only):
    if SCOPE == "selected":                  # no pre-check: the server answers per id
        df = read_table(IDS_FILE)
        col = ID_COLUMN or df.columns[0]
        ids = pd.to_numeric(df[col], errors="coerce").dropna().astype(int)
        return [{"id": i} for i in dict.fromkeys(ids)]
    want = lambda a: status(a).startswith("Open") if open_only else status(a) != "Archived"
    return [{"id": getv(a, "Id"), "who": getv(a, "ResponsibleName", "")} for a in fetch_all() if want(a)]


def close():
    run("close", _targets(open_only=True), lambda i: call("PATCH", f"/assignments/{i['id']}/close"))


def archive():
    run("archive", _targets(open_only=False), lambda i: call("PATCH", f"/assignments/{i['id']}/archive"))


def create():
    if "$" not in CREATE_QUESTIONNAIRE:
        sys.exit('Set CREATE_QUESTIONNAIRE as "guid$version", e.g. "00000000000000000000000000000000$2".')
    df = read_table(CREATE_FILE)
    if CREATE_RESPONSIBLE_COL not in df.columns:
        sys.exit(f"Column '{CREATE_RESPONSIBLE_COL}' not found in {CREATE_FILE}: {list(df.columns)}")
    if CREATE_FIRST_N:
        df = df.head(CREATE_FIRST_N)
    id_cols = [c for c in df.columns if c not in (CREATE_RESPONSIBLE_COL, CREATE_QUANTITY_COL)]
    items = []
    for n, row in df.iterrows():
        qty = row[CREATE_QUANTITY_COL] if CREATE_QUANTITY_COL in df.columns else 1
        items.append({"row": n + 2, "who": str(row[CREATE_RESPONSIBLE_COL]).strip(), "payload": {
            "Responsible": str(row[CREATE_RESPONSIBLE_COL]).strip(),
            "QuestionnaireId": CREATE_QUESTIONNAIRE,
            "Quantity": int(qty),
            "IdentifyingData": [{"Variable": str(c), "Answer": row[c]} for c in id_cols if not pd.isna(row[c])],
        }})
    if items:
        print("\nFirst request body (compare with POST /assignments on your /apidocs page):")
        print(json.dumps(items[0]["payload"], default=_py, indent=2))
    run("create", items, lambda i: call("POST", "/assignments", body=i["payload"]))


def rollover():
    rows = fetch_all()
    by = {}
    for a in rows:
        by.setdefault(guid_of(getv(a, "QuestionnaireId", "")), []).append(a)
    print("\nQuestionnaires on the server:")
    for g, rs in sorted(by.items()):
        act = [a for a in rs if status(a) != "Archived"]
        months = Counter(str(getv(a, "CreatedAtUtc", ""))[:7] for a in act)
        print(f"  {g}  {SCHEDULES[g][0] if g in SCHEDULES else '-- not mapped':<14}"
              f"active={len(act):<5} created: {dict(sorted(months.items()))}")
    if not SCHEDULES:
        sys.exit("Fill in SCHEDULES from the GUIDs above, then run again.")

    done_keys = set()
    if os.path.exists(STATE_FILE):
        with open(STATE_FILE, newline="", encoding="utf-8") as f:
            done_keys = {(r["cycle"], int(r["id"])) for r in csv.DictReader(f)}
    items, skipped = [], Counter()
    for a in rows:
        aid, g = getv(a, "Id"), guid_of(getv(a, "QuestionnaireId", ""))
        q, done = getv(a, "Quantity"), getv(a, "InterviewsCount", 0) or 0
        if g not in SCHEDULES:
            skipped["questionnaire not in SCHEDULES"] += 1
        elif getv(a, "Archived"):
            skipped["archived"] += 1
        elif (CYCLE, aid) in done_keys:
            skipped[f"already rolled in {CYCLE}"] += 1
        elif q is None or q < 0:
            skipped["unlimited quantity"] += 1
        elif done + SCHEDULES[g][1] == q:
            skipped["already at target"] += 1
        elif done + SCHEDULES[g][1] < q and not ALLOW_DECREASE:
            skipped["would decrease (ALLOW_DECREASE=False)"] += 1
        else:
            new = done + SCHEDULES[g][1]
            items.append({"id": aid, "who": getv(a, "ResponsibleName", ""), "old": q, "new": new,
                          "note": f"{SCHEDULES[g][0]}: qty {q} -> {new} (created {done})"})
    print(f"\nCycle {CYCLE}; skipped: {dict(skipped)}")
    print(f"  to change: {len(items)}, of which backlog dropped (quantity goes down): "
          f"{sum(i['new'] < i['old'] for i in items)}")

    def remember(i):
        new_file = not os.path.exists(STATE_FILE)
        with open(STATE_FILE, "a", newline="", encoding="utf-8") as f:
            w = csv.writer(f)
            if new_file:
                w.writerow(["cycle", "id", "old_quantity", "new_quantity", "time"])
            w.writerow([CYCLE, i["id"], i["old"], i["new"], datetime.now().isoformat(timespec="seconds")])

    run("rollover", items, lambda i: call("PATCH", f"/assignments/{i['id']}/changeQuantity", body=i["new"]), remember)


FUNCTIONS = {"export": export, "create": create, "close": close, "archive": archive, "rollover": rollover}


def main():
    global S
    bad = [a for a in ACTIONS if a not in FUNCTIONS]
    if bad:
        sys.exit(f"Unknown action(s) {bad}. Choose from {list(FUNCTIONS)}.")
    if "rollover" in ACTIONS and ({"close", "archive"} & set(ACTIONS)):
        sys.exit("Do not combine 'rollover' with 'close' or 'archive' in one run.")
    S = requests.Session()
    S.auth = HTTPBasicAuth(os.environ.get("SS_USER") or input("API user id: ").strip(),
                           os.environ.get("SS_PASSWORD") or input("API password (visible): ").strip())
    print(f"ACTIONS={ACTIONS}  DRY_RUN={DRY_RUN}  workspace={WORKSPACE}")
    try:
        for a in ACTIONS:
            FUNCTIONS[a]()
    finally:
        if LOG:
            name = f"assignment_log_{datetime.now():%Y%m%d_%H%M%S}.csv"
            with open(name, "w", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=["action", "key", "outcome", "http", "detail"])
                w.writeheader()
                w.writerows(LOG)
            print(f"\n{dict(Counter((r['action'], r['outcome']) for r in LOG))}  log: {name}")


if __name__ == "__main__":
    main()
