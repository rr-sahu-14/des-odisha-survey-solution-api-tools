# -*- coding: utf-8 -*-
"""
Survey Solutions Hybrid Exporter: GraphQL + REST API
- Uses GraphQL for interview list (fast pagination)
- Uses REST API for featured questions (first N answers from each interview)
"""

import json
import os
import sys
import threading
import time
import uuid
from datetime import datetime, timezone
from typing import Dict, Any, List, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import pandas as pd

# ===== CONFIG =====
BASE_URL = "https://your-server.example.org"   # no trailing slash (real value: local_config.py)
WORKSPACE = "primary"                           # (real value: local_config.py)
TAKE = 200
VERIFY_SSL = True
OUT_XLSX = "interviews.xlsx"                    # where the Excel file is saved (real value: local_config.py)
KEEP_HISTORY = False  # True = also save a copy with date and time in the name on every run
MAX_WORKERS = 5  # Parallel threads for REST API calls
NUM_FEATURED = 5  # Number of featured questions (first N answers)

# POWER BI MODE: Set to True to create a single flat file optimized for Power BI
POWER_BI_MODE = False  # 🔧 Set to False if you want schedule-wise sheets

# TEST MODE: Set to a number to limit interviews processed (e.g., 10 for testing)
# Set to None to process all interviews
TEST_LIMIT = None  # 🔧 Change to 10 for testing, None for full export

# UPDATE DATE WINDOW: only pull interviews last updated inside this window (dates are read as UTC).
# Use None for "no limit". Example: UPDATE_FROM = "2025-07-01", UPDATE_TO = "2025-09-30"
# Setting UPDATE_FROM saves the most time, because paging stops once it reaches older interviews.
UPDATE_FROM = "2026-10-01" 
UPDATE_TO = "2026-10-31"

# AGRI FILTER: keep only interviews whose Date_of_Data_Collection falls in these (agri_year, month) pairs.
# Agri year runs July to June, so agri_year 2025 = July 2025 to June 2026. Month is the calendar month (1-12).
# Example: [(2025, 7), (2025, 8), (2025, 9)]. Use an empty list [] to keep everything.
AGRI_FILTER = [(2026, 10)]  # 🔧 Change to desired agri year/month pairs, or [] for no filter
DATE_VARIABLE = "Date_of_Data_Collection"

# Schedule mapping: Questionnaire ID -> Schedule Name (real ids come from SCHEDULES in local_config.py)
SCHEDULE_MAP = {
    # "00000000-0000-0000-0000-000000000000": "1A",
}
# ==================

# Private values (real server, workspace, questionnaire ids, output path) live in local_config.py in
# the parent folder, which is git-ignored. If it is missing, the placeholders above stay in place.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    import local_config as _lc

    BASE_URL = _lc.SERVER
    WORKSPACE = _lc.WORKSPACE
    OUT_XLSX = getattr(_lc, "REPORT_OUT_XLSX", OUT_XLSX)
    # local_config stores ids as 32 hex characters ("Schedule 1A"); the server returns them with hyphens ("1A")
    SCHEDULE_MAP = {str(uuid.UUID(guid)): label.replace("Schedule", "", 1).strip()
                    for guid, (label, _cycles) in _lc.SCHEDULES.items()}
except ImportError:
    pass

# GraphQL Query
QUERY_INTERVIEWS = """
query GetInterviews($workspace: String!, $take: Int!, $skip: Int!) {
  interviews(workspace: $workspace, take: $take, skip: $skip, order: { updateDateUtc: DESC }) {
    totalCount
    nodes {
      id assignmentId questionnaireId questionnaireVersion status
      responsibleName responsibleId errorsCount updateDateUtc
    }
  }
}
"""


def parse_utc(value) -> Optional[pd.Timestamp]:
    ts = pd.to_datetime(value, errors="coerce", utc=True)
    return None if pd.isna(ts) else ts


def agri_year_month(date_value) -> Optional[Tuple[int, int]]:
    """Return (agri_year, calendar month). Agri year runs July to June."""
    ts = pd.to_datetime(date_value, errors="coerce")
    if pd.isna(ts):
        return None
    return (ts.year if ts.month >= 7 else ts.year - 1, ts.month)


def filter_by_agri(nodes: List[Dict[str, Any]], featured_data: Dict[str, List[Dict[str, Any]]]) -> List[Dict[str, Any]]:
    """Keep only interviews whose Date_of_Data_Collection falls in AGRI_FILTER."""
    if not AGRI_FILTER:
        return nodes
    wanted = set(AGRI_FILTER)
    kept = []
    for n in nodes:
        answers = featured_data.get(n.get("id"), [])
        date_val = next((a.get("Answer") for a in answers if a.get("Question") == DATE_VARIABLE), None)
        if agri_year_month(date_val) in wanted:
            kept.append(n)
    print(f"\n✓ Agri filter {sorted(wanted)}: kept {len(kept)} of {len(nodes)} interviews")
    return kept


def gql(session: requests.Session, query: str, variables: Dict[str, Any], user: str, pwd: str) -> Dict[str, Any]:
    r = session.post(
        f"{BASE_URL}/graphql",
        json={"query": query, "variables": variables},
        headers={"Content-Type": "application/json"},
        timeout=90,
        verify=VERIFY_SSL,
        auth=(user, pwd),
    )
    r.raise_for_status()
    data = r.json()
    if "errors" in data:
        raise RuntimeError(json.dumps(data["errors"]))
    return data


def fetch_all_interviews_graphql(user: str, pwd: str) -> Tuple[List[Dict[str, Any]], int]:
    """Fetch interviews inside the update window via GraphQL; returns (nodes, total)."""
    all_nodes: List[Dict[str, Any]] = []
    skip = 0
    total = None

    from_ts = parse_utc(UPDATE_FROM) if UPDATE_FROM else None
    # UPDATE_TO is a whole day, so the cut-off is the start of the next day
    to_ts = parse_utc(UPDATE_TO) + pd.Timedelta(days=1) if UPDATE_TO else None

    with requests.Session() as s:
        print(f"\nFetching interviews via GraphQL...")
        while True:
            vars_ = {"workspace": WORKSPACE, "take": TAKE, "skip": skip}
            resp = gql(s, QUERY_INTERVIEWS, vars_, user, pwd)
            block = resp["data"]["interviews"]
            if total is None:
                total = int(block["totalCount"])
                print(f"Total interviews: {total}")

            page = block["nodes"] or []
            n = len(page)
            print(f"  Page: skip={skip}, returned={n}")
            if n == 0:
                break

            # The list is sorted newest-updated first, so once we pass UPDATE_FROM every later row is older.
            reached_start = False
            for node in page:
                upd = parse_utc(node.get("updateDateUtc"))
                if upd is not None and from_ts is not None and upd < from_ts:
                    reached_start = True
                    break
                if upd is not None and to_ts is not None and upd >= to_ts:
                    continue
                all_nodes.append(node)

            skip += n
            if reached_start:
                print("  Reached UPDATE_FROM date - stopping early")
                break
            if skip >= total:
                break
            time.sleep(0.1)

    return all_nodes, total or 0


# Each worker thread keeps one session, so connections are reused instead of reopened per interview
_thread_local = threading.local()


def get_thread_session() -> requests.Session:
    if not hasattr(_thread_local, "session"):
        _thread_local.session = requests.Session()
    return _thread_local.session


def fetch_interview_answers_rest(interview_id: str, user: str, pwd: str) -> Optional[List[Dict[str, Any]]]:
    """Fetch first N answers (featured questions) from REST API."""
    try:
        url = f"{BASE_URL}/{WORKSPACE}/api/v1/interviews/{interview_id}"
        r = get_thread_session().get(
            url,
            auth=(user, pwd),
            headers={"Accept": "application/json"},
            timeout=30,
            verify=VERIFY_SSL
        )
        r.raise_for_status()

        # First N answers are the featured questions
        answers = r.json().get("Answers") or []
        return answers[:NUM_FEATURED]

    except Exception as e:
        print(f"    ⚠ Failed to fetch {interview_id}: {str(e)[:80]}")
        return None


def fetch_all_featured_questions_rest(interview_ids: List[str], user: str, pwd: str) -> Dict[
    str, List[Dict[str, Any]]]:
    """Fetch featured questions (first N answers) for all interviews via REST API in parallel."""
    print(f"\nFetching first {NUM_FEATURED} answers (featured questions) via REST API...")
    print(f"Processing {len(interview_ids)} interviews using {MAX_WORKERS} parallel threads...")

    featured_data: Dict[str, List[Dict[str, Any]]] = {}
    completed = 0

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        future_to_id = {
            executor.submit(fetch_interview_answers_rest, iid, user, pwd): iid
            for iid in interview_ids
        }

        for future in as_completed(future_to_id):
            iid = future_to_id[future]
            completed += 1
            if completed % 50 == 0 or completed == len(interview_ids):
                print(f"  Progress: {completed}/{len(interview_ids)}")

            try:
                answers = future.result()
                if answers:
                    featured_data[iid] = [
                        {
                            "Question": ans.get("VariableName"),
                            "Answer": ans.get("Answer")
                        }
                        for ans in answers
                    ]
            except Exception as e:
                print(f"    Error processing {iid}: {e}")

    print(f"✓ Fetched featured questions for {len(featured_data)} interviews")
    return featured_data


def make_dataframes(nodes: List[Dict[str, Any]],
                    featured_data: Dict[str, List[Dict[str, Any]]]) -> pd.DataFrame:
    """Return the wide interviews table (one row per interview)."""
    rows_wide = []
    featured_vars: List[str] = []  # variable names in the order first seen

    for n in nodes:
        base = {
            "id": n.get("id"),
            "assignmentId": n.get("assignmentId"),
            "questionnaireId": n.get("questionnaireId"),
            "questionnaireVersion": n.get("questionnaireVersion"),
            "schedule": SCHEDULE_MAP.get(n.get("questionnaireId"), "Unknown"),
            "status": n.get("status"),
            "responsibleName": n.get("responsibleName"),
            "responsibleId": n.get("responsibleId"),
            "errorsCount": n.get("errorsCount"),
            "updateDateUtc": n.get("updateDateUtc"),
        }
        # Featured questions - the variable names become the column names
        for fq in featured_data.get(base["id"], []):
            var_name = fq.get("Question")
            if var_name:
                base[var_name] = fq.get("Answer")
                if var_name not in featured_vars:
                    featured_vars.append(var_name)

        rows_wide.append(base)

    df_wide = pd.DataFrame(rows_wide)

    # Order columns: meta → featured variables (by actual names)
    meta = ["id", "assignmentId", "questionnaireId", "questionnaireVersion", "schedule", "status",
            "responsibleName", "responsibleId", "errorsCount", "updateDateUtc"]
    rest = [c for c in df_wide.columns if c not in meta + featured_vars]
    return df_wide[meta + featured_vars + rest]


def sorted_schedules(df_wide: pd.DataFrame) -> List[str]:
    """Schedule names in order, with 'Unknown' last."""
    schedules = df_wide["schedule"].unique()
    return sorted(s for s in schedules if s != "Unknown") + (["Unknown"] if "Unknown" in schedules else [])


def write_excel(df_wide: pd.DataFrame, total: int):
    # The "latest" file keeps one stable name, so Power BI can always point at it
    base_path = OUT_XLSX if OUT_XLSX.lower().endswith(".xlsx") else OUT_XLSX + ".xlsx"
    name_parts = os.path.splitext(base_path)
    latest_file = f"{name_parts[0]}_latest{name_parts[1]}"

    writer = pd.ExcelWriter(latest_file, engine="xlsxwriter")

    if POWER_BI_MODE:
        # Power BI Mode: Single flat table optimized for dashboards
        print("\n📊 Power BI Mode: Creating flat table...")

        # Add calculated columns for Power BI
        df_pbi = df_wide.copy()

        # Parse dates properly
        if "Date_of_Data_Collection" in df_pbi.columns:
            df_pbi["Date_of_Data_Collection"] = pd.to_datetime(df_pbi["Date_of_Data_Collection"], errors="coerce")
            df_pbi["Year"] = df_pbi["Date_of_Data_Collection"].dt.year
            df_pbi["Month"] = df_pbi["Date_of_Data_Collection"].dt.month
            df_pbi["MonthName"] = df_pbi["Date_of_Data_Collection"].dt.month_name()
            df_pbi["YearMonth"] = df_pbi["Date_of_Data_Collection"].dt.to_period("M").astype(str)
            df_pbi["Week"] = df_pbi["Date_of_Data_Collection"].dt.isocalendar().week

        # Add updateDate parsing - REMOVE TIMEZONE for Excel
        if "updateDateUtc" in df_pbi.columns:
            df_pbi["updateDateUtc"] = pd.to_datetime(df_pbi["updateDateUtc"], errors="coerce", utc=True)
            # Convert to timezone-naive datetime
            df_pbi["updateDateUtc"] = df_pbi["updateDateUtc"].dt.tz_localize(None)
            df_pbi["LastUpdateDate"] = df_pbi["updateDateUtc"].dt.date

        # Status cleanup
        if "status" in df_pbi.columns:
            df_pbi["StatusClean"] = df_pbi["status"].str.upper()

        # Write single sheet
        df_pbi.to_excel(writer, sheet_name="PowerBI_Data", index=False)
        print(f"  ✓ PowerBI_Data: {len(df_pbi)} rows with {len(df_pbi.columns)} columns")

        # Format the Power BI sheet
        worksheet = writer.sheets["PowerBI_Data"]
        worksheet.freeze_panes(1, 0)

        # Auto-fit columns
        for i, col in enumerate(df_pbi.columns):
            max_len = max([len(str(col))] + [len(str(v)) for v in df_pbi[col].astype(str).head(100)])
            worksheet.set_column(i, i, min(max_len + 2, 50))

    else:
        # Original Mode: Multiple sheets by schedule
        print("\nCreating Excel sheets...")
        schedules_sorted = sorted_schedules(df_wide)

        sheets_created = []

        for schedule in schedules_sorted:
            df_schedule = df_wide[df_wide["schedule"] == schedule].copy()
            df_schedule = df_schedule.dropna(axis=1, how='all')

            if len(df_schedule) > 0:
                sheet_name = f"Schedule_{schedule}"
                df_schedule.to_excel(writer, sheet_name=sheet_name, index=False)
                sheets_created.append((sheet_name, df_schedule))
                print(f"  ✓ {sheet_name}: {len(df_schedule)} interviews")

        df_wide.to_excel(writer, sheet_name="All_Interviews", index=False)
        sheets_created.append(("All_Interviews", df_wide))
        print(f"  ✓ All_Interviews: {len(df_wide)} interviews (combined)")

        about_rows = [
            ["RunAtUTC", datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S")],
            ["TotalCountReport", str(total)],
            ["RowsFetched", str(len(df_wide))],
            ["UniqueIDs", str(df_wide["id"].nunique(dropna=True))],
            ["NumFeaturedExtracted", str(NUM_FEATURED)],
            ["Workspace", WORKSPACE],
            ["BaseURL", BASE_URL],
            ["", ""],
            ["Schedule", "Count"],
        ]

        for schedule in schedules_sorted:
            count = len(df_wide[df_wide["schedule"] == schedule])
            about_rows.append([schedule, str(count)])

        about = pd.DataFrame(about_rows, columns=["Key", "Value"])
        about.to_excel(writer, sheet_name="About", index=False)
        sheets_created.append(("About", about))
        print(f"  ✓ About: Metadata and statistics")

        for name, src in sheets_created:
            ws = writer.sheets[name]
            ws.freeze_panes(1, 0)
            ws.set_zoom(110)
            for i, col in enumerate(src.columns):
                max_len = max([len(str(col))] + [len(str(v)) for v in src[col].astype(str).head(200)])
                ws.set_column(i, i, min(max_len + 2, 60))

    writer.close()
    print(f"\n✅ Excel written successfully → {latest_file}")

    # Optional history copy with date and time in the name
    if KEEP_HISTORY:
        import shutil
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        history_file = f"{name_parts[0]}_{timestamp}{name_parts[1]}"
        shutil.copy2(latest_file, history_file)
        print(f"🕘 History copy → {history_file}")

    return latest_file


if __name__ == "__main__":
    print("=" * 70)
    print("Survey Solutions Hybrid Exporter (GraphQL + REST API)")
    print(f"Extracting first {NUM_FEATURED} answers as featured questions")
    print("=" * 70)
    print("\nEnter Survey Solutions HQ credentials:")
    USERNAME = input("Username: ").strip()
    PASSWORD = input("Password: ").strip()

    # Step 1: Fetch all interviews via GraphQL (fast)
    nodes, total = fetch_all_interviews_graphql(USERNAME, PASSWORD)

    if TEST_LIMIT:
        nodes = nodes[:TEST_LIMIT]
        print(f"\n⚠️  TEST MODE: using only the first {TEST_LIMIT} interviews")

    interview_ids = [n["id"] for n in nodes if n.get("id")]
    print(f"\n✓ Fetched {len(interview_ids)} interview IDs via GraphQL")

    # Step 2: Fetch first N answers (featured questions) via REST API
    featured_data = fetch_all_featured_questions_rest(interview_ids, USERNAME, PASSWORD)

    # Step 2b: Keep only the chosen agri year / month combinations
    nodes = filter_by_agri(nodes, featured_data)
    if not nodes:
        print("No interviews left after filtering. Nothing to write.")
        raise SystemExit(0)

    # Step 3: Build dataframes
    df_wide = make_dataframes(nodes, featured_data)

    # Step 4: Write Excel
    out_file = write_excel(df_wide, total)

    print("\n" + "=" * 70)
    print("✅ EXPORT COMPLETE")
    print("=" * 70)
    print(f"File: {out_file}")
    print(f"Interviews on server: {total}")
    print(f"Interviews in report: {len(df_wide)}")
    print(f"\nSheets in Excel:")
    for schedule in sorted_schedules(df_wide):
        count = len(df_wide[df_wide["schedule"] == schedule])
        print(f"  • Schedule_{schedule}: {count} interviews")
    print(f"  • All_Interviews: {len(df_wide)} interviews (combined)")
    print(f"  • About: Metadata and statistics")
    print("\n💡 TIP: For analysis (month-wise, district-wise, status counts),")
    print("   you can now use Excel pivot tables or create a separate analysis script!")
    print("=" * 70)