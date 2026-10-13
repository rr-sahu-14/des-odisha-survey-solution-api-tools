# -*- coding: utf-8 -*-
"""
District x Schedule table.

Reads the Excel file made by ScheduleWiseReport.py and writes a small table:
one row per district, one column per schedule, with totals.

This script never contacts the survey server and never asks for a password.
It only reads one Excel file and writes another.
"""

import os
import sys

import pandas as pd

# ===== CONFIG =====
INPUT_XLSX = "interviews_latest.xlsx"         # (real folder: local_config.py -> REPORT_OUT_XLSX)
OUT_XLSX = "district_schedule_table.xlsx"     # (real folder: local_config.py -> REPORT_OUT_XLSX)

# Sheet to read. None = pick "All_Interviews", or "PowerBI_Data" if that is the one that exists.
SHEET_NAME = None

DISTRICT_COLUMN = "District_Name"
SCHEDULE_ORDER = ["1A", "1B", "2", "3", "4", "5", "6"]

# Count only these statuses, e.g. ["Completed", "ApprovedBySupervisor", "ApprovedByHeadquarters"].
# None = count every interview. Upper/lower case and underscores are ignored when matching.
STATUS_FILTER = None
# ==================

# The real folder lives in local_config.py in the parent folder, which is git-ignored.
# It is the folder that ScheduleWiseReport.py saves its Excel file in.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from local_config import REPORT_OUT_XLSX
    _folder = os.path.dirname(REPORT_OUT_XLSX)
    INPUT_XLSX = os.path.join(_folder, "interviews_latest.xlsx")
    OUT_XLSX = os.path.join(_folder, "district_schedule_table.xlsx")
except ImportError:
    pass


def clean_status(series: pd.Series) -> pd.Series:
    return series.astype(str).str.upper().str.replace("_", "", regex=False).str.strip()


def load_interviews(path: str):
    if not os.path.exists(path):
        raise SystemExit(f"Input file not found: {path}\nRun ScheduleWiseReport.py first.")
    book = pd.ExcelFile(path)
    sheet = SHEET_NAME or next((s for s in ("All_Interviews", "PowerBI_Data") if s in book.sheet_names), None)
    if sheet is None:
        raise SystemExit(f"No interview sheet found. Sheets in the file: {book.sheet_names}")
    return pd.read_excel(book, sheet_name=sheet), sheet


def build_table(df: pd.DataFrame) -> pd.DataFrame:
    # Schedule names like "2" may be read as numbers from Excel, so make them all text
    schedule = df["schedule"].astype(str).str.strip()
    district = df[DISTRICT_COLUMN].fillna("").astype(str).str.strip().replace("", "(Missing)")

    table = pd.crosstab(district, schedule)

    # Show every schedule, even if it has no interviews, plus any unexpected names (e.g. "Unknown") at the end
    extra = [c for c in table.columns if c not in SCHEDULE_ORDER]
    table = table.reindex(columns=SCHEDULE_ORDER + extra, fill_value=0)

    table["Total"] = table.sum(axis=1)
    table.loc["Total"] = table.sum(axis=0)
    table.index.name = "District"
    return table


def main():
    df, sheet = load_interviews(INPUT_XLSX)
    print(f"Read {len(df)} interviews from sheet '{sheet}'")

    for col in (DISTRICT_COLUMN, "schedule"):
        if col not in df.columns:
            raise SystemExit(f"Column '{col}' not found. Columns available: {list(df.columns)}")

    if STATUS_FILTER:
        wanted = set(clean_status(pd.Series(STATUS_FILTER)))
        df = df[clean_status(df["status"]).isin(wanted)]
        print(f"Status filter {STATUS_FILTER}: {len(df)} interviews left")

    if df.empty:
        raise SystemExit("No interviews left to count. Nothing to write.")

    missing = int(df[DISTRICT_COLUMN].isna().sum())
    if missing:
        print(f"Note: {missing} interviews have no {DISTRICT_COLUMN}; shown as '(Missing)'")

    table = build_table(df)
    print("\n" + table.to_string())

    with pd.ExcelWriter(OUT_XLSX) as writer:
        table.to_excel(writer, sheet_name="District_x_Schedule")
        ws = writer.sheets["District_x_Schedule"]
        ws.freeze_panes = "B2"
        ws.column_dimensions["A"].width = max(12, int(table.index.astype(str).str.len().max()) + 2)

    print(f"\nTable written to {OUT_XLSX}")


if __name__ == "__main__":
    main()
