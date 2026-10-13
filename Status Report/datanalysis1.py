import pandas as pd
import os
import sys
from datetime import datetime
from typing import Optional
import glob

# ===== CONFIG =====
EXCEL_FILE = "interviews_*.xlsx"   # (real folder: local_config.py -> REPORT_OUT_XLSX)
OUTPUT_FOLDER = "."                # (real folder: local_config.py -> REPORT_OUT_XLSX)

# The real folder lives in local_config.py in the parent folder, which is git-ignored.
# It is the folder that ScheduleWiseReport.py saves its Excel file in.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
try:
    from local_config import REPORT_OUT_XLSX
    OUTPUT_FOLDER = os.path.dirname(REPORT_OUT_XLSX)
    EXCEL_FILE = os.path.join(OUTPUT_FOLDER, "interviews_*.xlsx")
except ImportError:
    pass

# ==================

def find_latest_export(pattern: str) -> Optional[str]:
    """Find the most recent export file matching the pattern."""
    files = glob.glob(pattern)
    if not files:
        return None
    latest = max(files, key=os.path.getmtime)
    return latest


def load_data(excel_path: str) -> pd.DataFrame:
    """Load the All_Interviews sheet from exported Excel."""
    print(f"Loading data from: {excel_path}")
    df = pd.read_excel(excel_path, sheet_name="All_Interviews")
    print(f"✓ Loaded {len(df)} interviews")
    return df


def prepare_data(df: pd.DataFrame) -> pd.DataFrame:
    """Prepare data for analysis."""
    df = df.copy()

    if "Date_of_Data_Collection" in df.columns:
        df["Date_of_Data_Collection"] = pd.to_datetime(df["Date_of_Data_Collection"], errors="coerce")
        df["Year_Month"] = df["Date_of_Data_Collection"].dt.to_period("M").astype(str)
        df["Year"] = df["Date_of_Data_Collection"].dt.year
        df["Month"] = df["Date_of_Data_Collection"].dt.month_name()

    return df


def analyze_monthly(df: pd.DataFrame) -> pd.DataFrame:
    """Monthly data collection counts."""
    if "Year_Month" not in df.columns:
        return pd.DataFrame({"Error": ["Date_of_Data_Collection column not found"]})

    monthly = df.groupby("Year_Month").size().reset_index(name="Count")
    monthly = monthly.sort_values("Year_Month")
    return monthly


def analyze_district(df: pd.DataFrame) -> pd.DataFrame:
    """District-wise data collection counts."""
    if "District_Name" not in df.columns:
        return pd.DataFrame({"Error": ["District_Name column not found"]})

    district = df.groupby("District_Name").size().reset_index(name="Count")
    district = district.sort_values("Count", ascending=False)
    return district


def analyze_district_status(df: pd.DataFrame) -> pd.DataFrame:
    """District-wise status breakdown."""
    if "District_Name" not in df.columns or "status" not in df.columns:
        return pd.DataFrame({"Error": ["Required columns not found"]})

    pivot = pd.crosstab(
        df["District_Name"],
        df["status"],
        margins=True,
        margins_name="Total"
    )
    pivot = pivot.reset_index()

    desired_order = [
        "District_Name",
        "COMPLETED",
        "INTERVIEWERASSIGNED",
        "APPROVEDBYSUPERVISOR",
        "REJECTEDBYSUPERVISOR",
        "APPROVEDBYHEADQUARTER",
        "REJECTEDBYHEADQUARTER"
    ]

    existing_cols = [col for col in desired_order if col in pivot.columns]
    other_cols = [col for col in pivot.columns if col not in existing_cols and col != "Total"]
    final_order = existing_cols + other_cols + ["Total"]

    return pivot[final_order]


def analyze_schedule_district_status(df: pd.DataFrame) -> pd.DataFrame:
    """Schedule-wise, Month-wise, and District-wise status breakdown."""
    if "schedule" not in df.columns or "District_Name" not in df.columns or "status" not in df.columns:
        return pd.DataFrame({"Error": ["Required columns not found"]})

    # Check if month data is available
    if "Year_Month" not in df.columns:
        return pd.DataFrame({"Error": ["Date_of_Data_Collection not processed - run prepare_data first"]})

    # Create multi-index crosstab with schedule, month, and district as rows, status as columns
    pivot = pd.crosstab(
        index=[df["schedule"], df["Year_Month"], df["District_Name"]],
        columns=df["status"],
        margins=True,
        margins_name="Total"
    )
    pivot = pivot.reset_index()

    desired_order = [
        "schedule",
        "Year_Month",
        "District_Name",
        "COMPLETED",
        "INTERVIEWERASSIGNED",
        "APPROVEDBYSUPERVISOR",
        "REJECTEDBYSUPERVISOR",
        "APPROVEDBYHEADQUARTER",
        "REJECTEDBYHEADQUARTER"
    ]

    existing_cols = [col for col in desired_order if col in pivot.columns]
    other_cols = [col for col in pivot.columns if col not in existing_cols and col != "Total"]
    final_order = existing_cols + other_cols + ["Total"]

    return pivot[final_order]

def create_summary(df: pd.DataFrame) -> pd.DataFrame:
    """Create overall summary statistics."""
    summary_data = [
        ["Total Interviews", len(df)],
        ["Unique Interview IDs", df["id"].nunique() if "id" in df.columns else "N/A"],
        ["", ""],
    ]

    if "schedule" in df.columns:
        summary_data.append(["Schedules", ""])
        for schedule in sorted(df["schedule"].unique()):
            count = len(df[df["schedule"] == schedule])
            summary_data.append([f"  {schedule}", count])
        summary_data.append(["", ""])

    if "status" in df.columns:
        summary_data.append(["Status", ""])
        for status in df["status"].value_counts().index:
            count = len(df[df["status"] == status])
            summary_data.append([f"  {status}", count])
        summary_data.append(["", ""])

    if "District_Name" in df.columns:
        summary_data.append(["Districts", df["District_Name"].nunique()])

    if "Date_of_Data_Collection" in df.columns:
        min_date = df["Date_of_Data_Collection"].min()
        max_date = df["Date_of_Data_Collection"].max()
        summary_data.append(["Earliest Collection", min_date.strftime("%Y-%m-%d") if pd.notna(min_date) else "N/A"])
        summary_data.append(["Latest Collection", max_date.strftime("%Y-%m-%d") if pd.notna(max_date) else "N/A"])

    return pd.DataFrame(summary_data, columns=["Metric", "Value"])


def write_analysis_excel(df: pd.DataFrame, source_file: str):
    """Write all analysis results to Excel."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(OUTPUT_FOLDER, f"analysis_{timestamp}.xlsx")

    print(f"\nGenerating analysis...")

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book

        print("  ✓ Summary statistics")
        summary = create_summary(df)
        summary.to_excel(writer, sheet_name="Summary", index=False)

        print("  ✓ Monthly analysis")
        monthly = analyze_monthly(df)
        monthly.to_excel(writer, sheet_name="Monthly_Count", index=False)

        print("  ✓ District analysis")
        district = analyze_district(df)
        district.to_excel(writer, sheet_name="District_Count", index=False)

        print("  ✓ District × Status")
        district_status = analyze_district_status(df)
        district_status.to_excel(writer, sheet_name="District_x_Status", index=False)

        print("  ✓ Schedule × District × Status")
        schedule_district_status = analyze_schedule_district_status(df)
        schedule_district_status.to_excel(writer, sheet_name="Sch_Dist_Status", index=False)

        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.set_zoom(110)

            if sheet_name in ["Summary", "Monthly_Count", "District_Count"]:
                worksheet.set_column(0, 0, 30)
                worksheet.set_column(1, 1, 15)
            else:
                worksheet.set_column(0, 0, 25)
                worksheet.set_column(1, 20, 12)

    print(f"\n✅ Analysis complete: {output_file}")
    return output_file


def main():
    print("=" * 70)
    print("Survey Solutions Data Analysis")
    print("=" * 70)

    if "*" in EXCEL_FILE:
        excel_path = find_latest_export(EXCEL_FILE)
        if not excel_path:
            print(f"❌ No files found matching pattern: {EXCEL_FILE}")
            return
        print(f"📁 Auto-detected latest export")
    else:
        excel_path = EXCEL_FILE
        if not os.path.exists(excel_path):
            print(f"❌ File not found: {excel_path}")
            return

    df = load_data(excel_path)
    df = prepare_data(df)
    output_file = write_analysis_excel(df, excel_path)

    print("\n" + "=" * 70)
    print("✅ ANALYSIS COMPLETE")
    print("=" * 70)
    print(f"Output: {os.path.basename(output_file)}")
    print("\nGenerated sheets:")
    print("  • Summary - Overall statistics")
    print("  • Monthly_Count - Month-wise data collection")
    print("  • District_Count - District-wise breakdown")
    print("  • District_x_Status - District × Status cross-tab")
    print("  • Sch_Dist_Status - Schedule × District × Status")
    print("=" * 70)


if __name__ == "__main__":
    main()