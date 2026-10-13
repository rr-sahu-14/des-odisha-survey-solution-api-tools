"""Template for your PRIVATE settings.

Copy this file to local_config.py and fill in your own values.
local_config.py is git-ignored, so it never gets uploaded.
Any name defined there overrides the placeholder of the same name in the scripts.
"""

SERVER = "https://your-server.example.org"   # no trailing slash
WORKSPACE = "primary"

# questionnaire GUID (32 hex characters, no hyphens, no "$version") : (label, interviews per cycle)
SCHEDULES = {
    # "00000000000000000000000000000000": ("Schedule A", 2),
    # "00000000000000000000000000000000": ("Schedule B", 1),
}

# Where "Status Report/ScheduleWiseReport.py" saves its Excel file
# (the report also reads SERVER, WORKSPACE and SCHEDULES above; labels like "Schedule 1A" become sheets "Schedule_1A")
REPORT_OUT_XLSX = r"C:\path\to\interviews.xlsx"

# Inputs for "Schedule Rejection/tabfiles append.py": one .tab export per questionnaire
TAB_FILES = [
    # r"C:\path\to\export_1.tab",
    # r"C:\path\to\export_2.tab",
]
