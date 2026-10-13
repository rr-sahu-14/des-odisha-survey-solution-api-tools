# Survey Solutions API Tools

Python scripts that automate bulk work on a [Survey Solutions](https://docs.mysurvey.solutions/)
server through its HTTP API, replacing repetitive clicking in the Headquarters web interface.

| Tool | What it does |
|---|---|
| [`ss_assignment_tool.py`](ss_assignment_tool.py) | Manage **assignments**: export, create, close, archive, and a monthly quantity top-up (rollover) |
| [`Schedule Rejection/schedule_action.py`](Schedule%20Rejection/schedule_action.py) | Bulk **approve / reject** interviews from a list of `interview__key` values |
| [`Status Report/ScheduleWiseReport.py`](Status%20Report/ScheduleWiseReport.py) | **Monthly progress report, step 1.** Downloads interviews updated in a date window, keeps one agri month, writes an Excel file. Read-only. |
| [`Status Report/DistrictScheduleTable.py`](Status%20Report/DistrictScheduleTable.py) | **Step 2.** Reads that Excel file and makes a district x schedule count table. Never contacts the server. |
| [`Status Report/datanalysis1.py`](Status%20Report/datanalysis1.py) | **Step 2 (alternative).** Reads the same Excel file and writes a fuller analysis (month, district and status counts). Never contacts the server. |

Full guides: [assignment tool](README_ss_assignment_tool.md) · [interview actions](Schedule%20Rejection/interviewAPI.md) · [monthly status report](Status%20Report/ScheduleWiseReport_Guide.md)

## Design principles

- **Dry run by default.** Every script prints its plan and sends nothing until you switch it off.
- **Safety brakes.** A live run asks for confirmation, caps the number of changes, and stops at once on a bad login.
- **Logs.** Every item touched is written to a CSV with its outcome.
- **No secrets in code.** Credentials are asked for at run time (or read from the `SS_USER` / `SS_PASSWORD`
  environment variables). Server name and questionnaire ids live in a git-ignored `local_config.py`.

## Quick start

```
pip install -r requirements.txt
copy local_config.example.py local_config.py     # then edit it with your server details
python ss_assignment_tool.py                    # runs in dry-run mode first
```

Requires Python 3.8+ and an API user in your Survey Solutions workspace.

## Monthly status report (`Status Report/`)

These scripts only **read** data. They never approve, reject or change any interview or assignment.

1. In `local_config.py`, fill in `SERVER`, `WORKSPACE`, `SCHEDULES` (your questionnaire ids and schedule names)
   and `REPORT_OUT_XLSX` (where the Excel file should be saved).
2. Open `Status Report/ScheduleWiseReport.py` and set the month you want in the config block at the top:
   `UPDATE_FROM` and `UPDATE_TO` (interviews last updated in this window) and `AGRI_FILTER`
   (the agri year and month, e.g. `[(2026, 10)]`; the agri year runs July to June).
3. Run it: `python "Status Report/ScheduleWiseReport.py"`. It asks for your login and saves the Excel file
   (`interviews_latest.xlsx`, one sheet per schedule plus `All_Interviews`).
4. Run one of the two analysis scripts on that file:
   - `DistrictScheduleTable.py` makes a district x schedule count table.
   - `datanalysis1.py` makes month-wise, district-wise and status tables.

A narrower update window means fewer server calls, so the report runs faster.
See the [detailed guide](Status%20Report/ScheduleWiseReport_Guide.md) for every setting.

## Repository layout

```
ss_assignment_tool.py          assignment manager
Schedule Rejection/            bulk interview approve / reject
Status Report/                 monthly progress report: download, then analyse
local_config.example.py        template for your private settings
requirements.txt
```

## Data and privacy

Data files (`.csv`, `.xlsx`, `.tab`), exports, logs and `local_config.py` are listed in `.gitignore`
and are never committed.
