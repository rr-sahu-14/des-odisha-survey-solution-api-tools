# ScheduleWiseReport.py - How it works

A plain-English guide to `ScheduleWiseReport.py`, the first script of the monthly progress report.
The two scripts that analyse its output are described at the end of this guide.

## What the script does

It builds a **monthly progress report** from the Survey Solutions server and saves it as an Excel file.

It only **reads** data. It never approves, rejects, closes or changes any interview or assignment.

## The big idea: filter in two steps (for speed)

Pulling every interview from the server is slow. So the script narrows things down in two steps:

1. **Update window (fast, GraphQL):** get only interviews that were *last updated* between `UPDATE_FROM` and `UPDATE_TO`.
2. **Month filter (slower, REST):** for those interviews only, read `Date_of_Data_Collection` and keep the ones that fall in the wanted agri year and month (`AGRI_FILTER`).

The server cannot filter by `Date_of_Data_Collection` directly, so step 2 needs one REST call per interview from step 1. A narrower update window means fewer calls, which is the biggest speed-up.

## Two kinds of server calls

| | GraphQL | REST |
|---|---|---|
| Used for | The list of interviews (id, status, who is responsible, last update time...) | The answers inside one interview |
| Calls | A few, 200 interviews per page | One per interview |
| Speed | Fast | Slower, so it runs 5 at a time (`MAX_WORKERS`) |

The identifying-data field from GraphQL (the old `IDF_1`, `IDF_2`... columns) came back empty in this project, so it was removed. District, village and the date come from the REST answers instead.

## Steps when you run it

1. Asks for the username and password (typed in the terminal, never saved).
2. **Step 1 - GraphQL:** pages through interviews, newest-updated first.
   - Skips interviews updated after `UPDATE_TO`.
   - Stops early once it reaches interviews older than `UPDATE_FROM`. This is why the list is sorted by update time.
3. **Step 2 - REST:** for each interview, fetches the first `NUM_FEATURED` (5) answers. These include district, village and `Date_of_Data_Collection`.
4. **Step 2b - Month filter:** keeps only interviews whose date matches `AGRI_FILTER`. If nothing is left, it stops without writing a file.
5. **Step 3 - Build the table:** one row per interview. Columns are the basic details, then the 5 answers named by their variable names.
6. **Step 4 - Write Excel.**

## Settings (top of the file)

| Setting | Meaning |
|---|---|
| `BASE_URL`, `WORKSPACE` | Which server and workspace to read. **Private, never put real values on GitHub.** |
| `TAKE` | Interviews per GraphQL page (200) |
| `OUT_XLSX` | Where the Excel file is saved |
| `KEEP_HISTORY` | `True` = also save a copy with date and time in the name each run |
| `MAX_WORKERS` | How many REST calls run at the same time (5). Higher is faster but puts more load on the live server. |
| `NUM_FEATURED` | How many of the first answers to read (5). The date question must be inside these. |
| `POWER_BI_MODE` | `True` = one flat sheet with extra Year/Month/Week columns. `False` = one sheet per schedule. |
| `TEST_LIMIT` | A number (like 10) = use only that many interviews, for a quick test. `None` = everything. |
| `UPDATE_FROM`, `UPDATE_TO` | The update window, as whole days (UTC). `None` = no limit. |
| `AGRI_FILTER` | List of `(agri_year, month)` pairs. `[]` = keep everything. |
| `DATE_VARIABLE` | Variable name of the date question (`Date_of_Data_Collection`) |
| `SCHEDULE_MAP` | Converts a questionnaire id to a schedule name (1A, 1B, 2...) |

### Agri year

The agri year runs **July to June**. Agri year 2025 means July 2025 to June 2026.

- `(2026, 10)` = October 2026
- `(2025, 3)` = March 2026 (agri year 2025, calendar month 3)

## The functions

| Function | Job |
|---|---|
| `parse_utc` | Turns a text date into a proper UTC date |
| `agri_year_month` | Turns a date into `(agri_year, month)` |
| `filter_by_agri` | Applies `AGRI_FILTER` to the interviews |
| `gql` | Sends one GraphQL request and raises an error if the server says no |
| `fetch_all_interviews_graphql` | Step 1: pages through interviews, respecting the update window |
| `get_thread_session` | Gives each worker thread its own reusable connection |
| `fetch_interview_answers_rest` | Reads the first 5 answers of one interview |
| `fetch_all_featured_questions_rest` | Runs the REST call for all interviews in parallel |
| `make_dataframes` | Builds the one-row-per-interview table |
| `sorted_schedules` | Orders schedule names, with "Unknown" last |
| `write_excel` | Writes the Excel file |

## The Excel output

The file is saved as `<name>_latest.xlsx`. It keeps the same name every run, so Power BI can always point at it.

**With `POWER_BI_MODE = False`:**
- `Schedule_1A`, `Schedule_1B`... one sheet per schedule (columns that are empty are dropped)
- `All_Interviews` all rows together
- `About` run time, counts and settings summary

**With `POWER_BI_MODE = True`:**
- `PowerBI_Data` one flat sheet, plus `Year`, `Month`, `MonthName`, `YearMonth`, `Week`, `LastUpdateDate` and `StatusClean`

## Things to remember

- **Date must be in the first 5 answers.** If a questionnaire changes and the date moves beyond the 5th answer, those interviews are silently dropped by the filter.
- **Interviews with no date are dropped** when `AGRI_FILTER` is set.
- **Check the settings before every run:** `UPDATE_FROM`, `UPDATE_TO` and `AGRI_FILTER` must match the month you want.
- **An interview is only found if its last update falls in the window.** If an interview for the month was last updated later, move `UPDATE_TO` later.
- **Different questionnaires** can have different variable names for the first answers, so the Excel columns can differ between schedule sheets.

## What to run next: the two analysis scripts

Both read the Excel file made by `ScheduleWiseReport.py`. They **never contact the server and never ask for a password**.
They find the file's folder from `REPORT_OUT_XLSX` in `local_config.py`.

| Script | Input | Output |
|---|---|---|
| `DistrictScheduleTable.py` | `interviews_latest.xlsx`, sheet `All_Interviews` (or `PowerBI_Data`) | `district_schedule_table.xlsx`: one row per district, one column per schedule, with totals |
| `datanalysis1.py` | The newest `interviews_*.xlsx` in the folder, sheet `All_Interviews` | `analysis_<date_time>.xlsx` with sheets `Summary`, `Monthly_Count`, `District_Count`, `District_x_Status`, `Sch_Dist_Status` |

Things to know:
- `DistrictScheduleTable.py` uses the column `District_Name` (setting `DISTRICT_COLUMN`). If a questionnaire names its district column differently, change that setting.
- `DistrictScheduleTable.py` can count only some statuses with `STATUS_FILTER`. `None` counts every interview.
- Run `ScheduleWiseReport.py` first. If the Excel file is missing, both scripts stop and say so.
- `datanalysis1.py` reads the sheet `All_Interviews`, so it works with `POWER_BI_MODE = False` only.

## The whole flow

```
Survey Solutions server
        |  (ScheduleWiseReport.py: reads only)
        v
interviews_latest.xlsx
        |
        +--> DistrictScheduleTable.py --> district_schedule_table.xlsx
        |
        +--> datanalysis1.py ----------> analysis_<date_time>.xlsx
```
