# ss_assignment_tool.py — Survey Solutions assignment manager

One Python script that manages **assignments** on a Survey Solutions server through its HTTP API.
You choose which function(s) to run in one line (`ACTIONS`); everything else is set in a CONFIG
block at the top of the script. Nothing is sent to the server unless `DRY_RUN = False`.

| Function | What it does |
|---|---|
| `"export"` | Saves every assignment to an Excel file (with Status and Outstanding columns) |
| `"create"` | Creates new assignments from an Excel / CSV file, one row per assignment |
| `"close"` | Closes all open assignments, or only the ids listed in a file |
| `"archive"` | Archives all active assignments, or only the ids listed in a file |
| `"rollover"` | Monthly top-up: sets quantity = interviews already created + target for that schedule |

---

## 1. Requirements

**Software**
- Python 3.8 or newer
- Packages: `pip install requests pandas openpyxl`
- Only if one of your input files is a genuine old `.xls` file: `pip install xlrd`

**Access**
- Network access to your server (for example: `https://your-server.example.org`)
- An **API user** (user id + password) in the Survey Solutions workspace, allowed to read and change
  assignments. The script uses HTTP Basic authentication.

**Credentials**
- Either set environment variables `SS_USER` and `SS_PASSWORD` before running, or the script asks for them
  when it starts. The password is visible while you type it.

**Files**
- Put the script and any input files (ids list, new-assignment file) in the same folder, and run the
  script from that folder. File names in CONFIG are relative to where you run it.

---

## 2. Quick start

1. Open `ss_assignment_tool.py` and edit the CONFIG block:
   - `SERVER` and `WORKSPACE` (for example `"primary"`)
   - `ACTIONS` — the function(s) you want, e.g. `["export"]`
2. Run it with `DRY_RUN = True`:
   ```
   python ss_assignment_tool.py
   ```
3. Read the plan it prints. If it is what you expect, set `DRY_RUN = False`, run again, and type `YES`
   when asked.

Several functions can be listed, e.g. `ACTIONS = ["export", "archive"]`; they run in that order.
`"rollover"` cannot be combined with `"close"` or `"archive"` in the same run (the script refuses).

---

## 3. What the script writes

| File | When | Content |
|---|---|---|
| `assignments_export_<date>_<time>.xlsx` | `export` | All assignments with Status and Outstanding columns (the Password column is left out) |
| `assignment_log_<date>_<time>.csv` | any action that touches assignments | One row per assignment: action, id, outcome (`WOULD_DO`, `OK`, `FAILED_<code>`), HTTP code, detail |
| `rollover_state.csv` | `rollover` (live runs) | Which assignment was topped up in which cycle, so a second run in the same month does nothing |

Keep `rollover_state.csv` between months. Do not delete it during a month.

---

## 4. CONFIG reference

```python
ACTIONS = ["export"]        # "export" "create" "close" "archive" "rollover"
DRY_RUN = True              # True = print the plan only, send nothing
SERVER = "https://your-server.example.org"
WORKSPACE = "primary"
MAX_CHANGES = 5000          # a live action refuses to run on more items than this
```

**Used by close / archive**
```python
SCOPE = "all"               # "all" = every active assignment | "selected" = ids from IDS_FILE
IDS_FILE = "assignment_ids.xlsx"
ID_COLUMN = None            # column name; None = first column
```

**Narrowing filters** (apply to export, close, archive and rollover when `SCOPE = "all"`)
```python
FILTER_QUESTIONNAIRE = ""   # a questionnaire GUID; empty = all questionnaires
MIN_CREATED_UTC = ""        # e.g. "2026-08-01": ignore assignments created before this date
```

**Used by create** (see section 5.4)
```python
CREATE_FILE = "new_assignments.xlsx"
CREATE_QUESTIONNAIRE = ""              # "guid$version"
CREATE_RESPONSIBLE_COL = "Responsible"
CREATE_QUANTITY_COL = "Quantity"
CREATE_FIRST_N = 0                     # 0 = all rows; 1 = first row only (test)
```

**Used by rollover** (see section 5.5)
```python
CYCLE = datetime.now().strftime("%Y-%m")   # label of the month, e.g. "2026-10"
SCHEDULES = { "<guid>": ("Schedule 1A", 2), ... }
ALLOW_DECREASE = True
STATE_FILE = "rollover_state.csv"
```

---

## 5. How to run each function

### 5.1 export — get all assignments

```python
ACTIONS = ["export"]
```
Writes `assignments_export_<date>.xlsx` and prints a count by status.

- **Status** is `Open`, `Open (unlimited)`, `Closed` (quantity reached) or `Archived`.
- **Outstanding** is quantity minus interviews created, for open assignments.
- Use it before any bulk action, and to find questionnaire GUIDs (the part of `QuestionnaireId` before the `$`).

### 5.2 close — close open assignments

```python
ACTIONS = ["close"]
SCOPE = "all"          # every open assignment, or "selected"
```
- `SCOPE = "all"`: closes every assignment that is not archived and still open.
- `SCOPE = "selected"`: closes the ids in `IDS_FILE`. The script does not pre-check them; the server
  answers for each id (an assignment that cannot be closed shows up as `FAILED_<code>` in the log).
- Closing sets the quantity equal to the interviews already created, so no new interviews can be started.

### 5.3 archive — archive assignments

```python
ACTIONS = ["archive"]
SCOPE = "all"          # every assignment that is not archived, or "selected"
```
- `SCOPE = "all"` archives every assignment that is not already archived (open or closed).
- Narrow it first with `FILTER_QUESTIONNAIRE` or `MIN_CREATED_UTC` if you do not want everything.

**Ids file format** (close and archive, `SCOPE = "selected"`): an Excel or CSV file whose first column
holds assignment ids (the `Id` column in HQ or in the export file). The file may be a real `.xlsx`, an
old `.xls`, or CSV/tab text with an `.xlsx` name; the script detects the real format.

### 5.4 create — create new assignments

```python
ACTIONS = ["create"]
CREATE_FILE = "new_assignments.xlsx"
CREATE_QUESTIONNAIRE = "00000000000000000000000000000000$2"   # guid$version
CREATE_FIRST_N = 1          # test with one row first
```

**Input file** — one row per assignment:

| Responsible | Quantity | village | code |
|---|---|---|---|
| Ro205 | 2 | V0012 | 11 |

- `Responsible`: the login name of the enumerator / supervisor who receives the assignment.
- `Quantity`: optional column; if it is absent every assignment gets quantity 1.
- **Every other column is sent as identifying data.** Each column name must be the variable name
  in the questionnaire. Empty cells are not sent.
- For categorical questions use the numeric answer code, not the label.

**Every dry run prints the first request body.** Compare it with the `POST /assignments` entry on your
server's docs page (`/apidocs/index.html`) before going live.

**Always test first:** set `CREATE_FIRST_N = 1`, run live, and check that one assignment in HQ.
Then set `CREATE_FIRST_N = 0` for the rest.

> **Warning:** `create` does not remember what it created. Running it twice on the same file creates
> duplicates. Remove or move rows that are already created before running again.

### 5.5 rollover — monthly top-up

```python
ACTIONS = ["rollover"]
SCHEDULES = {
    "<1A guid>": ("Schedule 1A", 2),
    "<1B guid>": ("Schedule 1B", 2),
    "<2 guid>":  ("Schedule 2",  1),
    # ... Schedules 3, 4, 5, 6 with target 1
}
```

**Rule:** for each active assignment of a mapped schedule

> new quantity = interviews already created + target for the cycle

Schedules 1A and 1B (two rounds a month) use target 2; the other schedules (once a month) use target 1.
Interviews already created is a running total over the assignment's whole life, so earlier months
do not matter.

| Situation (target 2) | Quantity | Created | New quantity | Open for the month |
|---|---|---|---|---|
| Collected in both previous months | 4 | 4 | 6 | 2 |
| Missed last month entirely | 4 | 2 | 4 (skipped: already at target) | 2 |
| Nothing ever created | 4 | 0 | 2 (backlog dropped) | 2 |

Every village starts the month with exactly the target number of open interviews. A missed month's
backlog is dropped, so old months' prices are not entered late. Set `ALLOW_DECREASE = False` to skip
those assignments instead and handle them yourself.

**Skipped** (counted in the summary): questionnaire not in `SCHEDULES`, archived, already rolled in this
cycle, unlimited quantity, already at target.

**First run**
1. Leave `SCHEDULES` empty and run with `DRY_RUN = True`. The script lists every questionnaire GUID,
   its active-assignment count and the months in which the assignments were created.
2. Fill in `SCHEDULES` with the GUIDs of the seven schedules and their targets.
3. If one schedule shows several creation months, old batches are still active. Archive them first,
   or set `MIN_CREATED_UTC` to the date of the current batch. Otherwise all of them get topped up.
4. Test on a few assignments by temporarily setting `FILTER_QUESTIONNAIRE` to one schedule, and check them in HQ.

**Monthly routine**
1. Start of the month: run with `DRY_RUN = True` and read the summary.
2. Set `DRY_RUN = False`, run, type `YES`.
3. Ask enumerators to synchronise their tablets so they see the new quantities.
4. A second run in the same month does nothing for assignments already changed; assignments that failed
   are retried.

**Limit of this function:** it cannot see interview status. An interview that was created but left
unfinished, or one that was rejected and is waiting for correction, still counts as created. Keep your
usual monthly check on in-progress and rejected interviews.

---

## 6. Safety features

- `DRY_RUN = True` by default; a live run needs you to type `YES`.
- `MAX_CHANGES` stops a live run that would touch more assignments than expected.
- `rollover_state.csv` stops the same assignment being topped up twice in one cycle.
- A 401 (bad login) stops the run at once; any other error is logged for that assignment and the run continues.
- The log file records every assignment touched, including dry runs.
- The script never deletes anything. Archived assignments can be restored in HQ.

---

## 7. API calls used

All calls are under `{SERVER}/{WORKSPACE}/api/v1`, with Basic authentication.

| Function | Call | Status |
|---|---|---|
| export / close / archive / rollover (listing) | `GET /assignments?Limit=&Offset=` | Used in the same way as the interviews list |
| rollover | `PATCH /assignments/{id}/changeQuantity`, body = the new number as JSON (e.g. `4`) | **Confirmed working** on your server (swagger returned 200) |
| close | `PATCH /assignments/{id}/close` | Not yet confirmed on your server; check the verb on your `/apidocs/index.html` page |
| archive | `PATCH /assignments/{id}/archive` | Not yet confirmed on your server; same check |
| create | `POST /assignments`, JSON body with `Responsible`, `QuestionnaireId`, `Quantity`, `IdentifyingData` | Body matches community examples; not yet confirmed on your server |

The server's own list of calls is at `https://your-server.example.org/apidocs/index.html`
(not `/swagger`). Use the **Authorize** button on that page to log in before trying a call.

---

## 8. Troubleshooting

| Message | Likely cause and fix |
|---|---|
| `401 Unauthorized` | Wrong API user or password, or the user is not an API user in that workspace |
| `403` | The API user lacks permission for that action |
| `404` on a change | Wrong id, wrong workspace name, or a route that does not exist on your server version; check the id with the export file and the route on `/apidocs` |
| `405` | The server expects a different HTTP verb for that call; check `/apidocs` and tell the developer which one |
| `406` | The server refused the change for that assignment (for example a quantity below the interviews already created) |
| `415` (in Postman) | The request has no JSON body: Body tab → raw → JSON |
| `WARNING: listed X of Y assignments` | The listing did not return every assignment; do not run a live action until this is understood |
| `Excel file format cannot be determined` | The script handles this for the ids and create files; if you still see it, save the file again as `.xlsx` |
| `Column 'Responsible' not found` | Set `CREATE_RESPONSIBLE_COL` to the real column name in your file |

**Testing a single call in Postman:** method `PATCH`, Basic auth, Body → raw → JSON containing only the
number, for example `1`. A 200 or 204 means success.

---

## 9. Quick reference — common jobs

| Job | Settings |
|---|---|
| See everything | `ACTIONS = ["export"]` |
| Close everything that is open | `ACTIONS = ["close"]`, `SCOPE = "all"` |
| Close a list of ids | `ACTIONS = ["close"]`, `SCOPE = "selected"`, `IDS_FILE = "..."` |
| Archive all assignments of one old questionnaire | `ACTIONS = ["archive"]`, `FILTER_QUESTIONNAIRE = "<guid>"` |
| Create assignments for a new month or village list | `ACTIONS = ["create"]`, set the file and questionnaire, test with `CREATE_FIRST_N = 1` |
| Monthly top-up | `ACTIONS = ["rollover"]`, `SCHEDULES` filled in |
| Backup, then archive | `ACTIONS = ["export", "archive"]` |
