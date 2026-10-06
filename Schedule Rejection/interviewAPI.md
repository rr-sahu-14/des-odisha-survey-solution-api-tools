# Survey Solutions Bulk Interview Actions

A single-file Python script that bulk **approves or rejects** [Survey Solutions](https://docs.mysurvey.solutions/) interviews through its HTTP API, using an API user's **login and password** (HTTP Basic authentication). No API token is needed.

You give it a list of `interview__key` values (the short codes such as `31-10-76-63`). It looks up each key's `interview__id` locally, then sends one API call per interview.

---

## Contents

- [Features](#features)
- [How it works](#how-it-works)
- [Requirements](#requirements)
- [Input files](#input-files)
- [Configuration](#configuration)
- [Usage](#usage)
- [Supported actions](#supported-actions)
- [Output](#output)
- [Troubleshooting](#troubleshooting)
- [API behaviour notes](#api-behaviour-notes)
- [Security](#security)
- [Limitations](#limitations)

---

## Features

- Bulk **approve / reject** as Supervisor, or **approve / reject / unapprove** as Headquarters.
- Works from `interview__key`, the code people see in the web interface, not the internal GUID.
- **One API call per interview.** Run time depends only on how many keys you list, not on the size of your survey.
- **Dry-run mode** (on by default) shows what would be sent without changing anything.
- **Local key-to-id lookup** from a Survey Solutions `.tab` export, with no network round-trip.
- Fixes the **GUID format mismatch** between `.tab` exports (no hyphens) and the API (hyphenated).
- Reads "Excel" key files that are really CSV, tab-delimited text or old `.xls` files with a wrong extension.
- Writes a **results CSV** with the outcome for every key.
- Stops immediately on `401` (wrong credentials) instead of failing every row.

---

## How it works

```
 District's Excel file            Survey Solutions .tab export
 (interview__key only)            (interview__key + interview__id)
          │                                    │
          └────────────┬───────────────────────┘
                       ▼
        match key -> id locally (in memory)
                       ▼
   for each interview:  PATCH /api/v1/interviews/{id}/<action>
                       ▼
        HTTP 200/204 = done · 406 = wrong status · 404 = not found
                       ▼
                 results CSV
```

The script does **not** check an interview's status before acting. It sends the action and lets the server decide. If the interview is not in a status that allows the action, the server answers `HTTP 406`. The script logs that row as `FAILED` and moves on. This matches what the web interface does.

The public REST API cannot resolve an `interview__key` to an `interview__id`: the list endpoint never returns the key. The key-to-id lookup therefore comes from a `.tab` data export, which contains both columns.

---

## Requirements

- Python 3.8 or newer
- An API user on your Survey Solutions server with permission for the action you want (Supervisor or Headquarters role)

```bash
pip install pandas requests openpyxl
pip install xlrd        # only if your keys file is an old-style .xls
```

---

## Input files

| File | Source | Required columns |
|---|---|---|
| **Keys file** (`KEYS_EXCEL_FILE`) | Excel sheet listing the interviews to act on | `interview__key`. If there is no such header, the first column is used. |
| **Lookup file** (`LOOKUP_TAB_FILE`) | Survey Solutions tabular data export (`.tab`, tab-delimited) | `interview__key` **and** `interview__id` |

Notes:

- The lookup file must be a **tab-delimited** export, not comma-delimited.
- The `.tab` export stores `interview__id` without hyphens (for example `11111111222233334444555555555555`). The script converts it to the hyphenated form the API requires.
- Duplicate keys in the keys file are removed, and the order is kept.
- Keys that are missing from the lookup file are skipped and listed in a warning.
- Use a recent export. An interview created after the export will not be in the lookup file.

---

## Configuration

Edit the CONFIG block at the top of `ss_interview_actions.py`. There are no command-line flags.

```python
SERVER = "https://your-server.example.org"   # base URL of your Survey Solutions server
WORKSPACE = "primary"                        # workspace name

KEYS_EXCEL_FILE = "district_interview_keys.xlsx"   # interview__key values to act on
LOOKUP_TAB_FILE = "interview__actions.tab"          # export with interview__key + interview__id

ACTION = "hq-reject"       # approve | reject | hq-approve | hq-reject | unapprove
COMMENT = "Price outlier, please recheck"          # shown to the enumerator / supervisor

DRY_RUN = True             # True = send nothing; False = send for real
```

| Setting | Meaning |
|---|---|
| `SERVER` | Base URL, with no trailing path and no workspace. |
| `WORKSPACE` | Workspace name. `primary` is the default workspace. |
| `ACTION` | Which action to perform (see [Supported actions](#supported-actions)). |
| `COMMENT` | Text sent with the action. It is optional: use `""` to send none. For rejections it is the reason the enumerator sees. |
| `DRY_RUN` | Safety switch. Leave it `True` until the output looks right. |

---

## Usage

1. Put both input files next to the script, or give full paths in the CONFIG block.
2. Set `ACTION` and `COMMENT`.
3. **Dry run first** (`DRY_RUN = True`):

   ```bash
   python ss_interview_actions.py
   ```

   You will be asked for the API user id and password. The script prints each key with its resolved id and the action it would send. Nothing is changed on the server.

4. Check the output. Each key should resolve to an id.
5. Set `DRY_RUN = False` and run again. For a first live run, try a short keys file of one or two interviews.

### Credentials

The script asks for the API user id and password with a plain `input()` prompt. It does not use `getpass`, because that does not work in every terminal or IDE. **The password is visible on screen while you type it.**

To avoid the prompt, set environment variables. They take priority over the prompt:

```bash
export SS_USER="api_user"
export SS_PASSWORD="your_password"
python ss_interview_actions.py
```

### Example dry-run output

```
Server  : https://your-server.example.org/primary
Action  : hq-reject
Keys    : 3  (from district_interview_keys.xlsx, matched against interview__actions.tab)
Comment : Price outlier, please recheck
Mode    : DRY RUN -- nothing will be sent

[1/3] key 31-10-76-63 -> id 11111111-2222-3333-4444-555555555555  WOULD HQ-REJECT
[2/3] key 31-55-85-50 -> id 21841331-e9bb-433a-bb48-cec1b4231cd4  WOULD HQ-REJECT
[3/3] key 19-46-71-64 -> id 500b9b83-1e9e-4f65-a0f1-72a93020841b  WOULD HQ-REJECT

Dry run complete. Would act on 3, skip 0.
```

---

## Supported actions

All actions are `PATCH` requests to `/api/v1/interviews/{id}/<endpoint>`.

| `ACTION` | Endpoint | Performed as | Interview must be in status |
|---|---|---|---|
| `approve` | `/approve` | Supervisor | Completed |
| `reject` | `/reject` | Supervisor | Completed |
| `hq-approve` | `/hqapprove` | Headquarters | Completed or ApprovedBySupervisor |
| `hq-reject` | `/hqreject` | Headquarters | Completed or ApprovedBySupervisor |
| `unapprove` | `/unapprove` | Headquarters | ApprovedByHeadquarters |

The API user must hold the role the action needs. Otherwise the server returns `403`.

### Interview status codes (reference)

| Code | Status | Code | Status |
|---|---|---|---|
| -1 | Deleted | 95 | Restarted |
| 0 | Restored | 100 | Completed |
| 20 | Created | 120 | ApprovedBySupervisor |
| 40 | SupervisorAssigned | 125 | RejectedByHeadquarters |
| 60 | InterviewerAssigned | 130 | ApprovedByHeadquarters |
| 65 | RejectedBySupervisor | | |
| 80 | ReadyForInterview | | |
| 85 | SentToCapi | | |

---

## Output

### Console

One line per key with its HTTP result, then a summary: `Success`, `Failed`, `Skipped`.

### Results CSV

A live run writes `interview_<action>_<YYYYMMDD_HHMMSS>.csv`. A dry run does not write a file.

| Column | Meaning |
|---|---|
| `interview__key` | Key from your keys file |
| `interview__id` | Resolved GUID (empty if the key was not found) |
| `result` | `OK`, `FAILED`, or `SKIPPED` |
| `message` | Plain-language meaning of the HTTP code |
| `server_detail` | First 300 characters of the server's error text, for failures |

### HTTP codes

| Code | Meaning | Script behaviour |
|---|---|---|
| 200 / 204 | Action applied | `OK` |
| 400 | Bad request | `FAILED`, continues |
| 401 | Login or password rejected | **Stops the whole run** |
| 403 | User not allowed to do this action | `FAILED`, continues |
| 404 | Interview id not found | `FAILED`, continues |
| 406 | Interview is in a status that does not allow this action | `FAILED`, continues |

---

## Troubleshooting

**Every key is skipped as "not found".**
The keys are not in the `.tab` file. Check that the export is recent and that it has both `interview__key` and `interview__id` columns.

**`'...' must have columns ['interview__id']`**
The lookup file is missing a required column. Re-export it from Survey Solutions as tabular data.

**`ValueError: Excel file format cannot be determined`**
The script avoids this by checking the file's real format. If you still see it, open the file and check what it really is.

**`old-style .xls file ... needs the 'xlrd' package`**
Run `pip install xlrd`.

**HTTP 401 on the first row.**
The user id or password is wrong, or the account is not an API user.

**HTTP 403.**
The API user lacks the role for that action. Supervisor actions need a Supervisor-level API user. HQ actions need a Headquarters-level API user.

**HTTP 406.**
The interview's current status does not allow the action. For example, `hq-approve` on an interview that is already `ApprovedByHeadquarters`. Check its status in the web interface.

**HTTP 404.**
The id from the `.tab` file does not exist on the server, for example an interview that was deleted after the export.

---

## API behaviour notes

Findings from testing against a live server, useful if you extend the script:

| Endpoint | What it returns |
|---|---|
| `GET /api/v1/interviews` | Paginated list (`Limit` / `Offset`). Each row has `InterviewId` and `Status`. It **never returns `interview__key`** and cannot be filtered by id. |
| `GET /api/v1/interviews/{id}` | Only the questionnaire answers (`{"Answers": [...]}`). **No status field.** |
| `PATCH /api/v1/interviews/{id}/<action>` | The action endpoints used by this script. |
| `POST /graphql` | An internal endpoint used by the Headquarters web grid. It returns `id`, `key` and `status` together and accepts Basic Auth. It is undocumented and may change between server versions. |

Authentication: HTTP Basic (`Authorization: Basic base64(user:password)`) is accepted by both the REST API and the GraphQL endpoint, even though the official documentation focuses on tokens.

---

## Security

- **Never commit credentials.** Use the prompt or environment variables, not values written into the script.
- **Never commit data files.** The keys file, the `.tab` export and the results CSVs hold interview identifiers. Add them to `.gitignore`:

  ```gitignore
  *.xlsx
  *.xls
  *.tab
  interview_*.csv
  status_debug.txt
  ```
- The password is shown on screen while it is typed (see [Credentials](#credentials)). Do not run the script where others can see your screen.
- Use a dedicated API user with only the role the job needs.
- Approvals and rejections change survey data. Use a dry run and a small live test first.

---

## Limitations

- The `.tab` export is a snapshot. Interviews created after it was generated cannot be resolved.
- The script does not check an interview's status before sending an action. A wrong status shows up as `406` per interview, not as a pre-run warning.
- One request is sent at a time, with a 0.2 s delay between requests (`DELAY`). This is gentle on the server and is not tuned for very large batches.
- The optional `responsibleId` parameter, used to reassign an interview on rejection, is not supported.
- Tested end to end against `hq-reject`. The other actions use the same code path and the documented endpoints.

---

## License

Add a license of your choice (for example MIT) before publishing.
