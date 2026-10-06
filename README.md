# Survey Solutions API Tools

Python scripts that automate bulk work on a [Survey Solutions](https://docs.mysurvey.solutions/)
server through its HTTP API, replacing repetitive clicking in the Headquarters web interface.

| Tool | What it does |
|---|---|
| [`ss_assignment_tool.py`](ss_assignment_tool.py) | Manage **assignments**: export, create, close, archive, and a monthly quantity top-up (rollover) |
| [`Schedule Rejection/schedule_action.py`](Schedule%20Rejection/schedule_action.py) | Bulk **approve / reject** interviews from a list of `interview__key` values |

Full guides: [assignment tool](README_ss_assignment_tool.md) · [interview actions](Schedule%20Rejection/interviewAPI.md)

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

## Repository layout

```
ss_assignment_tool.py          assignment manager
Schedule Rejection/            bulk interview approve / reject
local_config.example.py        template for your private settings
requirements.txt
```

## Data and privacy

Data files (`.csv`, `.xlsx`, `.tab`), exports, logs and `local_config.py` are listed in `.gitignore`
and are never committed.
