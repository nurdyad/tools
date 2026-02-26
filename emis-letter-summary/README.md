# EMIS Letter Summary (Standalone)

Production-ready standalone tool to:

- resolve practice by name (via Mailroom)
- fetch EMIS login details (CDB, username, password)
- auto-login to EMIS
- scan Workflow Manager -> Document Management queue(s)
- output letter type counts (JSON + optional CSV)

## Folder

`C:\Tools\emis-letter-summary`

## Prerequisites

- Windows
- Python 3.11 installed (`py -3.11`)
- Access to EMIS app and CDB switcher
- Mailroom API URL + API key

## One-time install

```powershell
cd C:\Tools\emis-letter-summary
.\install.ps1
```

This creates `.venv` and installs dependencies in this folder only.

## Configure

Create `.env` (copy from `.env.example`) with:

```env
MAILROOM_API_URL=https://<your-mailroom>/api
MAILROOM_API_KEY=<your-api-key>
PRACTICE_NAME=<Practice display name>
```

Optional overrides:

```env
ODS_CODE=
EMIS_CDB=
EMIS_USERNAME=
EMIS_PASSWORD=
```

## Run

### Practice name (recommended)

```powershell
cd C:\Tools\emis-letter-summary
.\run.bat --practice "Ash Trees Surgery" --queues "Awaiting Filing" --csv-output .\output\summary.csv --strict-sidebar-match
```

### Manual credentials (no Mailroom lookup)

```powershell
.\run.bat --practice "Ash Trees Surgery" --cdb "YOUR_CDB" --emis-username "YOUR_USER" --emis-password "YOUR_PASS" --queues "Awaiting Filing"
```

### Attach to existing EMIS session only

```powershell
.\run.bat --skip-login --queues "Awaiting Filing"
```

## Useful flags

- `--queues all`
- `--strict-sidebar-match` (exit code 2 if sidebar total != scanned rows)
- `--no-kill-existing`
- `--no-clear-cache`
- `--close-on-exit`
- `--quiet`

## Output

- JSON: `C:\Tools\emis-letter-summary\output\emis_letter_type_summary_<timestamp>.json`
- CSV: path provided via `--csv-output`
- Logs: `C:\Tools\emis-letter-summary\logs\`

## UI (desktop)

Launch:

```powershell
cd C:\Tools\emis-letter-summary
.\run_ui.bat
```

UI lets you enter:

- Practice Name or ODS Code
- Queue list
- CSV output path
- Mailroom URL/API key
- runtime flags (`Strict Sidebar Match`, `Skip Login`, `No Kill Existing`, `No Clear Cache`)

Then click `Run` to execute and watch live logs.
