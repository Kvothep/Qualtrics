# iQor NY PCS ETL Pipeline — Code Block Reference

**File:** `iQor_NY_PCS_ETL_Blocks.ipynb`  
**Purpose:** Pulls daily post-call survey reports from iQor's SFTP server for three OpCos (CMP, RGE, NSE), transforms the data, and uploads it to the corresponding Qualtrics surveys.

---

## Table of Contents

1. [Environment Setup](#0-environment-setup)
2. [Library Import](#library-import)
3. [Qualtrics Credentials](#qualtrics-credentials)
4. [iQor SFTP Credentials](#iqor-sftp-credentials)
5. [Popup Notifications](#popup-notifications)
6. [SharePoint Folder Mapping](#sharepoint-folder-mapping)
7. [Extract — SFTP File Pull](#1-extract)
8. [Save Raw CSVs Locally](#save-raw-csvs-locally)
9. [Transform — Column Rename & Cast](#2-transform)
10. [Prepare — Validation & Tagging](#prepare-for-qualtrics-step-1-validation--tagging)
11. [Prepare — Qualtrics Header Mapping](#prepare-for-qualtrics-step-2-qualtrics-api-mapping)
12. [Load — CSV Writer](#3-load)
13. [Date Logic](#source-folder--date-logic)
14. [Upload to Qualtrics](#upload-to-qualtrics)
15. [Execute](#execute-cell)

---

## 0. Environment Setup

### Library Import

**What it does:** Loads all Python libraries needed across the full pipeline.

**Input:** None

**Output:** Libraries available in memory

| Library | Purpose |
|---|---|
| `pandas` | DataFrame operations |
| `datetime` | Date math for file routing |
| `re` | Regex for column detection |
| `requests` | HTTP calls to SFTP server and Qualtrics API |
| `json` | Parse Qualtrics API responses |
| `urllib3` | SSL warning suppression |
| `difflib.get_close_matches` | Fuzzy column name matching |
| `ctypes`, `threading`, `time` | Windows popup notifications |
| `os` | File path operations |
| `urllib.parse.quote` | URL-encode filenames with spaces |
| `io.StringIO` | Read CSV text as file-like object |

```python
import pandas as pd
from datetime import date, timedelta
import re
import requests
import json
import urllib3
from difflib import get_close_matches
import ctypes
import threading
import time
import os
from urllib.parse import quote
requests.packages.urllib3.disable_warnings()
from io import StringIO
```

---

### Qualtrics Credentials

**What it does:** Defines the API token, data center, and the survey ID for each OpCo.

**Input:** None (hardcoded — see security note)

**Output:** `API_TOKEN`, `DATA_CENTER`, `SURVEYS_ID` dict available globally

> **Security Note:** These values are exposed in plain text. Move them to environment variables (`os.environ`) or a `.env` file with `python-dotenv` before sharing or deploying this notebook.

| Variable | Description |
|---|---|
| `API_TOKEN` | Qualtrics API authentication token |
| `DATA_CENTER` | Qualtrics data center identifier (e.g. `iad1`) |
| `SURVEYS_ID` | Maps each OpCo code to its Qualtrics survey ID |

```python
API_TOKEN   = "..."
DATA_CENTER = "iad1"
SURVEYS_ID  = {
    "CMP": "SV_...",
    "RGE": "SV_...",
    "NSE": "SV_..."
}
```

---

### iQor SFTP Credentials

**What it does:** Defines the base URL, username, and password for the iQor managed file transfer server.

**Input:** None (hardcoded — see security note)

**Output:** `BASE_URL`, `USERNAME`, `PASSWORD` available globally

> **Security Note:** Same concern as above — externalize before sharing.

---

### Popup Notifications

**What it does:** Defines helper functions to show Windows message box dialogs for success and error states. Uses a background thread so Jupyter doesn't block, and auto-closes after a timeout.

**Input:**
- `message` (str) — Text to display
- `title` (str) — Window title
- `timeout` (int) — Seconds before auto-close (default: 10)
- `is_error` (bool) — True = red error icon, False = blue info icon

**Output:** A Windows popup dialog (non-blocking)

> **Platform Note:** This only works on Windows. On macOS/Linux or Azure, replace with email or Teams notifications.

```python
popup_info("Upload complete for CMP")
popup_error("File not found on SFTP", title="ETL Error")
```

---

### SharePoint Folder Mapping

**What it does:** Resolves the local OneDrive/SharePoint sync path for each OpCo to save output files. Falls back gracefully if the folder doesn't exist (e.g. running on a machine without SharePoint synced).

**Input:**
- `opco` (str) — One of `"CMP"`, `"RGE"`, `"NSE"`
- `filename` (str) — Target filename

**Output:**
- `get_sharepoint_folder(opco)` → full folder path (str)
- `get_output_path(filename, opco)` → tuple `(full_path, folder_exists_bool)`

| OpCo | SharePoint Subfolder |
|---|---|
| CMP | `iQor_CMP` |
| RGE | `iQor_RGE` |
| NSE | `iQor_NYSEG` |

```python
path, exists = get_output_path("CMP Daily Survey Report_20260615.csv", "CMP")
# → ('C:\Users\...\OneDrive - IBERDROLA S.A\iQor-Avangrid - General\iQor_CMP\CMP Daily Survey Report_20260615.csv', True)
```

---

## 1. Extract

### SFTP File Pull

**What it does:**
1. Opens an authenticated HTTP session to iQor's managed file transfer server
2. Lists files in each OpCo's survey report folder
3. Identifies the most recent `Daily Survey Report_YYYYMMDD.csv` (excludes `Triage` files)
4. Downloads the file content as raw text
5. Parses it into a pandas DataFrame

**Input:**
- iQor SFTP credentials (`BASE_URL`, `USERNAME`, `PASSWORD`)
- `SFTP_FOLDERS` dict mapping each OpCo to its remote directory path

**Output:**
- `dataframes` dict — `{ "CMP": df_cmp, "RGE": df_rge, "NSE": df_nse }`
- `raw_files` dict — `{ "CMP": (filename, raw_csv_text), ... }`
- Convenience variables: `df_cmp`, `df_rge`, `df_nse`

**Key helpers:**

| Function | Description |
|---|---|
| `list_dir(path)` | Parses the SFTP directory listing, returns `(files, folders)` |
| `get_latest_file(opco)` | Gets the most recent report for an OpCo, returns `(df, filename, raw_text)` |

**Example output shape:**
```
[CMP] Latest file : CMP Daily Survey Report_20260615.csv
[CMP] Server date : Jun 16 07:00:24
[CMP] Shape       : (83, 15)
```

> **Note:** The session warmup (`time.sleep(2)`) after the first GET is intentional — the iQor server requires a brief pause to fully establish the session before directory listing works reliably.

---

### Save Raw CSVs Locally

**What it does:** Saves the raw (untransformed) CSV files to the local SharePoint sync folder for audit purposes and manual pivot table use.

**Input:**
- `raw_files` dict — raw text content per OpCo from the extract step
- Requires SharePoint folder to be synced locally

**Output:** CSV files written to `OneDrive - IBERDROLA S.A\iQor-Avangrid - General\iQor_{OpCo}\`

**Failure modes handled:**
- Folder not found → skips with warning
- File locked by Excel → skips with warning

> **Deployment Note:** This cell is annotated `# LOCAL ONLY — Remove when deploying to Azure`. It's a convenience for local runs only.

---

## 2. Transform

### Column Rename & Cast

**What it does:**
1. Renames columns to consistent names (CMP uses different column names than RGE/NSE in the source file)
2. Standardizes the `Date/Time` format to `MM/DD/YYYY HH:MM:SS`
3. Casts score columns to nullable integers (`Int64`)

**Input:**
- `df` (DataFrame) — raw DataFrame from extract
- `opco` (str) — `"CMP"`, `"RGE"`, or `"NSE"`

**Output:** Transformed DataFrame with standardized column names and types

**Rename maps:**

| OpCo | Original → Standard |
|---|---|
| CMP | `Date_Time` → `Date/Time`, `CSAT1` → `CSAT`, `Survey_Status` → `Survey Completion` |
| RGE | `Date` → `Date/Time`, `Telephone` → `Phone Number` |
| NSE | `Date` → `Date/Time`, `Telephone` → `Phone Number` |

**Score columns cast to `Int64`:**

| OpCo | Columns |
|---|---|
| CMP | `NPS`, `CSAT`, `Call_Reason`, `Survey_Status_Count` |
| RGE/NSE | `NPS`, `FCR`, `CSAT`, `E_H`, `C_E`, `CallReason` |

```python
df_cmp = transform(df_cmp, "CMP")
# [CMP] Transform complete | Shape: (83, 15)
```

---

### Prepare for Qualtrics — Step 1: Validation & Tagging

**What it does:**
1. Warns (via popup) if any unexpected columns are detected
2. Validates that all required score columns are present (raises if missing)
3. Drops fully empty columns
4. Forward-fills `ID` and `Name` (handles multi-row survey responses where these repeat)
5. Tags rows where `Work Group` contains "test" (case-insensitive) as `"Test"`, others as `""`

**Input:**
- `df` (DataFrame) — transformed DataFrame
- `opco` (str) — `"CMP"`, `"RGE"`, or `"NSE"`

**Output:** DataFrame with validated columns + new `Tag` column

**Example:**
```
[CMP] Prepare complete | Shape: (83, 16)
```

> **Known Issue:** `EXPECTED_COLS` includes `company` and `report_date` which are commented out upstream, so the validator will always warn about them. Either add those columns or remove them from `EXPECTED_COLS`.

---

### Prepare for Qualtrics — Step 2: Qualtrics API Mapping

**What it does:**
1. Calls the Qualtrics Survey API to get the official question names, question text, and QIDs for the target survey
2. Fuzzy-matches DataFrame column names to Qualtrics `questionName` fields (cutoff: 0.6 similarity)
3. Renames columns to their official Qualtrics names
4. Builds a 3-row header block required by the Qualtrics Import Responses API:
   - Row 1: `questionName` (column headers)
   - Row 2: `questionText` (human-readable question labels)
   - Row 3: `{"ImportId": "QIDx_TEXT"}` (Qualtrics internal field mapping)
5. Stacks headers + data into the final upload-ready DataFrame

**Input:**
- `df` (DataFrame) — validated/tagged DataFrame from Step 1
- `opco` (str) — used to look up the correct `SURVEYS_ID`

**Output:** Upload-ready DataFrame with 2 extra header rows prepended

**Output shape:** `(original_rows + 2, columns)`

**Example header block:**
```
Row 0 (col names): ID | Name | Date/Time | Work Group | ...
Row 1 (ImportId):  {"ImportId": "QID1_TEXT"} | {"ImportId": "QID2_TEXT"} | ...
Row 2 (data):      38999530 | kevin.dallum | 06/15/2026 07:43:25 | ...
```

> **Important — Function Name Collision:** In the notebook, a function also named `prepare_for_qualtrics` is defined earlier (Step 1 above). When this cell runs, it **overwrites** the Step 1 function. Rename Step 1 to `validate_and_tag()` to avoid this.

---

## 3. Load

### CSV Writer

**What it does:** Writes a DataFrame to a CSV file at the given path.

**Input:**
- `df` (DataFrame) — any DataFrame (raw, transformed, or Qualtrics-ready)
- `output_path` (str) — full file path including filename

**Output:** CSV file written to disk (UTF-8, no index)

```python
load_data(df_cmp_q, "CMP Daily Survey Report_20260615_qualtrics.csv")
```

---

### Source Folder & Date Logic

**What it does:** Computes the effective date for file routing. Handles the Monday edge case: since no survey reports are generated on Sunday, Monday's pipeline run should look for Saturday's file.

**Input:** System date (`date.today()`)

**Output:** Date variables used to build file paths

| Variable | Example | Description |
|---|---|---|
| `source_year` | `2026` | Year for folder path |
| `source_month` | `"June"` | Month name for folder path |
| `source_day` | `"15"` | Zero-padded day for folder path |
| `landing_yesterday` | `2026-06-14` | Previous day — used in output filename |

---

### Upload to Qualtrics

**What it does:** POSTs the Qualtrics-ready CSV file to the Qualtrics Import Responses API and prints the server response.

**Input:**
- `file_path` (str) — path to the Qualtrics-formatted CSV temp file
- `DATA_CENTER` (str) — Qualtrics data center ID
- `SURVEY_ID` (str) — target survey ID
- `API_TOKEN` (str) — Qualtrics API token

**Output:** API response dict (or raw text if JSON parsing fails)

**API endpoint:** `POST /API/v3/surveys/{SURVEY_ID}/import-responses`

**Successful response example:**
```json
{
    "result": {
        "progressId": "f6b4eb2b-375f-4424-91d9-016eb072ab69",
        "percentComplete": 0.0,
        "status": "inProgress"
    },
    "meta": {
        "httpStatus": "200 - OK"
    }
}
```

> **Missing Step:** A `200 - OK` response with `"status": "inProgress"` means the job was **accepted, not completed**. You should poll `GET /API/v3/surveys/{SURVEY_ID}/import-responses/{progressId}` until `percentComplete` reaches `100` or status becomes `"complete"`. Without this, the pipeline can't detect partial upload failures.

---

## Execute Cell

**What it does:** Orchestrates the full pipeline end-to-end for a single run.

> **Known Issue:** This cell currently calls `extract_data()`, `transform_data()`, and `get_output_path(filename)` — functions that **do not exist** in this notebook (they belong to `Int_NY_PCS_ETL_Blocks.ipynb`). This cell will fail. It needs to be rewritten to call the correct iQor functions: `get_latest_file()`, `transform()`, `prepare_for_qualtrics()`, `load_data()`, and `upload_to_qualtrics()` for each OpCo.

**Intended flow:**
```
For each OpCo (CMP, RGE, NSE):
  1. get_latest_file(opco)          → raw DataFrame
  2. transform(df, opco)            → standardized DataFrame
  3. validate_and_tag(df, opco)     → validated + tagged DataFrame
  4. prepare_for_qualtrics(df, opco)→ upload-ready DataFrame
  5. load_data(df, output_path)     → save to SharePoint
  6. load_data(df_q, temp_path)     → save temp upload file
  7. upload_to_qualtrics(temp_path, ...) → POST to Qualtrics
```

---

## Issues Summary

| # | Severity | Issue |
|---|---|---|
| 1 | Critical | Execute cell calls non-existent functions (`extract_data`, `transform_data`) |
| 2 | Critical | Two functions named `prepare_for_qualtrics` — second silently overwrites first |
| 3 | High | Upload is not confirmed — `progressId` is never polled to check completion |
| 4 | High | Credentials hardcoded in plain text |
| 5 | Medium | `SURVEY_STATUS_COL` dict is defined but never used |
| 6 | Medium | `EXPECTED_COLS` lists `company`/`report_date` which are commented out upstream |
| 7 | Low | SSL verification disabled globally |
| 8 | Low | No deduplication guard against running twice in one day |
| 9 | Low | No retry logic for HTTP calls |
| 10 | Low | Popup notifications are Windows-only (no graceful fallback) |
