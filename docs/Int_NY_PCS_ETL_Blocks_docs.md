# Internal NY PCS ETL Pipeline — Code Block Reference

**File:** `Int_NY_PCS_ETL_Blocks.ipynb`  
**Purpose:** Reads the daily post-call survey Excel file from the internal network share (`\\clornas01\...`), detects and standardizes columns by content pattern, transforms the data, and uploads it to the NY Qualtrics survey.

> **Relationship to iQor notebook:** This notebook handles the *internal* data source (an `.xls` file from a network drive). The `iQor_NY_PCS_ETL_Blocks.ipynb` handles the *iQor* SFTP source. Both upload to Qualtrics but target different surveys and use different extract strategies.

---

## Table of Contents

1. [Environment Setup](#0-environment-setup)
2. [Library Import](#library-import)
3. [Qualtrics Credentials](#qualtrics-credentials)
4. [Popup Notifications](#popup-notifications)
5. [SharePoint Folder Mapping](#sharepoint-folder-mapping)
6. [Extract — Excel File Reader](#1-extract)
7. [Transform — Validation, Reorder & Status](#2-transform)
8. [Prepare — Qualtrics Header Mapping](#prepare-format-for-qualtrics)
9. [Load — CSV Writer](#3-load)
10. [Date Logic](#source-folder--date-logic)
11. [Upload to Qualtrics](#upload-to-qualtrics)
12. [Execute](#execute)

---

## 0. Environment Setup

### Library Import

**What it does:** Loads all Python libraries needed for the pipeline.

**Input:** None

**Output:** Libraries available in memory

| Library | Purpose |
|---|---|
| `pandas` | DataFrame operations |
| `datetime` | Date math for folder routing |
| `re` | Regex for column content detection |
| `requests` | Qualtrics API calls |
| `json` | Parse API responses |
| `urllib3` | SSL warning suppression |
| `difflib.get_close_matches` | Fuzzy column name matching |
| `ctypes`, `threading`, `time` | Windows popup dialogs |
| `os` | File path resolution |

---

### Qualtrics Credentials

**What it does:** Defines the API token, data center, and survey ID for the single NY survey.

**Input:** None (hardcoded)

**Output:** `API_TOKEN`, `DATA_CENTER`, `SURVEY_ID` available globally

> **Security Note:** Move these to environment variables before sharing this file.

---

### Popup Notifications

**What it does:** Shows Windows message box dialogs for success/error states. Each function spawns a background thread for the dialog box, sleeps for the timeout, then sends a close message to auto-dismiss it.

**Input:**
- `message` (str) — body text
- `title` (str) — window title
- `timeout` (int) — seconds before auto-close

**Output:** Windows dialog box (non-blocking)

> **Note vs iQor version:** This version uses `time.sleep(timeout)` in the main thread, which does block Jupyter for the duration of the timeout. The iQor notebook's version uses `t.join(timeout=timeout)` which is slightly better.

> **Platform Note:** Windows-only. Will crash on macOS/Linux.

---

### SharePoint Folder Mapping

**What it does:** Resolves the local OneDrive sync path where cleaned output files are saved.

**Input:**
- `filename` (str) — target filename

**Output:** Tuple `(full_path, folder_exists_bool)`

**Hardcoded target folder:**
```
OneDrive - IBERDROLA S.A\
  General - Customer Research\
    Post Call Survey\
      Post Call Survey Data 2025\
        Avangrid_NY\
```

> **Note vs iQor version:** This function has no `opco` parameter — it always maps to the same folder. The iQor version parameterizes this by OpCo.

---

## 1. Extract

### Excel File Reader (`extract_data`)

**What it does:** Reads the daily `.xls` file from the network share without assuming any fixed column order. Instead, it detects each column's semantic meaning by examining the cell content and the question text row, then assigns consistent standard column names.

**Input:**
- `input_path` (str) — full UNC path to the `.xls` file  
  e.g. `\\clornas01\DIGITAL_COE_CS_DATA\...\NY Feedback Daily.xls`

**Output:** DataFrame with standardized columns:

| Column | Detection Method |
|---|---|
| `ID` | First value matches `[UE]\d+` pattern |
| `Name` | Column immediately after `ID` |
| `Date/Time` | Column immediately after `Name` |
| `InteractionID` | 12+ alphanumeric chars, not starting with `+` |
| `Phone Number` | Value starts with `+` |
| `Survey Name` | Value contains "survey" or "surv" |
| `Work Group` | Value contains "cc", "vendor", or "new" |
| `NPS` | Question text contains "recommend" |
| `FCR` | Question text contains "resolve" or "call back" |
| `E_H` | Question text contains "help" |
| `C_E` | Question text contains "clear", "explain", or "explaine" |
| `CSAT` | Question text contains "satisfied" |
| `Call Reason` | Question text contains "payment", "billing", or "outage" |
| `Unknown_N` | Anything not matched — flagged for review |

**Source file structure assumed:**
- Row 7 (index 6): Question text labels
- Row 9+ (index 8+): Data rows

**Post-processing within extract:**
- `Date/Time` → parsed and reformatted to `MM/DD/YYYY HH:MM:SS`
- Score columns (`NPS`, `FCR`, `E_H`, `C_E`, `CSAT`, `Call Reason`) → cast to nullable `Int64`

```python
data = extract_data(r"\\clornas01\...\NY Feedback Daily.xls")
# Returns DataFrame with columns: ID, Name, Date/Time, InteractionID, Phone Number,
#   Survey Name, Work Group, NPS, FCR, E_H, C_E, CSAT, Call Reason
```

> **Resilience note:** If a column doesn't match any pattern, it's labeled `Unknown_N` rather than crashing. This lets the pipeline continue while flagging the issue for manual review.

---

## 2. Transform

### Validation, Reorder & Status (`transform_data`)

**What it does:**
1. Drops fully empty columns
2. Warns via popup if any unexpected columns were detected in extract
3. Validates all required score columns are present (raises `ValueError` if missing)
4. Reorders columns: moves `Work Group` to position 3, `CSAT` to position 7
5. Forward-fills `ID` and `Name` (handles multi-row agent entries)
6. Computes `Survey Status`: `"Complete"` if all score columns are filled, `"Abandoned"` otherwise
7. Creates `Tag` column: `"Test"` if `Work Group` contains "test" (case-sensitive), else `""`
8. Re-casts score columns to `Int64`

**Input:**
- `df` (DataFrame) — output from `extract_data()`

**Output:** Cleaned, reordered DataFrame with two new columns: `Survey Status`, `Tag`

**Final column order:**
```
ID | Name | Date/Time | Work Group | InteractionID | Phone Number |
Survey Name | CSAT | NPS | FCR | E_H | C_E | Call Reason | Survey Status | Tag
```

> **Known Warning:** Three `SettingWithCopyWarning` messages will appear:
> ```
> A value is trying to be set on a copy of a slice from a DataFrame.
> ```
> This happens because the function receives a DataFrame and mutates it without first doing `df = df.copy()`. Add `df = df.copy()` at the top of the function to eliminate these warnings.

> **Note — Case Sensitivity on Tag:** The `Tag` detection here uses `case=True` (case-sensitive "test"), while the iQor notebook uses `case=False` (case-insensitive). Align these to avoid missing test rows.

---

## Prepare Format for Qualtrics

### Qualtrics API Mapping (`prepare_for_qualtrics`)

**What it does:**
1. Calls the Qualtrics Survey API to retrieve official question names (`questionName`), question text (`questionText`), and QIDs
2. Fuzzy-matches each DataFrame column name to the Qualtrics `questionName` values (cutoff: 0.6 similarity)
3. Renames DataFrame columns to their official Qualtrics names
4. Builds the 3-row header block required by the Qualtrics Import Responses API:
   - Row 1: `questionName` (becomes the CSV column headers)
   - Row 2: `questionText` (human-readable question labels)
   - Row 3: `{"ImportId": "QIDx_TEXT"}` (Qualtrics internal mapping)
5. Stacks the headers + data into the final upload-ready DataFrame
6. Displays a preview of the first 5 rows

**Input:**
- `df` (DataFrame) — transformed/tagged DataFrame from `transform_data()`
- Uses global `DATA_CENTER`, `SURVEY_ID`, `API_TOKEN`

**Output:** Upload-ready DataFrame — shape: `(original_rows + 2, columns)`

**Qualtrics 3-row header structure:**
```
Row 0: ID | Name | Date/Time | Work Group | ... | Survey Status | Tag
Row 1: {"ImportId":"QID1_TEXT"} | {"ImportId":"QID2_TEXT"} | ...
Row 2: U357967 | Kathryn Tiesi | 06/15/2026 11:24:40 | ...
```

> **Note:** The function displays the first 5 rows for visual verification before returning.

---

## 3. Load

### CSV Writer (`load_data`)

**What it does:** Writes a DataFrame to a CSV file.

**Input:**
- `df` (DataFrame) — any DataFrame
- `output_path` (str) — full file path

**Output:** CSV file written to disk (UTF-8 encoding, no index)

Two files are created per run:
1. **Repository file** — cleaned/transformed CSV saved to SharePoint folder (for audit)
2. **Temp upload file** — Qualtrics-ready CSV at a fixed local path (`NY_qualtrics_upload.csv`)

---

### Source Folder & Date Logic

**What it does:** Determines the effective date for building the network share path to the source `.xls` file. Handles the Monday edge case where Sunday has no data.

**Input:** System date (`date.today()`)

**Output:** Path component variables

| Variable | Example | Description |
|---|---|---|
| `source_year` | `2026` | Year folder |
| `source_month` | `"June"` | Month name folder |
| `source_day` | `"15"` | Zero-padded day folder |
| `landing_yesterday` | `2026-06-14` | Used in output filename |

**Monday handling:** If today is Monday (`weekday() == 0`), effective date is set to last Saturday (`today - 2 days`), since Sunday has no reports.

---

### Upload to Qualtrics

**What it does:** POSTs the Qualtrics-ready CSV to the Qualtrics Import Responses API.

**Input:**
- `file_path` (str) — local path to the temp Qualtrics CSV
- `DATA_CENTER`, `SURVEY_ID`, `API_TOKEN` — Qualtrics credentials

**Output:** API response dict

**API endpoint:** `POST /API/v3/surveys/{SURVEY_ID}/import-responses`

**Response on success:**
```json
{
    "result": {
        "progressId": "...",
        "percentComplete": 0.0,
        "status": "inProgress"
    }
}
```

> **Missing Step:** `200 - OK` + `"status": "inProgress"` means the job was accepted, not finished. The `progressId` should be polled via `GET /API/v3/surveys/{SURVEY_ID}/import-responses/{progressId}` until completion. Without this check, silent partial upload failures can go undetected.

---

## Execute

**What it does:** Runs the full ETL pipeline end-to-end in a single try/except block.

**Flow:**
```
1. Build source path from date variables
2. extract_data(input_path)           → raw DataFrame
3. transform_data(raw_df)             → cleaned DataFrame
4. load_data(cleaned, sharepoint)     → save repository file
5. prepare_for_qualtrics(cleaned)     → upload-ready DataFrame
6. load_data(ready, temp_path)        → save temp upload file
7. upload_to_qualtrics(temp_path)     → POST to Qualtrics API
```

**On success:** `popup_info(filename, "Upload successful")` — shows the filename that was uploaded  
**On failure:** `popup_error(error_message)` — shows the exception with weekday context notes

**Source path pattern:**
```
\\clornas01\DIGITAL_COE_CS_DATA\data_delivery\qualtrics\
  NY_post_call_survey\daily\{year}\{month}\{day}\NY Feedback Daily.xls
```

---

## Issues Summary

| # | Severity | Issue |
|---|---|---|
| 1 | High | `SettingWithCopyWarning` in `transform_data` — missing `df = df.copy()` at top |
| 2 | High | Upload not confirmed — `progressId` never polled for completion status |
| 3 | High | Credentials hardcoded in plain text |
| 4 | Medium | `Tag` detection is case-sensitive (`"test"` not `"Test"`) — may miss test rows |
| 5 | Medium | `popup_error`/`popup_info` use `time.sleep()` in main thread — blocks notebook |
| 6 | Medium | SSL verification disabled globally |
| 7 | Low | Windows-only popup with no fallback for other environments |
| 8 | Low | No deduplication guard against running twice in one day |
| 9 | Low | `Unknown_N` columns are silently included in the output — consider raising or excluding |
