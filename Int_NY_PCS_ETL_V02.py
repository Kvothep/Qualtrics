"""
Avangrid NY Post-Call Survey ETL pipeline.

Extracts the daily survey export from the network share, cleans/tags it,
and uploads it to Qualtrics via the Import Responses API. Converted from
Int_NY_PCS_ETL_V02.ipynb so it can run standalone (and be packaged as a
.exe with PyInstaller -- see build.ps1).
"""

import configparser
import ctypes
import json
import os
import sys
import re
import threading
import time
from datetime import date, timedelta
from difflib import get_close_matches

import pandas as pd
import requests
import urllib3


# ============================================================
# CONFIGURATION
# ============================================================

# Update this (and rebuild) whenever a new version is cut, so an old,
# unmaintained .exe doesn't keep running indefinitely with a possibly
# revoked token.
EXPIRATION_DATE = date(2027, 1, 30)


def check_expiration():
    """Shows a popup and exits if this build is past its expiration date.
    Runs before config/token loading on purpose, so an expired .exe never
    even prompts about a missing config.ini."""
    if date.today() >= EXPIRATION_DATE:
        message = (
            "This version of NY_PCS_ETL has expired (valid through "
            f"{EXPIRATION_DATE.strftime('%b %d, %Y')}). Please request an "
            "updated version from Alejandro."
        )
        ctypes.windll.user32.MessageBoxW(0, message, "Expired", 0x10)  # MB_ICONERROR
        sys.exit(0)


check_expiration()


def get_app_dir():
    """Folder to look for external files (config.ini) in: the .exe's own
    folder when frozen by PyInstaller, or this script's folder otherwise.
    Deliberately NOT the current working directory (the .exe can be
    launched from anywhere) and NOT sys._MEIPASS (PyInstaller's temp
    extraction folder, which is wiped after each run)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def load_api_token():
    """Reads the Qualtrics API token from config.ini next to the script/exe
    instead of hardcoding it, so the token isn't baked into the .exe binary."""
    config_path = os.path.join(get_app_dir(), "config.ini")

    if not os.path.exists(config_path):
        raise RuntimeError(
            f"config.ini not found at: {config_path}\n\n"
            "Create a file named config.ini in that same folder containing:\n\n"
            "[qualtrics]\n"
            "api_token = YOUR_TOKEN_HERE"
        )

    parser = configparser.ConfigParser()
    parser.read(config_path)
    token = parser.get("qualtrics", "api_token", fallback="").strip()

    if not token:
        raise RuntimeError(
            f"config.ini at {config_path} is missing the api_token value.\n\n"
            "Add this to config.ini:\n\n"
            "[qualtrics]\n"
            "api_token = YOUR_TOKEN_HERE"
        )

    return token


API_TOKEN = load_api_token()
DATA_CENTER = "iad1"
SURVEY_ID = "SV_efytbaOtuAX2uF0"


# ============================================================
# Popup notifications (Windows only)
# ============================================================

def popup_info(message, title="Success", timeout=10):
    MB_OK = 0x0
    MB_ICONINFORMATION = 0x40

    def show_box():
        ctypes.windll.user32.MessageBoxW(0, message, title, MB_OK | MB_ICONINFORMATION)

    t = threading.Thread(target=show_box)
    t.start()

    time.sleep(timeout)

    hwnd = ctypes.windll.user32.FindWindowW(None, title)
    if hwnd:
        ctypes.windll.user32.PostMessageW(hwnd, 0x0010, 0, 0)  # WM_CLOSE


def popup_error(message, title="Error", timeout=10):
    MB_OK = 0x0
    MB_ICONERROR = 0x10

    def show_box():
        ctypes.windll.user32.MessageBoxW(0, message, title, MB_OK | MB_ICONERROR)

    t = threading.Thread(target=show_box)
    t.start()

    time.sleep(timeout)

    hwnd = ctypes.windll.user32.FindWindowW(None, title)
    if hwnd:
        ctypes.windll.user32.PostMessageW(hwnd, 0x0010, 0, 0)


# ============================================================
# SharePoint / OneDrive output folder
# ============================================================

def get_onedrive_path():
    return os.path.join(os.path.expanduser("~"), "OneDrive - IBERDROLA S.A")


def get_sharepoint_folder():
    onedrive_root = get_onedrive_path()
    return os.path.join(
        onedrive_root,
        "General - Customer Research",
        "Post Call Survey",
        "Post Call Survey Data 2025",
        "Avangrid_NY"
    )


def get_output_path(filename):
    sharepoint_folder = get_sharepoint_folder()

    if os.path.exists(sharepoint_folder):
        return os.path.join(sharepoint_folder, filename), True

    return filename, False


# ============================================================
# 1. Extract
# ============================================================

def resolve_input_file(folder, base_name):
    """Finds the daily export in `folder` whether it landed as .xls or
    .xlsx -- pandas picks the right engine (xlrd / openpyxl) automatically
    based on the extension it's given."""

    for ext in (".xls", ".xlsx"):
        candidate = os.path.join(folder, base_name + ext)
        if os.path.exists(candidate):
            return candidate

    raise FileNotFoundError(
        f"No '{base_name}.xls' or '{base_name}.xlsx' found in: {folder}"
    )


def extract_data(input_path):
    """Reads the daily export (.xls or .xlsx) and detects each column's
    meaning from its sample value / question-text row, instead of assuming
    a fixed column order."""

    # Read the full sheet (no assumptions about columns)
    raw = pd.read_excel(input_path, header=None, dtype=str)

    # Row 7 contains the question text
    question_row = raw.iloc[6].fillna("").astype(str)

    # Data starts at row 9
    data = raw.iloc[8:].reset_index(drop=True)

    # Prepare final column names list
    final_cols = []

    for col_idx, col_series in data.items():
        sample_value = col_series.dropna().astype(str).iloc[0] if col_series.dropna().size > 0 else ""
        question_text = question_row[col_idx].lower()

        # -------------------------
        # METADATA COLUMN DETECTION
        # -------------------------

        # ID
        if re.match(r"^[UE]\d+", sample_value):
            final_cols.append("ID")
            continue

        # Name (column immediately right of ID)
        if len(final_cols) > 0 and final_cols[-1] == "ID":
            final_cols.append("Name")
            continue

        # Date/Time (column immediately right of Name)
        if len(final_cols) > 0 and final_cols[-1] == "Name":
            final_cols.append("Date/Time")
            continue

        # InteractionID (alphanumeric)
        if re.match(r"^[A-Za-z0-9]{12,}$", sample_value) and not sample_value.startswith("+"):
            final_cols.append("InteractionID")
            continue

        # Phone Number
        if sample_value.startswith("+"):
            final_cols.append("Phone Number")
            continue

        # Survey Name
        if any(x in sample_value.lower() for x in ["survey", "surv"]):
            final_cols.append("Survey Name")
            continue

        # Work Group
        if any(x in sample_value.lower() for x in ["cc", "vendor", "new"]):
            final_cols.append("Work Group")
            continue

        # -------------------------
        # SCORING COLUMN DETECTION
        # -------------------------

        qt = question_text  # shorthand

        if "recommend" in qt:
            final_cols.append("NPS")
            continue

        if "resolve" in qt or "call back" in qt:
            final_cols.append("FCR")
            continue

        if "help" in qt:
            final_cols.append("E_H")
            continue

        if any(x in qt for x in ["clear", "explain", "explaine"]):
            final_cols.append("C_E")
            continue

        if "satisfied" in qt:
            final_cols.append("CSAT")
            continue

        if any(x in qt for x in ["payment", "billing", "outage"]):
            final_cols.append("Call Reason")
            continue

        # If nothing matches, mark as Unknown
        final_cols.append(f"Unknown_{col_idx}")

    # Apply the detected column names
    data.columns = final_cols

    # Convert Date/Time to proper format
    if "Date/Time" in data.columns:
        data["Date/Time"] = pd.to_datetime(data["Date/Time"], errors="coerce")
        data["Date/Time"] = data["Date/Time"].dt.strftime("%m/%d/%Y %H:%M:%S")

    # Convert scoring columns to integers
    score_cols = ["NPS", "FCR", "E_H", "C_E", "CSAT", "Call Reason"]

    for col in score_cols:
        if col in data.columns:
            data[col] = pd.to_numeric(data[col], errors="coerce").astype("Int64")

    return data


# ============================================================
# 2. Transform
# ============================================================

def transform_data(df):
    """Validates, reorders columns, and computes Survey Status / Tag."""

    # 0. Remove empty columns (copy first to avoid SettingWithCopyWarning
    # on the inserts/assignments below)
    df = df.dropna(axis=1, how="all").copy()

    # 1. Detect unknown columns
    expected_cols = ["ID", "Name", "Date/Time", "InteractionID", "Phone Number",
                      "Survey Name", "Work Group", "NPS", "FCR", "E_H",
                      "C_E", "CSAT", "Call Reason"]

    unknown_cols = [c for c in df.columns if c not in expected_cols]

    if unknown_cols:
        message = f"Unknown columns detected:\n{unknown_cols}"
        ctypes.windll.user32.MessageBoxW(0, message, "ETL Warning", 0x40)

    # 2. Validate scoring columns
    required_cols = ["NPS", "FCR", "E_H", "C_E", "CSAT", "Call Reason"]
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required scoring columns: {missing}")

    # 3. Validate Work Group and CSAT exist before popping
    if "Work Group" not in df.columns:
        raise ValueError("Column 'Work Group' is missing from the dataset.")

    if "CSAT" not in df.columns:
        raise ValueError("Column 'CSAT' is missing from the dataset.")

    # Move Columns
    csat = df.pop("CSAT")
    work_group = df.pop("Work Group")

    df.insert(3, "Work Group", work_group)
    df.insert(7, "CSAT", csat)

    # Null handling of Name and ID
    df[["ID", "Name"]] = df[["ID", "Name"]].ffill()

    # Create Survey Status/Completion column
    df["Survey Status"] = df[required_cols].notna().all(axis=1)
    df["Survey Status"] = df["Survey Status"].map({True: "Complete", False: "Abandoned"})

    # Tag column
    df["Tag"] = df["Work Group"].str.contains("test", case=True, na=False).map({True: "Test", False: ""})

    # Convert scoring fields into integers
    for col in required_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").astype("Int64")

    return df


# ============================================================
# Prepare -- Qualtrics API field mapping
# ============================================================

def prepare_for_qualtrics(df):
    """Fuzzy-matches df columns to official Qualtrics questionName values
    and builds the 3-row header block required by the Import Responses API."""

    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    url = f"https://{DATA_CENTER}.qualtrics.com/API/v3/surveys/{SURVEY_ID}"

    headers = {
        "X-API-TOKEN": API_TOKEN
    }

    response = requests.get(url, headers=headers, verify=False)
    response.raise_for_status()

    # Convert API response to JSON
    survey_json = response.json()

    # Pull the Qualtrics question metadata
    label_map = survey_json["result"]["questions"]

    # QID -> questionName / questionText
    qid_to_name = {qid: q["questionName"] for qid, q in label_map.items()}
    qid_to_text = {qid: q["questionText"] for qid, q in label_map.items()}

    # questionName -> QID
    name_to_qid = {name: qid for qid, name in qid_to_name.items()}

    # All official Qualtrics labels (questionName)
    qualtrics_names = list(name_to_qid.keys())

    # 1) Fuzzy-match df columns to Qualtrics questionName
    mapped_cols = {}
    used_names = set()

    for col in df.columns:
        match = get_close_matches(col, qualtrics_names, n=1, cutoff=0.6)
        if match:
            new_name = match[0]
            if new_name in used_names:
                new_name = col  # avoid duplicates
            mapped_cols[col] = new_name
            used_names.add(new_name)
        else:
            mapped_cols[col] = col  # leave unmapped as-is

    df = df.rename(columns=mapped_cols)

    # 2) Row 2: questionText (aligned to final column names)
    row2 = []
    for col in df.columns:
        qid = name_to_qid.get(col)
        row2.append(qid_to_text.get(qid, "") if qid else "")

    # 3) Row 3: {"ImportId": "QIDx_TEXT"} as STRING
    row3 = []
    for col in df.columns:
        qid = name_to_qid.get(col)
        if qid:
            row3.append(f'{{"ImportId": "{qid}_TEXT"}}')
        else:
            row3.append("")

    # 4) Stack as extra rows (no new columns)
    hdr1 = pd.DataFrame([list(df.columns)], columns=df.columns)  # row 1: questionName
    hdr2 = pd.DataFrame([row2], columns=df.columns)              # row 2: questionText
    hdr3 = pd.DataFrame([row3], columns=df.columns)              # row 3: ImportId/QID_TEXT

    df_final = pd.concat([hdr1, hdr2, hdr3, df.reset_index(drop=True)], ignore_index=True)

    # Remove duplicated header row and reset index
    df_final = df_final.iloc[1:].reset_index(drop=True)

    return df_final


# ============================================================
# 3. Load
# ============================================================

def load_data(df, output_path):
    df.to_csv(output_path, index=False, encoding="utf-8")


# ============================================================
# Date logic -- source folder routing
# ============================================================

def get_effective_date(today):
    """Returns (source_year, source_month, source_day, landing_yesterday).

    Handles the Monday edge case: if today is Monday, use last Saturday
    since Sunday has no reports.
    """
    if today.weekday() == 0:
        effective_date = today - timedelta(days=2)
    else:
        effective_date = today

    source_year = effective_date.year
    source_month = effective_date.strftime("%B")
    source_day = effective_date.strftime("%d")
    landing_yesterday = effective_date - timedelta(days=1)

    return source_year, source_month, source_day, landing_yesterday


# ============================================================
# Upload to Qualtrics
# ============================================================

def upload_to_qualtrics(file_path, data_center, survey_id, api_token):
    """Uploads a CSV file to Qualtrics using the Import Responses API.
    Expects a fully formatted Qualtrics-ready CSV at file_path."""

    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    url = f"https://{data_center}.qualtrics.com/API/v3/surveys/{survey_id}/import-responses"

    headers = {
        "X-API-TOKEN": api_token,
        "Content-Type": "text/csv",
        "charset": "UTF-8"
    }

    print(f"\nUploading file to Qualtrics: {file_path}")

    with open(file_path, "rb") as f:
        response = requests.post(
            url,
            headers=headers,
            data=f,
            verify=False
        )

    print("\n=== RAW RESPONSE TEXT ===")
    print(response.text)
    print("=========================\n")

    try:
        result = response.json()
        print("Upload response (parsed):")
        print(json.dumps(result, indent=4))
        return result
    except Exception:
        print("Could not parse JSON response.")
        return response.text


# ============================================================
# Execute
# ============================================================

def main():
    today = date.today()
    filename = None

    try:
        source_year, source_month, source_day, landing_yesterday = get_effective_date(today)

        source_folder = (
            rf"\\clornas01\DIGITAL_COE_CS_DATA\data_delivery\qualtrics\NY_post_call_survey"
            rf"\daily\{source_year}\{source_month}\{source_day}"
        )
        input_file_path = resolve_input_file(source_folder, "NY Feedback Daily")

        filename = f"NY Feedback Daily {landing_yesterday}.csv"
        output_file_path, used_sharepoint = get_output_path(filename)

        temp_file_path = "NY_qualtrics_upload.csv"  # local temp file

        data = extract_data(input_file_path)
        cleaned_data = transform_data(data)
        load_data(cleaned_data, output_file_path)
        ready_data = prepare_for_qualtrics(cleaned_data)
        load_data(ready_data, temp_file_path)
        upload_to_qualtrics(temp_file_path, DATA_CENTER, SURVEY_ID, API_TOKEN)

    except Exception as e:
        error_message = f"Error running the program:\n{e}"

        if today.weekday() == 0:
            error_message += "\n\nMonday run - Friday folder expected."
        elif today.weekday() == 6:
            error_message += "\n\nWeekend run - folder may be empty."

        popup_error(error_message)
        sys.exit(1)

    else:
        popup_info(filename, "Upload successful")


if __name__ == "__main__":
    main()
