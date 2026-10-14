"""
============================================================
 Alorica → Qualtrics Ingestion Pipeline
 Customer Research & Insights — Avangrid/Iberdrola
============================================================
 Reads daily Alorica CSV files (one per OpCo) and ingests
 them into the corresponding Qualtrics survey using the
 same prepare_for_qualtrics pattern used for iQor.

 Two survey layouts:
   - UIL Group  (UIC, SCG, CNG, CMP): NPS, FCR, CSAT, I_C, C_K, CallReason
   - NYSEG / RGE                    : NPS, FCR, CSAT, E_H, C_E, CallReason

 Usage (Jupyter cell):
   results = run_pipeline("2026-10-13")          # specific date
   results = run_pipeline()                      # yesterday
   results = run_pipeline(upload=False)          # dry run — prep only, no upload
"""

import os
import glob
import requests
import pandas as pd
import urllib3
from difflib import get_close_matches
from datetime import datetime, timedelta, timezone

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============================================================
#  QUALTRICS CONFIGURATION
# ============================================================

API_TOKEN   = "dZZueEgbaPrCuSTXEtIp2sz0EqlJNxg93jYEod7U"
DATA_CENTER = "iad1"

SURVEYS_ID = {
    "RGE"  : "SV_cvxFLJeNk0tQm6q",
    "NYSEG": "SV_5dLTpj4MDbbc1bU",
    "CMP"  : "SV_ebvEco2ILaRjPqS",
    "CNG"  : "SV_cuK9q1z7FFGaiLc",
    "SCG"  : "SV_6fZBPa0TfJq4Xrg",
    "UIC"  : "SV_3R9Pe8Lcy5ZaIQW",
}

# ============================================================
#  FILE PATHS
# ============================================================

BASE_PATH = os.path.join(
    os.path.expanduser("~"),
    "OneDrive - IBERDROLA S.A",
    "General - Customer Research",
    "Post Call Survey",
    "Post Call Survey Data 2025"
)

OPCO_CONFIG = {
    "RGE": {
        "folder"     : os.path.join(BASE_PATH, "Alorica_RGE"),
        "file_prefix": "ALO RGE_Survey_Data",
        "layout"     : "rge_nyseg",
    },
    "NYSEG": {
        "folder"     : os.path.join(BASE_PATH, "Alorica_NYSEG"),
        "file_prefix": "ALO NYSEG_Survey_Data",
        "layout"     : "rge_nyseg",
    },
    "CMP": {
        "folder"     : os.path.join(BASE_PATH, "Alorica_UIL"),
        "file_prefix": "ALO CMP_Survey_Data",
        "layout"     : "uil",
    },
    "SCG": {
        "folder"     : os.path.join(BASE_PATH, "Alorica_UIL"),
        "file_prefix": "ALO SCG_Survey_Data",
        "layout"     : "uil",
    },
    "CNG": {
        "folder"     : os.path.join(BASE_PATH, "Alorica_UIL"),
        "file_prefix": "ALO CNG_Survey_Data",
        "layout"     : "uil",
    },
    "UIC": {
        "folder"     : os.path.join(BASE_PATH, "Alorica_UIL"),
        "file_prefix": "ALO UIC_Survey_Data",
        "layout"     : "uil",
    },
}

# ============================================================
#  COLUMN NORMALISATION
#  Source CSV headers → canonical internal names
# ============================================================

# UIL group layout  (UIC / SCG / CNG / CMP)
UIL_RENAME = {
    "User"        : "ExternalDataReference",
    "Agent"       : "Agent",
    "Date"        : "Date",
    "WorkGroup"   : "WorkGroup",
    "Conversation": "ConversationID",
    "Survey"      : "SurveyName",
    "Phone"       : "Phone",
    "NPS"         : "NPS",
    "FCR"         : "FCR",
    "CSAT"        : "CSAT",
    "I_C"         : "I_C",
    "C_K"         : "C_K",
    "Call Reason" : "CallReason",
    "Status"      : "Status",
    "Tag"         : "Tag",
}

# RGE / NYSEG layout
RGE_NYSEG_RENAME = {
    "ID"                : "ExternalDataReference",
    "Name"              : "Agent",
    "Date/Time"         : "Date",
    "Work Group"        : "WorkGroup",
    "Interaction ID"    : "ConversationID",
    "Survey Name"       : "SurveyName",
    "Phone Number"      : "Phone",
    "NPS"               : "NPS",
    "FCR"               : "FCR",
    "CSAT"              : "CSAT",
    "E_H"               : "E_H",
    "C_E"               : "C_E",
    "CallReason"        : "CallReason",
    "Survey Completion" : "Status",
    "Tag"               : "Tag",
}

LAYOUT_RENAME = {
    "uil"      : UIL_RENAME,
    "rge_nyseg": RGE_NYSEG_RENAME,
}

# ============================================================
#  QUALTRICS PREP  (same pattern as iQor / internal pipelines)
# ============================================================

def prepare_for_qualtrics(df, opco):
    """
    Fetches the Qualtrics survey schema, fuzzy-matches CSV columns
    to questionName, then builds the 3-row header structure required
    for Qualtrics CSV import.
    """
    survey_id = SURVEYS_ID[opco]
    url       = f"https://{DATA_CENTER}.qualtrics.com/API/v3/surveys/{survey_id}"
    headers   = {"X-API-TOKEN": API_TOKEN}

    response = requests.get(url, headers=headers, verify=False)
    response.raise_for_status()

    label_map       = response.json()["result"]["questions"]
    qid_to_name     = {qid: q["questionName"] for qid, q in label_map.items()}
    qid_to_text     = {qid: q["questionText"] for qid, q in label_map.items()}
    name_to_qid     = {name: qid for qid, name in qid_to_name.items()}
    qualtrics_names = list(name_to_qid.keys())

    # Step 1 — Fuzzy match columns to Qualtrics questionName
    mapped_cols = {}
    used_names  = set()
    for col in df.columns:
        match = get_close_matches(col, qualtrics_names, n=1, cutoff=0.6)
        if match:
            new_name = match[0]
            if new_name in used_names:
                new_name = col
            mapped_cols[col] = new_name
            used_names.add(new_name)
        else:
            mapped_cols[col] = col
    df = df.rename(columns=mapped_cols)

    # Step 2 — Build questionText row
    row2 = [
        qid_to_text.get(name_to_qid.get(col), "")
        if name_to_qid.get(col) else ""
        for col in df.columns
    ]

    # Step 3 — Build ImportId row
    row3 = [
        f'{{"ImportId": "{name_to_qid[col]}_TEXT"}}'
        if col in name_to_qid else ""
        for col in df.columns
    ]

    # Step 4 — Stack headers + data
    hdr1     = pd.DataFrame([list(df.columns)], columns=df.columns)
    hdr2     = pd.DataFrame([row2],             columns=df.columns)
    hdr3     = pd.DataFrame([row3],             columns=df.columns)
    df_final = pd.concat(
        [hdr1, hdr2, hdr3, df.reset_index(drop=True)],
        ignore_index=True
    )

    # Step 5 — Remove duplicate header row (hdr1 mirrors column names)
    df_final = df_final.iloc[1:].reset_index(drop=True)

    print(f"  [{opco}] Qualtrics prep complete | Shape: {df_final.shape}")
    return df_final


# ============================================================
#  QUALTRICS UPLOAD
# ============================================================

def upload_to_qualtrics(df_prepared, opco):
    """
    Uploads the prepared DataFrame to Qualtrics as a CSV import.
    Returns the progressId or None on failure.
    """
    survey_id = SURVEYS_ID[opco]
    url       = (
        f"https://{DATA_CENTER}.qualtrics.com/API/v3/surveys/"
        f"{survey_id}/import-responses"
    )
    headers   = {"X-API-TOKEN": API_TOKEN}
    csv_bytes = df_prepared.to_csv(index=False).encode("utf-8")

    response = requests.post(
        url,
        headers=headers,
        files={"file": ("responses.csv", csv_bytes, "text/csv")},
        verify=False
    )

    if response.status_code == 200:
        progress_id = response.json().get("result", {}).get("progressId", "N/A")
        print(f"  [{opco}] ✅ Upload accepted | progressId: {progress_id}")
        return progress_id
    else:
        print(f"  [{opco}] ❌ Upload failed | {response.status_code}: {response.text}")
        return None


# ============================================================
#  FILE FINDER
# ============================================================

def find_file(opco, date_str):
    """
    Looks for:  <prefix> <YYYY-MM-DD>.csv  or  <prefix> <YYYY-MM-DD>.xlsx
    in the configured folder for the given OpCo.
    """
    cfg    = OPCO_CONFIG[opco]
    folder = cfg["folder"]
    prefix = cfg["file_prefix"]

    for ext in ["csv", "xlsx"]:
        pattern = os.path.join(folder, f"{prefix} {date_str}.{ext}")
        matches = glob.glob(pattern)
        if matches:
            return matches[0]

    return None


# ============================================================
#  READ + NORMALISE
# ============================================================

def read_and_normalise(filepath, opco):
    """
    Reads the CSV/Excel file, renames columns to canonical names,
    and filters to Completed surveys only.
    """
    cfg    = OPCO_CONFIG[opco]
    layout = cfg["layout"]
    rename = LAYOUT_RENAME[layout]

    if filepath.endswith(".xlsx"):
        df = pd.read_excel(filepath)
    else:
        df = pd.read_csv(filepath)

    print(f"  [{opco}] Read {len(df)} rows from {os.path.basename(filepath)}")

    # Rename known columns
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})

    # Filter to completed surveys only
    if "Status" in df.columns:
        before = len(df)
        df = df[df["Status"].str.strip().str.lower() == "completed"].copy()
        print(f"  [{opco}] Filtered to Completed: {len(df)} / {before} rows")

    # Drop rows with no ConversationID
    if "ConversationID" in df.columns:
        df = df[df["ConversationID"].notna() & (df["ConversationID"].astype(str).str.strip() != "")].copy()

    return df


# ============================================================
#  SAVE PREPARED FILE  (audit trail / manual re-upload)
# ============================================================

def save_prepared(df_prepared, opco, date_str):
    cfg    = OPCO_CONFIG[opco]
    folder = cfg["folder"]
    prefix = cfg["file_prefix"]
    out    = os.path.join(folder, f"{prefix} {date_str}_Qualtrics.csv")
    df_prepared.to_csv(out, index=False)
    print(f"  [{opco}] 💾 Saved: {os.path.basename(out)}")
    return out


# ============================================================
#  MAIN PIPELINE
# ============================================================

def run_pipeline(date_str=None, opcos=None, upload=True):
    """
    Full ingestion pipeline for a given date.

    Args:
        date_str : "YYYY-MM-DD"  (default: yesterday)
        opcos    : list of OpCo keys, e.g. ["RGE", "NYSEG"]
                   (default: all six)
        upload   : False = prepare & save files only, no Qualtrics upload
    """
    if date_str is None:
        date_str = (datetime.now(timezone.utc) - timedelta(days=1)).strftime("%Y-%m-%d")

    if opcos is None:
        opcos = list(OPCO_CONFIG.keys())

    print("=" * 60)
    print(f"  Alorica → Qualtrics Ingestion")
    print(f"  Date   : {date_str}")
    print(f"  OpCos  : {', '.join(opcos)}")
    print(f"  Upload : {'YES' if upload else 'DRY RUN — prep only, no upload'}")
    print("=" * 60)

    results = {}

    for opco in opcos:
        print(f"\n▶ {opco}")

        # 1 — Locate file
        filepath = find_file(opco, date_str)
        if not filepath:
            print(f"  [{opco}] ⚠️  No file found for {date_str} — skipping")
            results[opco] = "FILE_NOT_FOUND"
            continue

        # 2 — Read + normalise
        try:
            df = read_and_normalise(filepath, opco)
        except Exception as e:
            print(f"  [{opco}] ❌ Read error: {e}")
            results[opco] = f"READ_ERROR: {e}"
            continue

        if df.empty:
            print(f"  [{opco}] ⚠️  No completed rows — skipping")
            results[opco] = "NO_DATA"
            continue

        # 3 — Prepare for Qualtrics
        try:
            df_prepared = prepare_for_qualtrics(df, opco)
        except Exception as e:
            print(f"  [{opco}] ❌ Prep error: {e}")
            results[opco] = f"PREP_ERROR: {e}"
            continue

        # 4 — Save prepared CSV for audit trail
        try:
            save_prepared(df_prepared, opco, date_str)
        except Exception as e:
            print(f"  [{opco}] ⚠️  Could not save prepared file: {e}")

        # 5 — Upload
        if upload:
            try:
                progress_id = upload_to_qualtrics(df_prepared, opco)
                results[opco] = "UPLOADED" if progress_id else "UPLOAD_FAILED"
            except Exception as e:
                print(f"  [{opco}] ❌ Upload error: {e}")
                results[opco] = f"UPLOAD_ERROR: {e}"
        else:
            results[opco] = "DRY_RUN_OK"

    # ── Summary ───────────────────────────────────────────────
    print("\n" + "=" * 60)
    print("  SUMMARY")
    print("=" * 60)
    for opco, status in results.items():
        icon = "✅" if status in ("UPLOADED", "DRY_RUN_OK") else "⚠️ "
        print(f"  {icon} {opco:8} : {status}")
    print("=" * 60)

    return results


# ============================================================
#  QUICK TEST  — run this cell first with upload=False
# ============================================================
# results = run_pipeline("2026-10-13", upload=False)
#
# When dry run looks good:
# results = run_pipeline("2026-10-13", upload=True)
#
# Single OpCo:
# results = run_pipeline("2026-10-13", opcos=["RGE"], upload=True)
