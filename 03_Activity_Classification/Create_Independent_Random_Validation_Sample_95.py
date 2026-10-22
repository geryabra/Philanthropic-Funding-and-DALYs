# -*- coding: utf-8 -*-
"""
Create_Independent_Random_Validation_Sample_95.py

Purpose:
Create an independent random validation sample of 95 organizations.

Method:
1. Load all unmatched organizations.
2. Keep only rows with non-empty Org_mission.
3. Exclude organizations already reviewed in the development/manual sample.
4. Draw a simple random sample without replacement.
5. Use a fixed random seed for full reproducibility.
6. Save an Excel file with:
   - Validation_Sample_95
   - Sampling_Report
   - Excluded_Development_EINs

Required input files in the same folder:
1. Org_Mission_Disease_Dictionary_Matches_v2_enriched.xlsx
   Sheet: unmatched_orgs

2. Book3_classified_with_organization_type.xlsx
   Sheet: Classification_Template

Output:
Independent_Random_Validation_Sample_95.xlsx
"""

import pandas as pd
import numpy as np
import re
from pathlib import Path


# =========================================================
# 1. PATHS
# =========================================================

BASE_DIR = Path().resolve()

UNMATCHED_FILE = BASE_DIR / "Org_Mission_Disease_Dictionary_Matches_v2_enriched.xlsx"
UNMATCHED_SHEET = "unmatched_orgs"

DEVELOPMENT_FILE = BASE_DIR / "Book3_classified_with_organization_type.xlsx"
DEVELOPMENT_SHEET = "Classification_Template"

OUTPUT_FILE = BASE_DIR / "Independent_Random_Validation_Sample_95.xlsx"


# =========================================================
# 2. PARAMETERS
# =========================================================

SAMPLE_SIZE = 95
RANDOM_STATE = 42

EIN_COL = "extract_ein"
MISSION_COL_CANDIDATES = ["Org_mission", "org_mission", "mission", "extract_org_mission"]
REVENUE_COL_CANDIDATES = ["Revenue", "990extract_revenue", "extract_revenue", "revenue", "total_revenue"]


# =========================================================
# 3. HELPERS
# =========================================================

def clean_ein(x):
    """
    Standardize EIN for reliable matching.
    """
    if pd.isna(x):
        return ""
    s = str(x).strip()
    s = re.sub(r"\.0$", "", s)
    s = re.sub(r"[^0-9]", "", s)
    return s


def find_col(df, candidates, required=True):
    """
    Find column by exact case-insensitive name from candidate list.
    """
    lower_to_original = {str(c).lower().strip(): c for c in df.columns}
    for cand in candidates:
        key = cand.lower().strip()
        if key in lower_to_original:
            return lower_to_original[key]
    if required:
        raise ValueError(f"Required column not found. Tried: {candidates}")
    return None


def safe_revenue(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)


# =========================================================
# 4. LOAD DATA
# =========================================================

print("Loading unmatched organizations...")
df_unmatched = pd.read_excel(UNMATCHED_FILE, sheet_name=UNMATCHED_SHEET)
df_unmatched.columns = [str(c).strip() for c in df_unmatched.columns]

print("Unmatched rows loaded:", len(df_unmatched))

print("Loading development/manual sample...")
df_dev = pd.read_excel(DEVELOPMENT_FILE, sheet_name=DEVELOPMENT_SHEET)
df_dev.columns = [str(c).strip() for c in df_dev.columns]

print("Development rows loaded:", len(df_dev))


# =========================================================
# 5. STANDARDIZE COLUMNS
# =========================================================

mission_col = find_col(df_unmatched, MISSION_COL_CANDIDATES)
revenue_col = find_col(df_unmatched, REVENUE_COL_CANDIDATES, required=False)

if mission_col != "Org_mission":
    df_unmatched["Org_mission"] = df_unmatched[mission_col]

if revenue_col is None:
    print("WARNING: Revenue column was not found. Revenue will be set to 0.")
    df_unmatched["_revenue_for_sampling"] = 0
else:
    df_unmatched["_revenue_for_sampling"] = safe_revenue(df_unmatched[revenue_col])

if EIN_COL not in df_unmatched.columns:
    raise ValueError(f"Required column '{EIN_COL}' not found in unmatched file.")

if EIN_COL not in df_dev.columns:
    raise ValueError(f"Required column '{EIN_COL}' not found in development sample.")

df_unmatched["_ein_clean"] = df_unmatched[EIN_COL].apply(clean_ein)
df_dev["_ein_clean"] = df_dev[EIN_COL].apply(clean_ein)


# =========================================================
# 6. DEFINE POPULATION
# =========================================================

df_unmatched["_mission_clean"] = df_unmatched["Org_mission"].astype(str).str.strip()

df_population = df_unmatched[
    df_unmatched["Org_mission"].notna()
    & (df_unmatched["_mission_clean"] != "")
    & (df_unmatched["_mission_clean"].str.lower() != "nan")
].copy()

total_unmatched_rows = len(df_unmatched)
population_with_mission = len(df_population)

dev_eins = set(
    df_dev["_ein_clean"]
    .dropna()
    .astype(str)
    .str.strip()
)
dev_eins = {x for x in dev_eins if x != ""}

df_eligible = df_population[
    ~df_population["_ein_clean"].isin(dev_eins)
].copy()

development_excluded_count = population_with_mission - len(df_eligible)

print("Rows with non-empty mission:", population_with_mission)
print("Development EINs excluded:", development_excluded_count)
print("Eligible rows after exclusion:", len(df_eligible))


# =========================================================
# 7. DRAW RANDOM SAMPLE
# =========================================================

if len(df_eligible) < SAMPLE_SIZE:
    raise ValueError(
        f"Not enough eligible organizations for sample size {SAMPLE_SIZE}. "
        f"Eligible rows: {len(df_eligible)}"
    )

df_sample = df_eligible.sample(
    n=SAMPLE_SIZE,
    replace=False,
    random_state=RANDOM_STATE
).copy()

df_sample = df_sample.sort_values("_revenue_for_sampling", ascending=False).reset_index(drop=True)
df_sample.insert(0, "random_sample_id", range(1, len(df_sample) + 1))

# Empty manual classification columns
df_sample["Main Theme"] = ""
df_sample["Organization Type"] = ""
df_sample["Confidence"] = ""
df_sample["Comments"] = ""

# Useful order of columns
preferred_cols = [
    "random_sample_id",
    "Revenue",
    "990extract_revenue",
    "extract_ein",
    "extract_name",
    "Org_mission",
    "NTEE",
    "bmf_NTEE_NCCS",
    "Main Theme",
    "Organization Type",
    "Confidence",
    "Comments",
]

existing_preferred = [c for c in preferred_cols if c in df_sample.columns]
remaining_cols = [c for c in df_sample.columns if c not in existing_preferred and not c.startswith("_")]

df_sample_out = df_sample[existing_preferred + remaining_cols].copy()


# =========================================================
# 8. SAMPLING REPORT
# =========================================================

sampling_report = pd.DataFrame([
    {"metric": "sampling_method", "value": "Simple random sampling without replacement"},
    {"metric": "random_seed", "value": RANDOM_STATE},
    {"metric": "requested_sample_size", "value": SAMPLE_SIZE},
    {"metric": "actual_sample_size", "value": len(df_sample_out)},
    {"metric": "total_unmatched_rows", "value": total_unmatched_rows},
    {"metric": "unmatched_rows_with_non_empty_mission", "value": population_with_mission},
    {"metric": "development_sample_rows_loaded", "value": len(df_dev)},
    {"metric": "unique_development_eins", "value": len(dev_eins)},
    {"metric": "development_rows_excluded_from_population", "value": development_excluded_count},
    {"metric": "eligible_rows_after_exclusion", "value": len(df_eligible)},
    {"metric": "sampling_unit", "value": "organization / extract_ein"},
    {"metric": "manual_review_basis", "value": "Org_mission; extract_name and existing metadata may be used if already present in dataset"},
])

excluded_eins = (
    df_dev[[EIN_COL, "_ein_clean"]]
    .drop_duplicates()
    .sort_values("_ein_clean")
    .reset_index(drop=True)
)


# =========================================================
# 9. SAVE OUTPUT
# =========================================================

print("Saving output...")

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df_sample_out.to_excel(writer, sheet_name="Validation_Sample_95", index=False)
    sampling_report.to_excel(writer, sheet_name="Sampling_Report", index=False)
    excluded_eins.to_excel(writer, sheet_name="Excluded_Development_EINs", index=False)

print("DONE")
print("Output file:", OUTPUT_FILE)

print("\nSampling report:")
print(sampling_report)
