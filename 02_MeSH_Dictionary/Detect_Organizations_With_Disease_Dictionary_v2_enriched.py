







































































































































































































































































































































































































































































































































































































































































































































































































































































































































































































































































































































# -*- coding: utf-8 -*-
"""
Detect_Organizations_With_Disease_Dictionary_v2_enriched.py

Purpose:
Identify disease signals in organization missions using DALY_MeSH_Dictionary_v1,
while keeping the final classification results from FINAL_DATASET_STAGE1 / final_dataset
and enriching them with organization identifiers from FINAL_DATASET_STAGE1 / full_enriched_debug.

Inputs:
1. FINAL_DATASET_STAGE1.xlsx
   Main sheet: final_dataset
   Enrichment sheet: full_enriched_debug

2. DALY_MeSH_Dictionary_v1.xlsx
   Sheet: disease_master

Output:
Org_Mission_Disease_Dictionary_Matches_v2_enriched.xlsx

Main change vs v1:
- Uses final_dataset as the main analytical sheet.
- Adds organization-level metadata from full_enriched_debug:
  Unnamed: 0, extract_ein, extract_name, extract_state, extract_country,
  extract_ret_type, bmf_NTEE_NCCS.
"""

import pandas as pd
import re
from pathlib import Path


# =========================================================
# 1. PATHS
# =========================================================

BASE_DIR = Path().resolve()

ORG_FILE = BASE_DIR / "FINAL_DATASET_STAGE1.xlsx"

# Keep using final_dataset because it contains the final analytical classification.
ORG_SHEET = "final_dataset"

# Use this sheet only to enrich final_dataset with organization identifiers.
ENRICHED_SHEET = "full_enriched_debug"

DICT_FILE = BASE_DIR / "DALY_MeSH_Dictionary_v1.xlsx"
DICT_SHEET = "disease_master"

OUTPUT_FILE = BASE_DIR / "Org_Mission_Disease_Dictionary_Matches_v2_enriched.xlsx"


# =========================================================
# 2. PARAMETERS
# =========================================================

AUTO_CATEGORY_COL = "auto_category"
MISSION_COL = "Org_mission"

TARGET_CATEGORIES = ["B", "C"]

# Columns to bring from full_enriched_debug into final_dataset.
ENRICH_COLS = [
    "Unnamed: 0",
    "extract_ein",
    "extract_name",
    "extract_state",
    "extract_country",
    "extract_ret_type",
    "bmf_NTEE_NCCS",
]

MIN_TERM_LENGTH = 4

# Short medical acronyms allowed despite length < 4.
# "ms" is intentionally excluded because it can create false positives.
SHORT_TERM_WHITELIST = {
    "hiv",
    "aids",
    "als",
}

EXCLUDE_TERMS = {
    "",
    "disease",
    "diseases",
    "disorder",
    "disorders",
    "syndrome",
    "syndromes",
    "condition",
    "conditions",
}


# =========================================================
# 3. HELPERS
# =========================================================

def normalize_text(text):
    if pd.isna(text):
        return ""

    text = str(text).lower().strip()
    text = text.replace("&", " and ")
    text = text.replace("/", " ")
    text = text.replace("-", " ")
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def split_terms(value):
    if pd.isna(value):
        return []

    parts = re.split(r"[;|,]", str(value))
    return [p.strip() for p in parts if p and p.strip()]


def get_col(df, candidates, required=True):
    cols_lower = {str(c).lower().strip(): c for c in df.columns}

    for cand in candidates:
        key = cand.lower().strip()
        if key in cols_lower:
            return cols_lower[key]

    if required:
        raise ValueError(f"Could not find required column. Tried: {candidates}")

    return None


def compile_term_pattern(term_norm):
    return re.compile(r"\b" + re.escape(term_norm) + r"\b")


def join_unique(values):
    clean = []
    for v in values:
        if pd.isna(v):
            continue
        for part in str(v).split(";"):
            p = part.strip()
            if p:
                clean.append(p)
    return "; ".join(sorted(set(clean)))


def clean_enrich_cols(df, cols):
    existing_cols = [c for c in cols if c in df.columns]
    missing_cols = [c for c in cols if c not in df.columns]

    if missing_cols:
        print("\nWARNING: These enrichment columns were not found and will be skipped:")
        print(missing_cols)

    return existing_cols


# =========================================================
# 4. LOAD FILES
# =========================================================

print("Loading organizations from final_dataset...")
df_org = pd.read_excel(ORG_FILE, sheet_name=ORG_SHEET)
df_org.columns = [str(c).strip() for c in df_org.columns]

print("final_dataset rows loaded:", len(df_org))
print("final_dataset columns:")
print(df_org.columns.tolist())

print("\nLoading organization enrichment columns from full_enriched_debug...")
df_enriched = pd.read_excel(ORG_FILE, sheet_name=ENRICHED_SHEET)
df_enriched.columns = [str(c).strip() for c in df_enriched.columns]

print("full_enriched_debug rows loaded:", len(df_enriched))
print("full_enriched_debug columns:")
print(df_enriched.columns.tolist())

print("\nLoading disease dictionary...")
df_dict = pd.read_excel(DICT_FILE, sheet_name=DICT_SHEET)
df_dict.columns = [str(c).strip() for c in df_dict.columns]

print("Dictionary rows loaded:", len(df_dict))
print("Dictionary columns:")
print(df_dict.columns.tolist())


# =========================================================
# 5. ENRICH final_dataset WITH ORGANIZATION IDENTIFIERS
# =========================================================

# IMPORTANT:
# This merge assumes final_dataset and full_enriched_debug were created from the same
# source dataframe and preserve the same row order.
df_org["_merge_row_id"] = df_org.index
df_enriched["_merge_row_id"] = df_enriched.index

existing_enrich_cols = clean_enrich_cols(df_enriched, ENRICH_COLS)

df_org = df_org.merge(
    df_enriched[["_merge_row_id"] + existing_enrich_cols],
    on="_merge_row_id",
    how="left"
)

print("\nAfter enrichment:")
print("Rows:", len(df_org))
print("Added columns:", existing_enrich_cols)

# Basic validation of row-index enrichment.
if len(df_org) != len(df_enriched):
    print("\nWARNING: final_dataset and full_enriched_debug have different row counts.")
    print("This does not necessarily break the run, but index-based enrichment should be reviewed.")
    print("final_dataset rows:", len(df_org))
    print("full_enriched_debug rows:", len(df_enriched))


# =========================================================
# 6. VALIDATE COLUMNS
# =========================================================

auto_col = get_col(df_org, [AUTO_CATEGORY_COL, "auto_category_v4_1"])
mission_col = get_col(df_org, [MISSION_COL, "org_mission", "mission"])

daly_col = get_col(df_dict, ["daly_disease", "Disease", "disease"])
mesh_id_col = get_col(df_dict, ["mesh_id", "MeSH ID", "DescriptorUI"], required=False)
preferred_col = get_col(
    df_dict,
    ["preferred_term_from_mesh", "mesh_preferred_term", "preferred_term"],
    required=False
)
entry_terms_col = get_col(df_dict, ["all_entry_terms", "entry_terms"], required=False)
matched_term_col = get_col(df_dict, ["matched_term"], required=False)
mapping_group_col = get_col(df_dict, ["mapping_group"], required=False)
handling_decision_col = get_col(df_dict, ["handling_decision"], required=False)


# =========================================================
# 7. FILTER ORGANIZATIONS B + C
# =========================================================

df_org["_source_row_id"] = df_org.index
df_org["_auto_category_clean"] = df_org[auto_col].astype(str).str.strip()

df_bc = df_org[df_org["_auto_category_clean"].isin(TARGET_CATEGORIES)].copy()
df_bc["_mission_norm"] = df_bc[mission_col].apply(normalize_text)

print("\nFiltered organizations:")
print(df_bc["_auto_category_clean"].value_counts(dropna=False))
print("Total B+C organizations:", len(df_bc))


# =========================================================
# 8. BUILD TERM DICTIONARY
# =========================================================

term_rows = []
skipped_rows = []

for _, row in df_dict.iterrows():

    daly_disease = row.get(daly_col, "")
    mesh_id = row.get(mesh_id_col, "") if mesh_id_col else ""
    preferred = row.get(preferred_col, "") if preferred_col else ""
    mapping_group = row.get(mapping_group_col, "") if mapping_group_col else ""
    handling_decision = row.get(handling_decision_col, "") if handling_decision_col else ""

    candidates = []

    if pd.notna(daly_disease) and str(daly_disease).strip():
        candidates.append(("daly_disease", str(daly_disease).strip()))

    if preferred_col and pd.notna(preferred) and str(preferred).strip():
        candidates.append(("preferred_term", str(preferred).strip()))

    if matched_term_col and pd.notna(row.get(matched_term_col)) and str(row.get(matched_term_col)).strip():
        candidates.append(("matched_term", str(row.get(matched_term_col)).strip()))

    if entry_terms_col and pd.notna(row.get(entry_terms_col)):
        for term in split_terms(row.get(entry_terms_col)):
            candidates.append(("entry_term", term))

    for source, term in candidates:
        term_norm = normalize_text(term)

        skip_reason = ""

        if term_norm in EXCLUDE_TERMS:
            skip_reason = "excluded_generic_term"
        elif len(term_norm) < MIN_TERM_LENGTH and term_norm not in SHORT_TERM_WHITELIST:
            skip_reason = "too_short"
        elif len(term_norm.split()) == 1 and term_norm.isdigit():
            skip_reason = "numeric_only"

        if skip_reason:
            skipped_rows.append({
                "term": term,
                "term_norm": term_norm,
                "skip_reason": skip_reason,
                "match_source": source,
                "daly_disease": daly_disease,
                "mesh_id": mesh_id,
                "preferred_term_from_mesh": preferred
            })
            continue

        term_rows.append({
            "term": term,
            "term_norm": term_norm,
            "term_length": len(term_norm),
            "term_word_count": len(term_norm.split()),
            "match_source": source,
            "daly_disease": daly_disease,
            "mesh_id": mesh_id,
            "preferred_term_from_mesh": preferred,
            "mapping_group": mapping_group,
            "handling_decision": handling_decision
        })

df_terms = pd.DataFrame(term_rows)

if len(df_terms) == 0:
    raise ValueError("No usable terms were created from the dictionary. Check dictionary columns.")

df_terms = (
    df_terms
    .drop_duplicates(subset=["term_norm", "daly_disease"])
    .sort_values(["term_word_count", "term_length"], ascending=False)
    .reset_index(drop=True)
)

df_skipped = pd.DataFrame(skipped_rows)

print("\nUsable dictionary terms:", len(df_terms))
print("Skipped terms:", len(df_skipped))
print("Unique normalized terms:", df_terms["term_norm"].nunique())
print("Unique DALY diseases in dictionary terms:", df_terms["daly_disease"].nunique())


compiled_terms = []
for _, t in df_terms.iterrows():
    compiled_terms.append((compile_term_pattern(t["term_norm"]), t.to_dict()))


# =========================================================
# 9. MATCH ORGANIZATION MISSIONS
# =========================================================

print("\nMatching organization missions...")

raw_matches = []

for n, (_, org_row) in enumerate(df_bc.iterrows(), start=1):

    if n % 250 == 0:
        print(f"Processed organizations: {n} / {len(df_bc)}")

    mission_norm = org_row["_mission_norm"]

    if not mission_norm:
        continue

    org_id_data = {
        col: org_row.get(col, "")
        for col in existing_enrich_cols
        if col in df_bc.columns
    }

    for pattern, term_data in compiled_terms:

        if pattern.search(mission_norm):

            raw_matches.append({
                "source_row_id": org_row["_source_row_id"],
                **org_id_data,
                "auto_category": org_row["_auto_category_clean"],
                "Org_mission": org_row.get(mission_col, ""),
                "matched_term": term_data["term"],
                "matched_term_norm": term_data["term_norm"],
                "match_source": term_data["match_source"],
                "term_length": term_data["term_length"],
                "term_word_count": term_data["term_word_count"],
                "daly_disease": term_data["daly_disease"],
                "mesh_id": term_data["mesh_id"],
                "preferred_term_from_mesh": term_data["preferred_term_from_mesh"],
                "mapping_group": term_data["mapping_group"],
                "handling_decision": term_data["handling_decision"]
            })

df_raw_matches = pd.DataFrame(raw_matches)

print("Raw term matches:", len(df_raw_matches))


# =========================================================
# 10. COLLAPSE TO ORG + DISEASE LEVEL
# =========================================================

id_group_cols = [c for c in existing_enrich_cols if c in df_raw_matches.columns]

if len(df_raw_matches) > 0:

    df_org_disease_matches = (
        df_raw_matches
        .groupby(
            [
                "source_row_id",
                *id_group_cols,
                "auto_category",
                "Org_mission",
                "daly_disease",
                "mesh_id",
                "preferred_term_from_mesh",
                "mapping_group",
                "handling_decision"
            ],
            dropna=False,
            as_index=False
        )
        .agg(
            matched_terms=("matched_term", join_unique),
            matched_terms_norm=("matched_term_norm", join_unique),
            match_sources=("match_source", join_unique),
            raw_term_match_count=("matched_term", "count")
        )
    )

else:
    df_org_disease_matches = pd.DataFrame(columns=[
        "source_row_id", *existing_enrich_cols, "auto_category", "Org_mission",
        "daly_disease", "mesh_id", "preferred_term_from_mesh",
        "mapping_group", "handling_decision",
        "matched_terms", "matched_terms_norm", "match_sources",
        "raw_term_match_count"
    ])


matched_source_ids = set(df_org_disease_matches["source_row_id"].unique()) if len(df_org_disease_matches) else set()

df_unmatched = df_bc[~df_bc["_source_row_id"].isin(matched_source_ids)].copy()


# =========================================================
# 11. SUMMARIES
# =========================================================

if len(df_org_disease_matches) > 0:

    org_level_group_cols = [
        "source_row_id",
        *[c for c in existing_enrich_cols if c in df_org_disease_matches.columns],
        "auto_category",
        "Org_mission"
    ]

    df_org_level_summary = (
        df_org_disease_matches
        .groupby(org_level_group_cols, dropna=False, as_index=False)
        .agg(
            matched_disease_count=("daly_disease", "nunique"),
            matched_diseases=("daly_disease", join_unique),
            matched_mesh_ids=("mesh_id", join_unique),
            matched_terms=("matched_terms", join_unique)
        )
    )

    df_disease_summary = (
        df_org_disease_matches
        .groupby(["daly_disease", "mesh_id", "preferred_term_from_mesh"], dropna=False, as_index=False)
        .agg(
            matched_org_count=("source_row_id", "nunique"),
            total_org_disease_matches=("source_row_id", "count"),
            matched_terms=("matched_terms", join_unique)
        )
        .sort_values("matched_org_count", ascending=False)
    )

else:
    df_org_level_summary = pd.DataFrame()
    df_disease_summary = pd.DataFrame()


total_bc = len(df_bc)
matched_orgs = len(matched_source_ids)
unmatched_orgs = len(df_unmatched)
coverage_pct = round((matched_orgs / total_bc * 100), 2) if total_bc else 0

coverage_summary = pd.DataFrame([
    {"metric": "total_BC_orgs", "value": total_bc},
    {"metric": "matched_BC_orgs", "value": matched_orgs},
    {"metric": "unmatched_BC_orgs", "value": unmatched_orgs},
    {"metric": "coverage_percent", "value": coverage_pct},
    {"metric": "raw_term_matches", "value": len(df_raw_matches)},
    {"metric": "org_disease_matches", "value": len(df_org_disease_matches)},
    {"metric": "unique_DALY_diseases_matched", "value": df_org_disease_matches["daly_disease"].nunique() if len(df_org_disease_matches) else 0},
    {"metric": "dictionary_terms_used", "value": len(df_terms)},
    {"metric": "unique_dictionary_terms_norm", "value": df_terms["term_norm"].nunique()},
    {"metric": "skipped_terms", "value": len(df_skipped)},
    {"metric": "enrichment_columns_added", "value": len(existing_enrich_cols)}
])

category_summary = (
    df_bc
    .assign(has_disease_dictionary_match=df_bc["_source_row_id"].isin(matched_source_ids))
    .groupby(["_auto_category_clean", "has_disease_dictionary_match"], as_index=False)
    .size()
    .rename(columns={"_auto_category_clean": "auto_category", "size": "org_count"})
)


# =========================================================
# 12. SAVE OUTPUT
# =========================================================

print("\nSaving output...")

cols_to_drop = ["_mission_norm", "_merge_row_id"]

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    coverage_summary.to_excel(writer, sheet_name="coverage_summary", index=False)
    category_summary.to_excel(writer, sheet_name="category_summary", index=False)
    df_org_disease_matches.to_excel(writer, sheet_name="org_disease_matches", index=False)
    df_raw_matches.to_excel(writer, sheet_name="raw_term_matches", index=False)
    df_org_level_summary.to_excel(writer, sheet_name="org_level_summary", index=False)
    df_disease_summary.to_excel(writer, sheet_name="disease_summary", index=False)
    df_unmatched.drop(columns=cols_to_drop, errors="ignore").to_excel(writer, sheet_name="unmatched_orgs", index=False)
    df_terms.to_excel(writer, sheet_name="term_dictionary_used", index=False)
    df_skipped.to_excel(writer, sheet_name="skipped_terms", index=False)

print("DONE")
print("Output file:", OUTPUT_FILE)

print("\nCoverage summary:")
print(coverage_summary)

print("\nCategory summary:")
print(category_summary)

print("\nTop matched diseases:")
print(df_disease_summary.head(20) if len(df_disease_summary) else "No matches found")

print("\nEnrichment columns included:")
print(existing_enrich_cols)
