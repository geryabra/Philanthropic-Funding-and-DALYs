# -*- coding: utf-8 -*-
"""
Created on Tue May 19 10:35:04 2026

@author: Admin
"""

# -*- coding: utf-8 -*-

# =========================================================
# B ENTITY EXTRACTION MAPPING - STAGE 2 PREPARATION
# =========================================================
# מטרת הקוד:
# להוציא מתוך עמותות B ישויות רפואיות נקיות:
# 1. Disease aliases ברורים
# 2. Medical domains
# 3. Specialties / procedures
#
# הקוד לא משתמש עדיין ב-Cosine Similarity.
# הוא מכין input נקי לשלב הבא.
# =========================================================

import pandas as pd
import numpy as np
import re
from pathlib import Path

# =========================================================
# PATHS
# =========================================================

BASE_DIR = Path().resolve()

FINAL_FILE = BASE_DIR / "FINAL_DATASET_STAGE1.xlsx"
DALY_FILE = BASE_DIR / "DALYS_2022.xlsx"

OUTPUT_FILE = BASE_DIR / "B_entity_extraction_mapping_v1.xlsx"

# =========================================================
# HELPERS
# =========================================================

def normalize_text(text):
    if pd.isna(text):
        return ""

    text = str(text).lower().strip()
    text = text.replace("&", " and ")
    text = text.replace("/", " ")
    text = text.replace("-", " ")

    text = re.sub(r"[^a-z0-9\s']", " ", text)
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def build_word_boundary_pattern(term):
    term_norm = normalize_text(term)
    return rf"(?<![a-z0-9]){re.escape(term_norm)}(?![a-z0-9])"


def extract_terms(text, terms):
    text_norm = normalize_text(text)
    found = []

    for term in terms:
        pattern = build_word_boundary_pattern(term)
        if re.search(pattern, text_norm):
            found.append(term)

    return list(dict.fromkeys(found))


def join_terms(x):
    if len(x) == 0:
        return np.nan
    return "; ".join(x)


# =========================================================
# LOAD FILES
# =========================================================

df_final = pd.read_excel(
    FINAL_FILE,
    sheet_name="final_dataset"
)

df_daly = pd.read_excel(DALY_FILE)

df_daly["Disease_norm"] = df_daly["Disease"].apply(normalize_text)

daly_terms = set(df_daly["Disease_norm"].dropna())

print("Rows in FINAL:", len(df_final))
print("DALY rows:", len(df_daly))

# =========================================================
# FILTER B ONLY
# =========================================================

df_B = df_final[
    df_final["auto_category"].astype(str).str.strip() == "B"
].copy()

print("B rows:", len(df_B))

# =========================================================
# DISEASE ALIASES
# =========================================================
# אלה מונחים שאם מופיעים במשימה,
# ב-move_to_A_candidate.
# =========================================================

disease_aliases = {
    # neurological / rare
    "als": "Motor neuron disease",
    "amyotrophic lateral sclerosis": "Motor neuron disease",
    "lou gehrig": "Motor neuron disease",
    "lou gehrig disease": "Motor neuron disease",
    "lou gehrig's disease": "Motor neuron disease",

    "parkinson": "Parkinson disease",
    "parkinsons": "Parkinson disease",
    "parkinson disease": "Parkinson disease",
    "parkinson's disease": "Parkinson disease",

    "huntington": "Huntington disease",
    "huntingtons": "Huntington disease",
    "huntington disease": "Huntington disease",
    "huntington's disease": "Huntington disease",

    "epilepsy": "Epilepsy",
    "cerebral palsy": "Cerebral palsy",
    "brain injury": "Brain injury",
    "traumatic brain injury": "Brain injury",
    "spinal cord injury": "Spinal cord injury",

    # cancer
    "breast cancer": "Breast cancer",
    "leukemia": "Leukaemia",
    "leukaemia": "Leukaemia",
    "acute myeloid leukemia": "Leukaemia",
    "aml": "Leukaemia",
    "skin cancer": "Malignant skin melanoma",
    "melanoma": "Malignant skin melanoma",
    "rare cancers": "Other neoplasms",

    # respiratory / genetic / immune
    "cystic fibrosis": "Cystic fibrosis",
    "sickle cell": "Sickle cell disorders",
    "sickle cell disease": "Sickle cell disorders",
    "sickle cell anemia": "Sickle cell disorders",

    # skin / inflammatory / allergy
    "psoriasis": "Skin diseases",
    "food allergy": "Asthma",
    "food allergies": "Asthma",
    "anaphylaxis": "Asthma",
    "amyloidosis": "Other cardiovascular and circulatory diseases",
    "lymphedema": "Lymphedema",

    # mental / developmental
    "adhd": "Attention-deficit/hyperactivity disorder",
    "attention deficit hyperactivity disorder": "Attention-deficit/hyperactivity disorder",
    "autism": "Autism and Asperger syndrome",
    "asd": "Autism and Asperger syndrome",
    "down syndrome": "Down syndrome",
    "fragile x": "Fragile X syndrome",
    "dyslexia": "Dyslexia",

    # women / reproductive
    "endometriosis": "Gynecological diseases",
    "fibroid": "Gynecological diseases",
    "fibroids": "Gynecological diseases",

    # cardiovascular / kidney / diabetes
    "heart disease": "Other circulatory diseases",
    "cardiovascular disease": "Other circulatory diseases",
    "kidney disease": "Chronic kidney disease",
    "renal failure": "Chronic kidney disease",
    "dialysis": "Chronic kidney disease",
    "diabetes": "Diabetes mellitus",
    "diabetic": "Diabetes mellitus",

    # oral / dental
    "dental caries": "Dental caries",
    "oral disorder": "Other oral disorders",
    "oral disorders": "Other oral disorders",
}

disease_alias_terms = sorted(
    disease_aliases.keys(),
    key=len,
    reverse=True
)

# =========================================================
# MEDICAL DOMAIN TERMS
# =========================================================
# רפואי ברור, אבל לא בהכרח מחלה ספציפית.
# =========================================================

medical_domain_terms = [
    "dermatology",
    "cardiology",
    "neurology",
    "neuroscience",
    "ophthalmology",
    "optometry",
    "dentistry",
    "urology",
    "gynecology",
    "gastroenterology",
    "rheumatology",
    "orthopedic",
    "orthopaedic",
    "anesthesiology",
    "radiology",
    "pathology",
    "oncology",
    "hematology",
    "immunology",
    "allergy",
    "allergists",
    "immunologists",
    "pediatrics",
    "pediatric",
    "mental health",
    "behavioral health",
    "public health",
    "biomedical",
    "life science",
    "bioscience",
    "biotechnology",
]

# =========================================================
# PROCEDURE / SYSTEM / CARE TERMS
# =========================================================
# רפואי תשתיתי / מקצועי / טיפולי.
# =========================================================

procedure_terms = [
    "surgery",
    "surgeries",
    "surgical",
    "surgeon",
    "surgeons",
    "transplant",
    "transplantation",
    "organ donation",
    "organ procurement",
    "tissue",
    "medical devices",
    "pharmaceutical",
    "pharmaceuticals",
    "medicines",
    "rehabilitation",
    "rehab",
    "physical therapy",
    "physical therapists",
    "clinical research",
    "medical research",
    "patient care",
    "medical care",
    "healthcare",
    "health care",
    "diagnostics",
    "diagnostic",
    "vaccines",
    "vaccine",
    "treatment",
    "treatments",
    "therapy",
    "therapies",
]

# =========================================================
# EXTRACT ENTITIES
# =========================================================

df_B["extracted_disease_alias_terms"] = df_B["Org_mission"].apply(
    lambda x: extract_terms(x, disease_alias_terms)
)

df_B["extracted_disease_targets"] = df_B["extracted_disease_alias_terms"].apply(
    lambda terms: [disease_aliases[t] for t in terms]
)

df_B["extracted_disease_alias_terms"] = df_B["extracted_disease_alias_terms"].apply(join_terms)
df_B["extracted_disease_targets"] = df_B["extracted_disease_targets"].apply(join_terms)

df_B["extracted_medical_domains"] = df_B["Org_mission"].apply(
    lambda x: join_terms(extract_terms(x, medical_domain_terms))
)

df_B["extracted_procedure_terms"] = df_B["Org_mission"].apply(
    lambda x: join_terms(extract_terms(x, procedure_terms))
)

# =========================================================
# MATCH TYPE
# =========================================================

df_B["has_extracted_disease"] = df_B["extracted_disease_targets"].notna()
df_B["has_medical_domain"] = df_B["extracted_medical_domains"].notna()
df_B["has_procedure_terms"] = df_B["extracted_procedure_terms"].notna()

def classify_entity_action(row):

    if row["has_extracted_disease"]:
        return "move_to_A_candidate"

    if row["has_medical_domain"] or row["has_procedure_terms"]:
        return "keep_B_entity_proxy"

    return "keep_B_unmapped"


df_B["entity_stage_action"] = df_B.apply(
    classify_entity_action,
    axis=1
)

# =========================================================
# CLEAN ENTITY TEXT FOR FUTURE COSINE SIMILARITY
# =========================================================

def build_clean_entity_text(row):

    parts = []

    for col in [
        "extracted_disease_targets",
        "extracted_medical_domains",
        "extracted_procedure_terms"
    ]:
        val = row.get(col)

        if pd.notna(val) and str(val).strip() != "":
            parts.append(str(val))

    if len(parts) == 0:
        return np.nan

    return "; ".join(parts)


df_B["clean_entity_text_for_similarity"] = df_B.apply(
    build_clean_entity_text,
    axis=1
)

# =========================================================
# SUMMARY TABLES
# =========================================================

action_summary = (
    df_B["entity_stage_action"]
    .value_counts(dropna=False)
    .rename_axis("entity_stage_action")
    .reset_index(name="count")
)

domain_summary = (
    df_B["extracted_medical_domains"]
    .dropna()
    .str.split("; ")
    .explode()
    .value_counts()
    .rename_axis("medical_domain")
    .reset_index(name="count")
)

procedure_summary = (
    df_B["extracted_procedure_terms"]
    .dropna()
    .str.split("; ")
    .explode()
    .value_counts()
    .rename_axis("procedure_term")
    .reset_index(name="count")
)

disease_candidate_summary = (
    df_B["extracted_disease_targets"]
    .dropna()
    .str.split("; ")
    .explode()
    .value_counts()
    .rename_axis("disease_target")
    .reset_index(name="count")
)

# =========================================================
# SPLIT TABLES
# =========================================================

df_move_to_A = df_B[
    df_B["entity_stage_action"] == "move_to_A_candidate"
].copy()

df_B_proxy = df_B[
    df_B["entity_stage_action"] == "keep_B_entity_proxy"
].copy()

df_B_unmapped = df_B[
    df_B["entity_stage_action"] == "keep_B_unmapped"
].copy()

# =========================================================
# SAVE
# =========================================================

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:

    df_B.to_excel(
        writer,
        sheet_name="B_entity_matches",
        index=False
    )

    df_move_to_A.to_excel(
        writer,
        sheet_name="move_to_A_candidates",
        index=False
    )

    df_B_proxy.to_excel(
        writer,
        sheet_name="B_entity_proxy",
        index=False
    )

    df_B_unmapped.to_excel(
        writer,
        sheet_name="B_unmapped",
        index=False
    )

    action_summary.to_excel(
        writer,
        sheet_name="action_summary",
        index=False
    )

    disease_candidate_summary.to_excel(
        writer,
        sheet_name="disease_candidate_summary",
        index=False
    )

    domain_summary.to_excel(
        writer,
        sheet_name="domain_summary",
        index=False
    )

    procedure_summary.to_excel(
        writer,
        sheet_name="procedure_summary",
        index=False
    )

print("Saved:", OUTPUT_FILE)

print("\nAction summary:")
print(action_summary)

print("\nMove to A candidates:", len(df_move_to_A))
print("B entity proxy:", len(df_B_proxy))
print("B unmapped:", len(df_B_unmapped))