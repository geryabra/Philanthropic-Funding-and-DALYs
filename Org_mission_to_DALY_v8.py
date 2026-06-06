# -*- coding: utf-8 -*-
"""
Created on Thu Apr 23 10:57:38 2026

@author: Admin
"""

# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
import re
from pathlib import Path
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.cluster import KMeans

# =========================================================
# PATHS
# =========================================================
BASE_DIR = Path().resolve()

REVENUE_FILE = BASE_DIR / "2022_Revenue.xlsx"
DALY_FILE = BASE_DIR / "DALYS_2022.xlsx"
OUTPUT_FILE = BASE_DIR / "FINAL_DATASET_STAGE1.xlsx"

# =========================================================
# LOAD SOURCE FILES
# =========================================================
df_rev = pd.read_excel(REVENUE_FILE).copy()
df_daly = pd.read_excel(DALY_FILE).copy()

# =========================================================
# RENAME COLUMNS
# =========================================================
revenue_rename_map = {
    "extract_revenue": "Revenue",
    "extract_org_mission": "Org_mission",
    "NTEE": "NTEE"
}

daly_rename_map = {
    "% total": "TotalPerc",
    "DALYs": "DALYs",
    "Disease": "Disease"
}

df_rev = df_rev.rename(columns=revenue_rename_map)
df_daly = df_daly.rename(columns=daly_rename_map)

required_rev_cols = ["Revenue", "Org_mission", "NTEE"]
required_daly_cols = ["Disease", "DALYs", "TotalPerc"]

missing_rev = [c for c in required_rev_cols if c not in df_rev.columns]
missing_daly = [c for c in required_daly_cols if c not in df_daly.columns]

if missing_rev:
    raise ValueError(f"Missing columns in 2020_Revenue.xlsx: {missing_rev}")

if missing_daly:
    raise ValueError(f"Missing columns in DALYS_2020.xlsx: {missing_daly}")

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
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def is_research_context(text):
    text = normalize_text(text)
    
    research_terms = [
        "research", "study", "scientific", "investigation",
        "clinical research", "laboratory", "innovation",
        "prevention", "therapeutics"
    ]
    
    return any(t in text for t in research_terms)

def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).lower()
    text = re.sub(r"[^a-z\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def unique_preserve_order(items):
    seen = set()
    result = []
    for item in items:
        if item not in seen:
            seen.add(item)
            result.append(item)
    return result

def text_present(x):
    return pd.notna(x) and str(x).strip() != ""

def build_word_boundary_pattern(term):
    return rf"(?<![a-z0-9]){re.escape(term)}(?![a-z0-9])"

# =========================================================
# NORMALIZE CORE TEXT
# =========================================================
df_rev["Org_mission_norm"] = df_rev["Org_mission"].apply(normalize_text)
df_rev["text_clean"] = df_rev["Org_mission"].apply(clean_text)

df_daly["Disease_norm"] = df_daly["Disease"].apply(normalize_text)

# =========================================================
# BUILD DALY DICTIONARY
# term -> official DALY disease
# =========================================================
term_to_disease = {}

for _, row in df_daly.iterrows():
    disease = str(row["Disease"]).strip()
    disease_norm = normalize_text(disease)

    if disease_norm:
        term_to_disease[disease_norm] = disease
################ שינוי V7

# =========================================================
# CONTEXT FILTERS BEFORE B / CLUSTERING
# מטרת השלב:
# 1. לא לתת למילים חלשות כמו support/community/education להכניס שורה ל-B
# 2. לשמור B רק כאשר יש הקשר רפואי אמיתי
# 3. לסמן B_signal בנפרד מ-B_clean לפני clustering
# =========================================================
def has_strong_medical_context(text):
    strong_terms = [
        "patient", "patients", "hospital", "clinical",
        "treatment", "therapy", "medical", "healthcare",
        "diagnosis", "disease", "disorder"
    ]
    text = normalize_text(text)
    return any(t in text for t in strong_terms)


strong_medical_context_terms = [
    "medical", "healthcare", "patient", "patients", "physician", "physicians",
    "clinical", "clinic", "hospital", "hospitals", "treatment", "treatments",
    "therapy", "therapies", "medicine", "public health", "health care",
    "ambulance", "rescue", "emergency medical", "nursing", "diagnosis",
    "diagnostic", "rehabilitation", "surgery", "surgeries", "surgeon",
    "eyecare", "eye care", "eye exams", "ophthalmic", "ophthalmology",
    "optometric", "optometry", "dental", "dentistry", "urology",
    "pediatric dentistry", "medical equipment", "medical supplies"
    "disease", "diseases",
    "infectious disease", "infectious diseases",
    "diagnostics", "diagnostic",
    "vaccine", "vaccines",
    "adjuvant", "adjuvants",
    "drug", "drugs",
    "detect", "detection",
    "prevent", "prevention",
    "treat", "treating",
        # added medical-domain terms
    "disease", "diseases",
    "disorder", "disorders",
    "rare disease", "rare diseases",

    "immunology", "immune", "immunity", "immunization",
    "immunotherapy",

    "therapy", "therapies", "therapeutic", "therapeutics",

    "psychiatry", "psychiatric", "psychology", "psychological",

    "neurology", "neurological", "neurologic",

    "biomedical",
    "cure", "curing",
    "healing",

    "doctor", "doctors",

    "cardiology", "cardio", "cardiac",
    "heart disease",

    "sickle cell",
    "parkinson", "parkinsons", "parkinson disease",
        
        # surgery / clinical
    "clinical",
    "clinic",
    "medical clinic",
    "surgical",
    "surgery",
    "surgeon",
    "surgeons",
    
    # rehab / mobility
    "rehabilitation",
    "rehab",
    "prosthetic",
    "prosthetics",
    "mobility assistance",
    "assistive devices",
    
    # transplant / tissue / pharma
    "organ",
    "organ donation",
    "transplant",
    "transplantation",
    "tissue",
    "pharmaceutical",
    "pharmaceuticals",
    "medicines",
    "orthopaedic",
    "orthopedic",
    "medical devices",
    "procurement",
    "specialty services",
        # transplant / organ donation
    "organ",
    "organs",
    "tissue",
    "tissues",
    "transplant",
    "transplants",
    "transplantation",
    "organ donation",
    "organ procurement",

    # developmental disability / habilitation
    "developmental disability",
    "developmental disabilities",
    "developmentally disabled",
    "habilitation",
    "community habilitation",
    "day program",
    "residential services",
    # visual impairment / blindness
    "visual impairment",
    "visually impaired",
    "vision impairment",
    "blind",
    "blindness",
    "low vision",
]

medium_medical_context_terms = [
    "health", "care", "prevention", "wellness", "screening",
    "caregivers", "registry", "diagnosed", "cure", "research",
    "financial assistance", "travel expenses"
]

weak_general_terms = [
    "education", "support", "community", "services", "families",
    "awareness", "learning", "wellbeing", "well being",
    "quality of life", "advocacy", "professionals", "resources"
]

social_non_medical_terms = [
    "music", "dance", "performance", "cultural interaction",
    "workforce", "training certification", "homeownership",
    "housing", "mentoring", "youth", "underprivileged",
    "poor and underprivileged", "love and support",
    "early childhood program", "creative expression",
    "student members", "audiences", "industry"
        # disability / special needs / illness broad medical support
    "disabilities",
    "disability",
    "disabled",
    "special needs",
    "illness",
    "illnesses",
    "injury",
    "injuries",
    "conditions",

    # therapy / intervention services
    "intervention",
    "interventions",
    "hyperbaric",
    "hyperbaric services",

    # ophthalmology profession
    "ophthalmologist",
    "ophthalmologists",
    "ophthalmology",
]

def extract_phrase_matches(text, terms):
    text_norm = normalize_text(text)
    matches = []

    for term in terms:
        term_norm = normalize_text(term)
        pattern = build_word_boundary_pattern(term_norm)
        if re.search(pattern, text_norm):
            matches.append(term)

    return unique_preserve_order(matches)

def has_any_phrase(text, terms):
    return len(extract_phrase_matches(text, terms)) > 0

def is_likely_social_non_medical(text):
    has_social = has_any_phrase(text, social_non_medical_terms)
    has_strong_medical = has_any_phrase(text, strong_medical_context_terms)

    if has_social and not has_strong_medical:
        return True

    return False

def classify_b_signal_type(row):
    has_specific = text_present(row.get("matched_specific_terms"))
    has_disease = bool(row.get("has_disease_signal"))

    if has_specific:
        return np.nan

    if has_disease:
        return "B_signal"

    return "B_clean"


# =========================================================
# MANUAL ALIASES
# הרחבתי יחסית כדי לצמצם זליגה מ-A ל-B
# =========================================================
manual_aliases = {
    "autism": "Autism and Asperger syndrome",
    "autism spectrum disorder": "Autism and Asperger syndrome",
    "autism spectrum disorders": "Autism and Asperger syndrome",
    "asperger": "Autism and Asperger syndrome",
    "asperger syndrome": "Autism and Asperger syndrome",
     # ALS
    "als": "Motor neuron disease",
    "lou gehrig": "Motor neuron disease",
    "lou gehrigs disease": "Motor neuron disease",
    # CRMO
    "crmo": "Other musculoskeletal disorders",
    "chronic recurrent multifocal osteomyelitis": "Other musculoskeletal disorders",
    # Mossy foot
    "mossy foot": "Lymphatic filariasis",
    # Canser
    "cancer": "Other neoplasms",
    "childhood cancer": "Other neoplasms",
    "childhood cancers": "Other neoplasms",
    "pediatric cancer": "Other neoplasms",
    "pediatric cancers": "Other neoplasms",
    "solid tumor": "Other neoplasms",
    "solid tumors": "Other neoplasms",
    "brain tumor": "Brain and central nervous system cancer",
    "brain tumors": "Brain and central nervous system cancer",
    "oral cancer": "Lip and oral cavity cancer",
    "head neck cancer": "Other pharynx cancer",
    "maxillofacial cancers": "Lip and oral cavity cancer",
    "breast cancer": "Breast cancer",
    "melanoma": "Malignant skin melanoma",
    "skin cancer": "Malignant skin melanoma",
    "dipg": "Brain and central nervous system cancer",
    "astrocytoma": "Brain and central nervous system cancer",
    "astrocytomas": "Brain and central nervous system cancer",

    "hiv": "HIV/AIDS",
    "aids": "HIV/AIDS",
    "hiv aids": "HIV/AIDS",
    "hivaids": "HIV/AIDS",

    "epilepsy": "Epilepsy",
    "alzheimer": "Alzheimer's disease and other dementias",
    "alzheimers": "Alzheimer's disease and other dementias",
    "alzheimers disease": "Alzheimer's disease and other dementias",
    "dementia": "Alzheimer's disease and other dementias",
    "dementias": "Alzheimer's disease and other dementias",
    "stroke": "Stroke",
    "aneurysm": "Stroke",
    "brain aneurysm": "Stroke",
    "neurological disorders": "Neurological disorders",
    "neurological impairments": "Neurological disorders",
    "United Cerebral Palsy": "Cerebral palsy",
    "ucp": "Cerebral palsy",
  

    "diabetes": "Diabetes mellitus",
    "heart failure": "Other circulatory diseases",
    "heart condition": "Other circulatory diseases",
    "cardiac": "Other circulatory diseases",
    "cardiovascular diseases": "Other circulatory diseases",
    "pulmonary diseases": "Chronic respiratory diseases",
    
    # Kidney
    "kidney patient": "Chronic kidney disease",
    "kidney patients": "Chronic kidney disease",
    "kidney disease": "Chronic kidney disease",
    "kidney diseases": "Chronic kidney disease",
    "kidney related disorders": "Chronic kidney disease",
    "renal failure": "Chronic kidney disease",
    "renal": "Chronic kidney disease",
    "dialysis": "Chronic kidney disease",
    "urinary tract diseases": "Chronic kidney disease",

    "lupus": "Other musculoskeletal disorders",
    "muscular dystrophy": "Musculoskeletal disorders",
    "limb girdle muscular dystrophy": "Musculoskeletal disorders",
    "lgmd2d": "Musculoskeletal disorders",
    "prader willi syndrome": "Other congenital birth defects",
    "wolff parkinson white syndrome": "Other circulatory diseases",
    "scoliosis": "Other musculoskeletal disorders",
    "blindness": "Blindness and vision loss",
    "vision impairment": "Blindness and vision loss",
    "blind": "Blindness and vision loss",
    "cerebral palsy": "Cerebral palsy",
    "substance abuse disorder": "Substance use disorders",
    "substance abuse": "Substance use disorders",
    "pain": "Low back pain",
    "chronic pain": "Low back pain",
    "trauma": "Transport injuries",
    
        # Additional disease-specific aliases
    "sickle cell": "Sickle cell disorders",
    "sickle cell disease": "Sickle cell disorders",
    "sickle cell disorder": "Sickle cell disorders",
    "sickle cell anemia": "Sickle cell disorders",

    "parkinson": "Parkinson disease",
    "parkinsons": "Parkinson disease",
    "parkinson disease": "Parkinson disease",
    "parkinson's disease": "Parkinson disease",

    "heart disease": "Other circulatory diseases",
    
        # Additional A-level disease aliases
    "huntington": "Huntington disease",
    "huntingtons": "Huntington disease",
    "huntington disease": "Huntington disease",
    "huntington's disease": "Huntington disease",

    "parkinson": "Parkinson disease",
    "parkinsons": "Parkinson disease",
    "parkinson disease": "Parkinson disease",
    "parkinson's disease": "Parkinson disease",

    "lymphedema": "Lymphedema",

    "heart disease": "Other circulatory diseases",
    
    # Review findings - missed A
    "amyotrophic lateral sclerosis": "Motor neuron disease",
    "als": "Motor neuron disease",
    "lou gehrig": "Motor neuron disease",
    "lou gehrig disease": "Motor neuron disease",
    "lou gehrig's disease": "Motor neuron disease",
    
    "acute myeloid leukemia": "Leukaemia",
    "aml": "Leukaemia",
    "leukemia": "Leukaemia",
    "leukaemia": "Leukaemia",
    
    "cystic fibrosis": "Cystic fibrosis",
    
    "psoriasis": "Skin diseases",
    
    "adhd": "Attention-deficit/hyperactivity disorder",
    "attention deficit hyperactivity disorder": "Attention-deficit/hyperactivity disorder",
    
    "amyloidosis": "Other cardiovascular and circulatory diseases",
    
    "nemaline myopathy": "Other neurological conditions",
    
    "food allergy": "Asthma",
    "food allergies": "Asthma",
    "anaphylaxis": "Asthma",
    # Dental / oral
    "dental caries": "Dental caries",
    
    "oral disorder": "Other oral disorders",
    "oral disorders": "Other oral disorders",
    
    "oral surgery": "Other oral disorders",
    "maxillofacial": "Other oral disorders",
    "oral and maxillofacial": "Other oral disorders",
    
    "dentistry": "Other oral disorders",
    "dental surgery": "Other oral disorders",
        

}

# נכניס alias רק אם היעד קיים ב-DALY;
# אם לא קיים בדיוק, נשמור את המחרוזת כיעד בכל זאת, כדי לא לאבד אינפורמציה.
for alias, target in manual_aliases.items():
    alias_norm = normalize_text(alias)
    target_norm = normalize_text(target)

    if target_norm in term_to_disease:
        term_to_disease[alias_norm] = term_to_disease[target_norm]
    else:
        term_to_disease[alias_norm] = target

# =========================================================
# EXPAND DICTIONARY WITH BASIC VARIANTS
# =========================================================
expanded_term_to_disease = {}

for term, disease in term_to_disease.items():
    variants = {
        term,
        term.replace(" diseases", " disease"),
        term.replace(" disorders", " disorder"),
        term.replace(" syndromes", " syndrome"),
        term.replace(" cancers", " cancer"),
        term.replace(" and ", " "),
    }

    for v in variants:
        v = normalize_text(v)
        if v:
            expanded_term_to_disease.setdefault(v, disease)

term_to_disease = expanded_term_to_disease

# =========================================================
# SPECIFIC DISEASE MATCHING
# =========================================================
all_specific_terms = sorted(term_to_disease.keys(), key=len, reverse=True)

def extract_specific_diseases(text):
    text_norm = normalize_text(text)
    matches = []

    for term in all_specific_terms:
        pattern = build_word_boundary_pattern(term)
        if re.search(pattern, text_norm):
            matches.append(term_to_disease[term])

    matches = unique_preserve_order(matches)
    return matches

# =========================================================
# GENERAL MEDICAL TERMS
# חשוב: הרשימה הזו כבר לא משמשת לבד כדי להכניס ל-B.
# היא נשמרת בעיקר כעמודת debug.
# =========================================================
general_medical_terms = (
    strong_medical_context_terms
    + medium_medical_context_terms
    + weak_general_terms
)

def extract_general_medical_terms(text):
    matches = []
    matches.extend(extract_phrase_matches(text, strong_medical_context_terms))
    matches.extend(extract_phrase_matches(text, medium_medical_context_terms))
    matches.extend(extract_phrase_matches(text, weak_general_terms))
    return unique_preserve_order(matches)


# =========================================================
# BUILD MATCH COLUMNS
# =========================================================
df_rev["matched_specific_terms"] = df_rev["Org_mission"].apply(extract_specific_diseases)
df_rev["matched_specific_terms"] = df_rev["matched_specific_terms"].apply(
    lambda x: "; ".join(x) if len(x) > 0 else np.nan
)

df_rev["matched_general_disease_terms"] = df_rev["Org_mission"].apply(extract_general_medical_terms)
df_rev["matched_general_disease_terms"] = df_rev["matched_general_disease_terms"].apply(
    lambda x: "; ".join(x) if len(x) > 0 else np.nan
)

# =========================================================
# HAS DISEASE SIGNAL
# משמש להבנת B_dirty
# =========================================================
disease_keywords = sorted(
    [
        "cancer", "diabetes", "autism", "hiv", "aids", "hivaids",
        "epilepsy", "stroke", "kidney", "cardiac", "heart failure",
        "tumor", "brain tumor", "lupus", "aneurysm", "als",
        "lou gehrig", "alzheimer", "dementia", "dialysis", "renal",
        "muscular dystrophy", "prader willi", "cerebral palsy",
        "blindness", "vision impairment", "scoliosis",
        "wolff parkinson white", "addiction", "substance abuse",
        "crmo", "osteomyelitis", "mossy foot"
        "disease", "diseases",
        "infectious disease", "infectious diseases",
        "diagnostics", "vaccines", "vaccine",
        "adjuvants", "adjuvant",
        "disease", "diseases",
        "disorder", "disorders",
        "rare disease", "rare diseases",
        "sickle cell",
        "sickle cell disease",
        "sickle cell disorder",
        "parkinson",
        "parkinsons",
        "parkinson disease",
        "parkinson's disease",
        "heart disease",
        "cardio",
        "cardiac",
        "neurolog",
        "psychiatr",
        "psycholog",
        "immun",
        "biomedical",
        "cure",
        "therapy",
        "therapeutic",
        "huntington",
        "huntingtons",
        "huntington disease",
        "huntington's disease",
        "parkinson",
        "parkinsons",
        "parkinson disease",
        "parkinson's disease",
        "lymphedema",
        "heart disease",
        "transplant",
        "transplants",
        "transplantation",
        "organ procurement",
        "developmental disability",
        "developmental disabilities",
        "developmentally disabled",
        "habilitation",
        "illness",
        "illnesses",
        "injury",
        "injuries",
        "disability",
        "disabilities",
        "disabled",
        "special needs",
        "hyperbaric",
        "ophthalmologist",
        "ophthalmologists",
        "blindness",
        "visual impairment",
        "visually impaired",
        
        
    ],
    key=len,
    reverse=True
)

def has_disease_signal(text):
    text_norm = normalize_text(text)
    for keyword in disease_keywords:
        pattern = build_word_boundary_pattern(normalize_text(keyword))
        if re.search(pattern, text_norm):
            return True
    return False

df_rev["has_disease_signal"] = df_rev["Org_mission"].apply(has_disease_signal)

# =========================================================
# CLASSIFICATION
# A = specific disease
# B = general medical / health-related but not specific
# C = non-medical / no signal
# =========================================================

### classify_row - עודכן בגרסה v8
def classify_row(row):
    mission_text = row.get("Org_mission", "")

    has_specific = text_present(row.get("matched_specific_terms"))
    has_disease = bool(row.get("has_disease_signal"))

    strong_matches = extract_phrase_matches(mission_text, strong_medical_context_terms)
    medium_matches = extract_phrase_matches(mission_text, medium_medical_context_terms)
    weak_matches = extract_phrase_matches(mission_text, weak_general_terms)

    has_strong_medical = len(strong_matches) > 0
    has_medium_medical = len(medium_matches) > 0
    has_weak_general = len(weak_matches) > 0

    if has_specific:
        return "A"

    if is_likely_social_non_medical(mission_text) and not has_disease:
        return "C"

    if has_disease:
        return "B"

    if has_strong_medical:
        return "B"

    if has_medium_medical and not has_weak_general:
        return "B"

    if has_medium_medical and has_weak_general:
        return "B"

    return "C"

df_rev["auto_category"] = df_rev.apply(classify_row, axis=1)

df_rev["medical_strong_terms"] = df_rev["Org_mission"].apply(
    lambda x: "; ".join(extract_phrase_matches(x, strong_medical_context_terms)) or np.nan
)

df_rev["medical_medium_terms"] = df_rev["Org_mission"].apply(
    lambda x: "; ".join(extract_phrase_matches(x, medium_medical_context_terms)) or np.nan
)

df_rev["weak_general_terms_found"] = df_rev["Org_mission"].apply(
    lambda x: "; ".join(extract_phrase_matches(x, weak_general_terms)) or np.nan
)

df_rev["is_social_non_medical"] = df_rev["Org_mission"].apply(is_likely_social_non_medical)

df_rev["B_signal_type"] = df_rev.apply(
    lambda row: classify_b_signal_type(row) if row["auto_category"] == "B" else np.nan,
    axis=1
)
# =========================================================
# PRIMARY DISEASE FOR A
# =========================================================
def get_primary_disease(x):
    if pd.isna(x) or str(x).strip() == "":
        return np.nan
    return str(x).split(";")[0].strip()

df_rev["Disease"] = df_rev["matched_specific_terms"].apply(get_primary_disease)
df_rev["Disease_norm"] = df_rev["Disease"].apply(normalize_text)




# =========================================================

# MERGE DALY DATA
# =========================================================
disease_merge_aliases = {
    "autism and asperger syndrome": "austism and asperger syndrome",
    "alzheimer s disease and other dementias": "alzheimer disease and other dementias",
    "blindness and vision loss": "other vision loss",
    "musculoskeletal disorders": "other musculoskeletal disorders",
    "cardiovascular diseases": "Other circulatory diseases",
    "neurological disorders": "neurological conditions",
    "transport injuries": "road injury",
    "cerebral palsy": "other neurological conditions",
    "chronic kidney disease":  "other chronic kidney disease",
    "low back pain": "back and neck pain",
    "brain and central nervous system cancer": "brain and nervous system cancers",
    "motor neuron disease": "other neurological conditions",
    "other congenital birth defects": "other congenital anomalies",
    "substance use disorders": "mental and substance use disorders",
    "lip and oral cavity cancer": "lip and oral cavity",
    "stroke": "stroke",
    "chronic respiratory diseases": "other respiratory diseases",
    "lymphedema": "other cardiovascular and circulatory diseases",
    "sickle cell disorders": "sickle cell disorders and trait",
    "attention deficit hyperactivity disorder":
    "attention deficit hyperactivity syndrome",
    "attention-deficit hyperactivity disorder":
    "attention deficit hyperactivity syndrome",
    "attention-deficit/hyperactivity disorder":"attention deficit hyperactivity syndrome",
    "other cardiovascular and circulatory diseases":"other circulatory diseases",
    "huntington disease": "other neurological conditions",
}


    
df_rev["Disease_merge_norm"] = df_rev["Disease_norm"].replace(disease_merge_aliases)

daly_disease = df_daly[["Disease_norm", "DALYs", "TotalPerc"]].copy()

daly_level2 = (
    df_daly.groupby("Level_2", dropna=False)
    .agg(
        DALYs=("DALYs", "sum"),
        TotalPerc=("TotalPerc", "sum")
    )
    .reset_index()
)

daly_level2["Disease_norm"] = daly_level2["Level_2"].apply(normalize_text)
daly_level2 = daly_level2[["Disease_norm", "DALYs", "TotalPerc"]]

# ---- Level_3 aggregation fallback ----
# Used for generic categories such as Stroke,
# where Revenue file contains subtypes but no generic "Stroke" disease row.
daly_level3 = (
    df_daly.groupby("Level_3", dropna=False)
    .agg(
        DALYs=("DALYs", "sum"),
        TotalPerc=("TotalPerc", "sum")
    )
    .reset_index()
)

daly_level3["Disease_norm"] = daly_level3["Level_3"].apply(normalize_text)
daly_level3 = daly_level3[["Disease_norm", "DALYs", "TotalPerc"]]

daly_lookup = (
    pd.concat([daly_disease, daly_level2, daly_level3], ignore_index=True)
    .drop_duplicates(subset=["Disease_norm"], keep="first")
)
 
df_final = df_rev.merge(
    daly_lookup,
    left_on="Disease_merge_norm",
    right_on="Disease_norm",
    how="left",
    suffixes=("", "_daly")
)


# =========================================================
# NLP CLUSTERING ON B_CLEAN ONLY
# B_clean = B with no disease signal
# =========================================================
df_B = df_final[df_final["auto_category"] == "B"].copy()
df_B_clean = df_B[df_B["B_signal_type"] == "B_clean"].copy()
df_B_signal = df_B[df_B["B_signal_type"] == "B_signal"].copy()

# ---- Post-B refinement before semantic matching ----
# Research כאן הוא מחקר רפואי כללי בלבד; מחקר ספציפי למחלה אמור כבר להיכנס ל-A.
def is_patient_support(text):
    text = normalize_text(text)

    patient_terms = [
        "patient", "patients", "care", "treatment", "diagnosis",
        "support", "therapy", "access", "services", "surgery", "surgeries"
    ]

    professional_noise_terms = [
        "members", "represent our members", "practice of", "profession",
        "association", "society"
    ]

    has_patient_signal = any(t in text for t in patient_terms)
    has_professional_noise = any(t in text for t in professional_noise_terms)

    return has_patient_signal and not has_professional_noise

def has_daly_hint(text):
    text = normalize_text(text)

    strong_daly_terms = [
        "disease", "diseases",
        "disorder", "disorders",
        "condition", "conditions",
        "diagnosis", "diagnostic",
        "syndrome", "syndromes",
        "rare disease", "rare diseases",
        "neurological",
        "behavioral health", "mental health",
        "medical care",
        "treatment", "therapy",
        "surgery", "surgeries",
        "patients", "patient"
    ]

    professional_noise_terms = [
        "members", "represent our members", "surgeons",
        "physicians", "practice of", "profession",
        "association", "society"
    ]

    has_strong_hint = any(t in text for t in strong_daly_terms)
    has_professional_noise = any(t in text for t in professional_noise_terms)

    return has_strong_hint and not has_professional_noise


df_B_clean["is_research"] = df_B_clean["Org_mission"].apply(is_research_context)
df_B_clean["is_patient_support"] = df_B_clean["Org_mission"].apply(is_patient_support)
df_B_clean["has_daly_hint"] = df_B_clean["Org_mission"].apply(has_daly_hint)

df_B_clean["B_final_type"] = np.select(
    [
        df_B_clean["is_research"] == True,
        df_B_clean["is_patient_support"] == True
    ],
    [
        "research",
        "patient_support"
    ],
    default="system_or_org"
)

df_final["Cluster"] = np.nan

# Initialize refinement columns with compatible dtypes
df_final["is_research"] = pd.Series(pd.NA, index=df_final.index, dtype="boolean")
df_final["is_patient_support"] = pd.Series(pd.NA, index=df_final.index, dtype="boolean")
df_final["has_daly_hint"] = pd.Series(pd.NA, index=df_final.index, dtype="boolean")
df_final["B_final_type"] = pd.Series(pd.NA, index=df_final.index, dtype="object")

if len(df_B_clean) >= 4:
    vectorizer = TfidfVectorizer(
        max_features=1000,
        stop_words="english",
        min_df=3,
        max_df=0.7
    )
    X = vectorizer.fit_transform(df_B_clean["text_clean"])

    kmeans = KMeans(n_clusters=4, random_state=42, n_init=10)
    df_B_clean["Cluster"] = kmeans.fit_predict(X)

    df_final.loc[df_B_clean.index, "Cluster"] = df_B_clean["Cluster"]

elif len(df_B_clean) > 0:
    df_B_clean["Cluster"] = 0
    df_final.loc[df_B_clean.index, "Cluster"] = df_B_clean["Cluster"]

# להחזיר את עמודות ה-refinement ל-df_final אחרי clustering
refine_cols = ["is_research", "is_patient_support", "has_daly_hint", "B_final_type"]
for col in refine_cols:
    df_final.loc[df_B_clean.index, col] = df_B_clean[col]

# =========================================================
# MINIMAL DEBUG COLUMNS
# =========================================================
debug_cols = [
    "matched_specific_terms",
    "matched_general_disease_terms",
    "has_disease_signal",
    "B_signal_type",
    "medical_strong_terms",
    "medical_medium_terms",
    "weak_general_terms_found",
    "is_social_non_medical",
    "is_research",
    "is_patient_support",
    "has_daly_hint",
    "B_final_type"
]

# =========================================================
# FINAL DATASET
# =========================================================
core_cols = [
    "Revenue",
    "Org_mission",
    "NTEE",
    "Disease",
    "DALYs",
    "TotalPerc",
    "auto_category",
    "B_signal_type",
    "B_final_type",
    "Cluster"
]

final_cols = core_cols + debug_cols
# Remove duplicate column names while preserving order.
# Without this, df_output["B_signal_type"] may become a DataFrame, causing value_counts() errors.
final_cols = list(dict.fromkeys([c for c in final_cols if c in df_final.columns]))

df_output = df_final[final_cols].copy()

# =========================================================
# OPTIONAL SUMMARY SHEETS
# =========================================================
category_summary = (
    df_output["auto_category"]
    .value_counts(dropna=False)
    .rename_axis("auto_category")
    .reset_index(name="count")
)

cluster_summary = (
    df_output["Cluster"]
    .value_counts(dropna=False)
    .rename_axis("Cluster")
    .reset_index(name="count")
)

# =========================================================
# SAVE
# =========================================================
with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df_output.to_excel(writer, sheet_name="final_dataset", index=False)
    df_final.to_excel(writer, sheet_name="full_enriched_debug", index=False)
    category_summary.to_excel(writer, sheet_name="category_summary", index=False)
    cluster_summary.to_excel(writer, sheet_name="cluster_summary", index=False)

print("Saved:", OUTPUT_FILE)
print("\nCategory distribution:")
print(df_output["auto_category"].value_counts(dropna=False))
print("\nCluster distribution:")
print(df_output["Cluster"].value_counts(dropna=False))
print("\nB signal split:")
print(df_output["B_signal_type"].value_counts(dropna=False))
print("\nPreview:")
print(df_output.head())



# =========================================================
# SEMANTIC DALY PROXIMITY FOR B_CLEAN
# Models:
# 1. all-mpnet-base-v2
# 2. all-MiniLM-L6-v2
# =========================================================

# אם חסר:
# pip install sentence-transformers

from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity
import pandas as pd
import numpy as np

# =========================
# SETTINGS
# =========================
TOP_N = 3

models_to_run = {
    "mpnet": "sentence-transformers/all-mpnet-base-v2",
    "minilm": "sentence-transformers/all-MiniLM-L6-v2"
}

# =========================
# PREPARE TEXTS
# =========================
# Semantic matching ירוץ רק על patient_support עם רמז DALY חזק.
df_for_matching = df_B_clean[
    (df_B_clean["B_final_type"] == "patient_support") &
    (df_B_clean["has_daly_hint"] == True)
].copy()

print("\nRows sent to semantic matching:")
print(len(df_for_matching))

print("\nB_final_type distribution:")
print(df_B_clean["B_final_type"].value_counts(dropna=False))

print("\nDALY hint distribution inside patient_support:")
print(
    df_B_clean[df_B_clean["B_final_type"] == "patient_support"]["has_daly_hint"]
    .value_counts(dropna=False)
)

print("\nSample rows sent to semantic:")
print(df_for_matching["Org_mission"].head(5))

b_texts = df_for_matching["Org_mission"].fillna("").astype(str).tolist()
daly_labels = df_daly["Disease"].fillna("").astype(str).tolist()

# אם יש לך גם Level_1/2/3 ורוצה להעשיר את טקסט ה-DALY:
# daly_labels = (
#     df_daly["Disease"].fillna("").astype(str)
#     + " "
#     + df_daly.get("Level_1", "").fillna("").astype(str)
#     + " "
#     + df_daly.get("Level_2", "").fillna("").astype(str)
# ).tolist()

# =========================
# FUNCTION: RUN ONE MODEL
# =========================
def run_semantic_matching(model_name, model_path, b_texts, daly_labels, top_n=3):
    print(f"\nRunning model: {model_name} | {model_path}")

    model = SentenceTransformer(model_path)

    b_embeddings = model.encode(
        b_texts,
        convert_to_numpy=True,
        show_progress_bar=True
    )

    daly_embeddings = model.encode(
        daly_labels,
        convert_to_numpy=True,
        show_progress_bar=True
    )

    sim_matrix = cosine_similarity(b_embeddings, daly_embeddings)

    results = []

    for i in range(sim_matrix.shape[0]):
        top_idx = np.argsort(sim_matrix[i])[::-1][:top_n]

        row_result = {}

        for rank, idx in enumerate(top_idx, start=1):
            row_result[f"{model_name}_top{rank}_daly"] = daly_labels[idx]
            row_result[f"{model_name}_top{rank}_score"] = round(float(sim_matrix[i][idx]), 4)

        results.append(row_result)

    return pd.DataFrame(results)

# =========================
# RUN BOTH MODELS
# =========================
results_list = []

if len(df_for_matching) == 0:
    print("\nNo rows were sent to semantic matching after filtering.")
    semantic_results = pd.DataFrame()
else:
    for model_name, model_path in models_to_run.items():
        res = run_semantic_matching(
            model_name=model_name,
            model_path=model_path,
            b_texts=b_texts,
            daly_labels=daly_labels,
            top_n=TOP_N
        )
        results_list.append(res)

    semantic_results = pd.concat(results_list, axis=1)

# =========================
# ATTACH RESULTS TO FILTERED B_CLEAN
# =========================
if len(df_for_matching) == 0:
    df_B_semantic = df_for_matching.copy()
    semantic_summary = pd.DataFrame(columns=["semantic_match_status", "count"])
    agreement_summary = pd.DataFrame(columns=["model_agreement_top1", "count"])
else:
    df_B_semantic = pd.concat(
        [
            df_for_matching.reset_index(drop=True),
            semantic_results.reset_index(drop=True)
        ],
        axis=1
    )

    # =========================
    # MODEL AGREEMENT CHECK
    # =========================
    df_B_semantic["model_agreement_top1"] = np.where(
        df_B_semantic["mpnet_top1_daly"] == df_B_semantic["minilm_top1_daly"],
        "agree",
        "disagree"
    )

    df_B_semantic["score_gap_mpnet"] = (
        df_B_semantic["mpnet_top1_score"] - df_B_semantic["mpnet_top2_score"]
    ).round(4)

    df_B_semantic["score_gap_minilm"] = (
        df_B_semantic["minilm_top1_score"] - df_B_semantic["minilm_top2_score"]
    ).round(4)

    # =========================
    # REVIEW FLAG
    # =========================
    def semantic_review_flag(row):
        if row["model_agreement_top1"] == "agree":
            if row["mpnet_top1_score"] >= 0.45 and row["minilm_top1_score"] >= 0.45:
                return "strong_semantic_match"
            return "agreement_but_low_score"

        if row["mpnet_top1_score"] < 0.40 and row["minilm_top1_score"] < 0.40:
            return "weak_match"

        return "needs_review"

    df_B_semantic["semantic_match_status"] = df_B_semantic.apply(
        semantic_review_flag,
        axis=1
    )

# =========================
# SUMMARY
# =========================
semantic_summary = (
    df_B_semantic["semantic_match_status"]
    .value_counts(dropna=False)
    .rename_axis("semantic_match_status")
    .reset_index(name="count")
)

agreement_summary = (
    df_B_semantic["model_agreement_top1"]
    .value_counts(dropna=False)
    .rename_axis("model_agreement_top1")
    .reset_index(name="count")
)

print("\nSemantic match status:")
print(semantic_summary)

print("\nModel agreement:")
print(agreement_summary)

# =========================
# SAVE
# =========================
output_semantic_file = "B_clean_semantic_DALY_similarity_v8.xlsx"

with pd.ExcelWriter(output_semantic_file, engine="openpyxl") as writer:
    df_B_semantic.to_excel(writer, sheet_name="B_clean_semantic_results", index=False)
    semantic_summary.to_excel(writer, sheet_name="semantic_summary", index=False)
    agreement_summary.to_excel(writer, sheet_name="agreement_summary", index=False)

print("\nSaved:", output_semantic_file)


df_B_clean[df_B_clean["is_research"] == True]["Org_mission"].head(10)

######
for c in sorted(df_B_clean["Cluster"].dropna().unique()):
    print(f"\n=== Cluster {c} ===")
    print(df_B_clean[df_B_clean["Cluster"] == c]["Org_mission"].head(15).to_string())
    
#############
print("\nMissing DALY values:")
print(df_final["DALYs"].isna().sum())
missing = df_final[df_final["DALYs"].isna()]["Disease"].value_counts()
print(missing.head(15))

missing = df_final[
    (df_final["auto_category"] == "A") &
    (df_final["DALYs"].isna())
]["Disease"].value_counts()

print(missing)

### 
check_terms = (
    "disease|disorder|immun|sickle cell|therap|psychiatr|"
    "parkinson|neurolog|biomedical|cure|cancer|heart disease|"
    "healing|doctor|cardio|psycholog"
)

print(
    df_output[
        df_output["Org_mission"].str.contains(check_terms, case=False, na=False)
    ][["Org_mission", "auto_category", "B_signal_type", "matched_specific_terms"]]
    .to_string()
)

print("\nMissing DALY values:")
print(df_final[df_final["auto_category"] == "A"]["DALYs"].isna().sum())

print(
    df_output[
        df_output["Org_mission"].str.contains(
            "amyotrophic lateral sclerosis|als|leukemia|aml|cystic fibrosis|psoriasis|adhd|amyloidosis|nemaline myopathy|food allerg",
            case=False,
            na=False
        )
    ][["Org_mission", "auto_category", "Disease", "DALYs", "TotalPerc"]]
    .to_string()
)