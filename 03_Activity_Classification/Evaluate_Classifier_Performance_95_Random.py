# -*- coding: utf-8 -*-
"""
Rule_Based_Activity_OrgType_Classifier_v1_2.py

Purpose:
Rule-based classifier for organizations that remained unmatched after the Disease Dictionary step.

Version 1.2 changes vs V1.1:
1. Stronger Non-health / Other rules.
2. Stronger Patient Support rules:
   - disability services
   - visual impairment / blindness
   - mental / behavioral health support
   - children and families support
3. Stronger Medical Research rules:
   - cure / finding cures
   - developing medicines
   - therapeutics
   - treatment options
   - research + cure / therapy / medicine combinations
4. Adds combo rules, not just single keyword rules.
5. Keeps the scope narrow:
   - Main Theme
   - Organization Type
   - Confidence
   - Validation against manual Excel sample

Required input files in the same folder:
1. Org_Mission_Disease_Dictionary_Matches_v2_enriched.xlsx
   Sheet: unmatched_orgs

2. Book3_classified_with_organization_type.xlsx
   Sheet: Classification_Template

3. 2022_Revenue.xlsx
   Sheet: Sheet1
   Optional but recommended. Used to add extract_name_2022 by extract_ein.

Output:
Rule_Based_Activity_OrgType_Classifier_v1_2_output.xlsx
"""

import pandas as pd
import numpy as np
import re
from pathlib import Path
from collections import defaultdict


# =========================================================
# 1. PATHS
# =========================================================

BASE_DIR = Path().resolve()

UNMATCHED_FILE = BASE_DIR / "Org_Mission_Disease_Dictionary_Matches_v2_enriched.xlsx"
UNMATCHED_SHEET = "unmatched_orgs"

MANUAL_FILE = BASE_DIR / "Random_classified_with_organization_type.xlsx"
MANUAL_SHEET = "Classification_Template"

REVENUE_2022_FILE = BASE_DIR / "2022_Revenue.xlsx"
REVENUE_2022_SHEET = "Sheet1"

OUTPUT_FILE = BASE_DIR / "Rule_Based_Activity_OrgType_Classifier__95_Random.xlsx"


# =========================================================
# 2. PARAMETERS
# =========================================================

MISSION_COL_CANDIDATES = ["Org_mission", "org_mission", "mission", "extract_org_mission"]
REVENUE_COL_CANDIDATES = ["Revenue", "990extract_revenue", "extract_revenue", "revenue", "total_revenue"]
EIN_COL = "extract_ein"

TEXT_COLS_FOR_CLASSIFICATION = [
    "Org_mission",
    "extract_name",
    "extract_name_2022",
    "bmf_NTEE_NCCS",
    "NTEE",
]

OTHER_MIN_SCORE = 7
HEALTH_BLOCK_OTHER_SCORE = 3

HIGH_MIN_SCORE = 8
HIGH_MIN_GAP = 4
MEDIUM_MIN_SCORE = 5
MEDIUM_MIN_GAP = 2


# =========================================================
# 3. HELPERS
# =========================================================






















































def normalize_text(x):
    if pd.isna(x):
        return ""
    x = str(x).lower().strip()
    x = x.replace("&", " and ")
    x = x.replace("/", " ")
    x = x.replace("-", " ")
    x = re.sub(r"[^a-z0-9\s]", " ", x)
    x = re.sub(r"\s+", " ", x)
    return x.strip()


def get_col(df, candidates, required=True):
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


def has_phrase(text, phrase):
    return re.search(r"\b" + re.escape(phrase) + r"\b", text) is not None


def phrase_hits(text, phrases):
    return [p for p in phrases if has_phrase(text, p)]


def has_any(text, phrases):
    return any(has_phrase(text, p) for p in phrases)


def join_text(row, cols):
    parts = []
    for c in cols:
        if c in row.index and pd.notna(row[c]):
            parts.append(str(row[c]))
    return " ".join(parts)


def add_score(scores, evidence, bucket, label, points, rule_name, matched_text):
    scores[bucket][label] += points
    evidence.append({
        "bucket": bucket,
        "label": label,
        "points": points,
        "rule_name": rule_name,
        "matched_text": matched_text,
    })


def top_two(score_dict):
    clean = [(k, v) for k, v in score_dict.items() if v > 0]
    if not clean:
        return "Needs Review", 0, "", 0
    clean = sorted(clean, key=lambda x: x[1], reverse=True)
    first_label, first_score = clean[0]
    if len(clean) > 1:
        second_label, second_score = clean[1]
    else:
        second_label, second_score = "", 0
    return first_label, first_score, second_label, second_score


def confidence_from_scores(top_score, second_score):
    gap = top_score - second_score
    if top_score >= HIGH_MIN_SCORE and gap >= HIGH_MIN_GAP:
        return "High"
    if top_score >= MEDIUM_MIN_SCORE and gap >= MEDIUM_MIN_GAP:
        return "Medium"
    if top_score > 0:
        return "Low"
    return "Manual Review"


def scores_to_str(score_dict):
    return "; ".join(
        [f"{k}:{v}" for k, v in sorted(score_dict.items(), key=lambda x: x[1], reverse=True)]
    )


def evidence_to_str(evidence):
    return "; ".join(
        [f'{e["rule_name"]}="{e["matched_text"]}" +{e["points"]}->{e["bucket"]}:{e["label"]}'
         for e in evidence]
    )


def clean_ein(x):
    if pd.isna(x):
        return ""
    s = str(x).strip()
    s = re.sub(r"\.0$", "", s)
    s = re.sub(r"[^0-9]", "", s)
    return s


# =========================================================
# 4. RULE KNOWLEDGE BASE
# =========================================================

KB = []

def rule(phrases, main=None, org_type=None, health=0, other=0, name=""):
    if isinstance(phrases, str):
        phrases = [phrases]
    KB.append({
        "phrases": phrases,
        "main": main or {},
        "org_type": org_type or {},
        "health": health,
        "other": other,
        "name": name,
    })


# =========================================================
# 4A. GENERAL HEALTH SIGNALS
# =========================================================

rule(
    ["health", "healthcare", "health care", "medical", "medicine", "clinical",
     "clinic", "hospital", "patient", "patients", "care", "treatment",
     "disease", "diseases", "biomedical", "public health", "physician",
     "physicians", "surgery", "pathology", "pediatric", "pediatrics",
     "therapy", "therapeutic", "therapeutics"],
    health=1,
    name="general_health_signal"
)


# =========================================================
# 4B. PATIENT SUPPORT RULES
# =========================================================

rule(
    ["patient assistance", "financial need", "uninsured", "insurance plan",
     "access to medicine", "access to medicines", "access to medication",
     "qualifying patients", "copay", "co pay", "medicine access",
     "medication access"],
    main={"Patient Support": 9},
    org_type={"Patient Assistance Foundation": 9},
    health=3,
    name="patient_assistance_access"
)

rule(
    ["patient support", "support patients", "support for patients",
     "individuals with", "people with", "persons with"],
    main={"Patient Support": 5},
    org_type={"Patient Support Organization": 5},
    health=2,
    name="patient_support_population"
)

rule(
    ["advocacy", "advocate", "awareness", "advance and improve the care",
     "improve the care of individuals"],
    main={"Patient Support": 5},
    org_type={"Patient Advocacy Organization": 6},
    health=2,
    name="patient_advocacy"
)

# V1.1: disability / impairment support
rule(
    ["developmental disabilities", "developmentally disabled",
     "intellectual disabilities", "special needs", "handicapped persons",
     "disabled persons", "disability services", "learning disabilities",
     "frail and elderly", "supported living", "residential habilitation"],
    main={"Patient Support": 8},
    org_type={"Patient Support Organization": 7},
    health=3,
    name="patient_support_disability_services"
)

rule(
    ["visually impaired", "vision impaired", "blind", "blindness",
     "hearing impaired", "deaf", "low vision"],
    main={"Patient Support": 8},
    org_type={"Patient Support Organization": 7},
    health=3,
    name="patient_support_impairment"
)

rule(
    ["mental health", "behavioral health", "emotional health",
     "children and families improve their mental health",
     "improve their mental health", "families improve their mental health"],
    main={"Patient Support": 7},
    org_type={"Patient Support Organization": 6},
    health=3,
    name="patient_support_mental_behavioral_health"
)

rule(
    ["children and families", "children and adults", "children with",
     "families with", "helping children", "help children and young people",
     "children reach their potential"],
    main={"Patient Support": 5},
    org_type={"Patient Support Organization": 5},
    health=1,
    name="patient_support_children_families"
)


# =========================================================
# 4C. HEALTHCARE SERVICES RULES
# =========================================================

rule(
    ["academic medical center", "university medical center"],
    main={"Healthcare Services": 10},
    org_type={"Academic Medical Center": 10},
    health=4,
    name="academic_medical_center"
)

rule(
    ["hospital", "medical center", "children s hospital", "children hospital"],
    main={"Healthcare Services": 9},
    org_type={"Hospital / Medical Center": 9},
    health=4,
    name="hospital_medical_center"
)

rule(
    ["physician group", "medical group", "orthopaedic specialty services",
     "orthopedic specialty services", "specialty services"],
    main={"Healthcare Services": 8},
    org_type={"Physician Group / Clinical Practice": 8},
    health=3,
    name="physician_group_clinical_practice"
)

rule(
    ["provide healthcare", "provide health care", "healthcare services",
     "health care services", "clinical care", "patient centered care",
     "primary care", "quality healthcare", "quality health care",
     "medical services", "provide medical services", "treatment of", "therapy", "therapies",
     "rehabilitation", "clinical services", "medical specialists", "medical specialty services",
     "subspecialists", "quality health care to children", "provide quality health care"],
    main={"Healthcare Services": 7},
    org_type={"Clinical Services Organization": 6},
    health=3,
    name="direct_healthcare_services"
)

rule(
    ["controls healthcare facilities", "controls health care facilities",
     "operates healthcare facilities", "operates health care facilities"],
    main={"Healthcare Services": 8},
    org_type={"Hospital / Medical Center": 6},
    health=3,
    name="controls_healthcare_facilities"
)

rule(
    ["community clinic", "community clinics", "community health center"],
    main={"Healthcare Services": 8},
    org_type={"Community Health Center / Clinic": 8},
    health=3,
    name="community_clinic"
)

rule(
    ["organ transplant", "organ transplantation", "organ donation",
     "organs and tissues", "transplants", "transplantation"],
    main={"Healthcare Services": 6},
    org_type={"Organ Procurement / Transplant Organization": 9},
    health=3,
    name="organ_procurement_transplant"
)

rule(
    ["hospital ships", "ships filled with volunteer professionals"],
    main={"Healthcare Services": 6},
    org_type={"Hospital Ship / Mobile Care": 9},
    health=3,
    name="hospital_ship_mobile_care"
)



rule(
    ["employs pediatric medicine specialists", "employs medical specialists",
     "specialists who practice", "subspecialists who practice",
     "provide quality health care to children", "quality health care to children in the community"],
    main={"Healthcare Services": 9},
    org_type={"Physician Group / Clinical Practice": 7},
    health=3,
    name="healthcare_specialists_practice_care"
)

# =========================================================
# 4D. MEDICAL RESEARCH RULES
# =========================================================

rule(
    ["medical research", "biomedical research", "clinical research",
     "scientific research", "laboratory research", "translational research",
     "research and education", "research institute", "research center",
     "research foundation", "clinical trials", "federally sponsored research",
     "privately sponsored research", "advance medical research",
     "advancing clinically relevant medical research", "conduct of research"],
    main={"Medical Research": 9},
    org_type={"Research Organization": 8},
    health=3,
    name="medical_research_strong"
)

rule(
    ["research support service", "archiving and distributing", "plasmids",
     "library of plasmids", "research support", "data generated from the materials",
     "repository", "biobank", "bio bank", "registry"],
    main={"Health Infrastructure": 7, "Medical Research": 3},
    org_type={"Health Technology / Infrastructure Organization": 7},
    health=3,
    name="research_infrastructure_support"
)

rule(
    ["innovative ideas", "new ways to treat illnesses", "medical innovators",
     "practical application of innovative ideas", "new clinical treatments"],
    main={"Medical Research": 6},
    org_type={"Research Organization": 5},
    health=2,
    name="medical_innovation_research"
)

# V1.1: cure / medicine / therapeutics research
rule(
    ["finding cures", "find cures", "finding a cure", "find a cure",
     "cure for", "toward a cure", "advance research toward a cure",
     "research toward a cure", "aggressive research agenda"],
    main={"Medical Research": 10},
    org_type={"Research Organization": 8},
    health=3,
    name="medical_research_cure"
)

rule(
    ["developing medicines", "develop medicines", "develop new medicines",
     "drug development", "medicine development", "new therapies",
     "new treatments", "treatment options", "therapeutics", "therapeutic development",
     "device development", "medical device development"],
    main={"Medical Research": 8},
    org_type={"Research Organization": 7},
    health=3,
    name="medical_research_therapeutics_development"
)

rule(
    ["funding research agenda", "funding an aggressive research agenda",
     "research agenda", "clinical treatment development"],
    main={"Medical Research": 8},
    org_type={"Research Organization": 7},
    health=3,
    name="medical_research_agenda"
)


rule(
    ["disease oriented science", "new era of discovery", "drive a new era of discovery",
     "translate innovative laboratory findings", "practical application of chemistry physics and biology"],
    main={"Medical Research": 10},
    org_type={"Research Organization": 8},
    health=3,
    name="medical_research_discovery_disease_science"
)

# Weak research indicators: should not dominate alone.
rule(
    ["research", "discovery", "laboratory", "scientific", "science", "institute"],
    main={"Medical Research": 2},
    org_type={"Research Organization": 2},
    health=0,
    name="research_weak_indicator"
)


# =========================================================
# 4E. PUBLIC HEALTH RULES
# =========================================================

rule(
    ["public health", "population health", "health department",
     "epidemiology", "epidemiological", "health surveillance",
     "disease prevention", "prevention", "global health",
     "low and middle income countries"],
    main={"Public Health": 9},
    org_type={"Public Health Organization": 8},
    health=3,
    name="public_health_population"
)

rule(
    ["worldwide", "international health", "global mission"],
    main={"Public Health": 3},
    org_type={"Global Health Organization": 4},
    health=1,
    name="global_health_weak"
)


# =========================================================
# 4F. EDUCATION & PROFESSIONAL DEVELOPMENT RULES
# =========================================================

rule(
    ["certify", "certification", "diplomates", "accreditation",
     "accreditation program", "standards", "standards of excellence",
     "quality assurance", "implement standards", "competency"],
    main={"Education & Professional Development": 9},
    org_type={"Standards / Certification Organization": 9},
    health=2,
    name="standards_certification_accreditation"
)

rule(
    ["professional development", "continuing education", "training",
     "educate healthcare professionals", "educate health care professionals",
     "healthcare careers", "health care careers", "scholarships to students",
     "needs based scholarships", "fellowship"],
    main={"Education & Professional Development": 7},
    org_type={"Education / Training Organization": 7},
    health=2,
    name="education_training"
)

# In V1 this was too strong when combined with research.
# Kept, but lower weight for mentor/mentoring.
rule(
    ["mentor", "mentoring"],
    main={"Education & Professional Development": 4},
    org_type={"Education / Training Organization": 4},
    health=1,
    name="mentoring_education_weak"
)

rule(
    ["promote scholarship", "leadership", "high standards of excellence",
     "pursue scientific knowledge", "advance the profession", "foster excellence",
     "college of", "society of", "association of", "board certified"],
    main={"Education & Professional Development": 7},
    org_type={"Professional Society / Association": 8},
    health=2,
    name="professional_society_association"
)


# =========================================================
# 4G. HEALTH INFRASTRUCTURE RULES
# =========================================================

rule(
    ["medical equipment", "laboratory equipment", "equipment", "facilities",
     "technology platform", "data platform", "health technology", "medical technology",
     "technology and innovation"],
    main={"Health Infrastructure": 7},
    org_type={"Health Technology / Infrastructure Organization": 7},
    health=2,
    name="health_infrastructure_technology"
)


# =========================================================
# 4H. HEALTH FUNDING & GRANTMAKING RULES
# =========================================================

rule(
    ["provide grants", "grantmaking", "grant making", "endowment fund",
     "health endowment", "funds research", "funding research",
     "awards grants", "grant support"],
    main={"Health Funding & Grantmaking": 9},
    org_type={"Grantmaking Foundation / Endowment": 9},
    health=2,
    name="grantmaking_endowment"
)

rule(
    ["scientific educational financial and project management support",
     "project management support during the conduct of", "support during the conduct of research"],
    main={"Medical Research": 7},
    org_type={"Research Organization": 6},
    health=2,
    name="research_project_management_support"
)


# =========================================================
# 4I. NON-HEALTH / OTHER RULES
# =========================================================

rule(
    ["housing", "affordable housing", "homeless", "homelessness",
     "food bank", "food pantry", "arts", "music", "religion", "church",
     "synagogue", "economic development", "workforce development",
     "job placement", "employment", "environmental conservation", "conservation",
     "animal", "animal welfare", "community development", "community foundation",
     "faith", "spirituality", "ministers to the community",
     "rural america", "manufacturing industry", "paint and coatings",
     "coatings industry", "copper", "worldwide copper", "build thriving communities",
     "thriving communities"],
    main={"Non-health / Other": 8},
    org_type={"Non-health Organization": 8},
    other=5,
    name="non_health_strong_signal"
)

rule(
    ["social services"],
    main={"Needs Review": 3},
    org_type={"Needs Review": 3},
    other=1,
    name="social_services_too_general"
)


# =========================================================
# 5. COMBO RULES
# =========================================================

def apply_combo_rules(text, scores, evidence):
    """
    Combo rules give extra score when two concepts appear together.
    This is still rule-based and explainable, but avoids relying on isolated keywords.
    """


    # Direct care + professional providers should strengthen Healthcare Services.
    direct_care_terms = ["provide medical services", "quality health care", "clinical services",
                         "medical specialists", "subspecialists", "therapy", "rehabilitation"]
    provider_terms = ["employs", "provide", "services", "practice", "care"]

    if has_any(text, direct_care_terms) and has_any(text, provider_terms):
        add_score(scores, evidence, "main", "Healthcare Services", 6, "combo_direct_care_services", "direct care + services/providers")
        add_score(scores, evidence, "org_type", "Clinical Services Organization", 4, "combo_direct_care_services", "direct care + services/providers")

    # Medical research should beat education when discovery/disease-science is central.
    research_core_terms = ["discovery", "disease oriented science", "laboratory findings",
                           "clinical research", "scientific research", "new treatments",
                           "new clinical treatments", "medical research"]
    education_secondary_terms = ["mentor", "mentoring", "education", "educational"]

    if has_any(text, research_core_terms) and has_any(text, education_secondary_terms):
        add_score(scores, evidence, "main", "Medical Research", 7, "combo_research_beats_secondary_education", "research core + education/mentor")
        add_score(scores, evidence, "org_type", "Research Organization", 5, "combo_research_beats_secondary_education", "research core + education/mentor")

    # Research + cure/therapy/development should beat advocacy/support in research-heavy organizations.
    research_terms = ["research", "clinical trials", "scientific", "laboratory", "discovery"]
    cure_terms = ["cure", "cures", "therapeutics", "therapy", "treatment options",
                  "developing medicines", "drug development", "new treatments"]

    if has_any(text, research_terms) and has_any(text, cure_terms):
        add_score(scores, evidence, "main", "Medical Research", 8, "combo_research_plus_cure_therapy", "research + cure/therapy")
        add_score(scores, evidence, "org_type", "Research Organization", 6, "combo_research_plus_cure_therapy", "research + cure/therapy")

    # Patient + disability/impairment means Patient Support, not generic Healthcare.
    disability_terms = ["developmental disabilities", "intellectual disabilities", "learning disabilities",
                        "special needs", "visually impaired", "blind", "hearing impaired",
                        "disabled", "handicapped"]
    patient_population_terms = ["individuals", "people", "persons", "children", "families", "adults"]

    if has_any(text, patient_population_terms) and has_any(text, disability_terms):
        add_score(scores, evidence, "main", "Patient Support", 7, "combo_population_plus_disability", "population + disability")
        add_score(scores, evidence, "org_type", "Patient Support Organization", 6, "combo_population_plus_disability", "population + disability")

    # Non-health mission with no health signal gets an additional push to Non-health.
    non_health_terms = ["community foundation", "homelessness", "job placement", "faith", "spirituality",
                        "manufacturing industry", "paint and coatings", "copper", "rural america",
                        "build thriving communities", "thriving communities"]
    if has_any(text, non_health_terms) and not has_any(text, ["health", "medical", "patient", "patients", "clinical", "hospital", "disease"]):
        add_score(scores, evidence, "main", "Non-health / Other", 6, "combo_non_health_no_health_signal", "non-health + no health signal")
        add_score(scores, evidence, "org_type", "Non-health Organization", 6, "combo_non_health_no_health_signal", "non-health + no health signal")

    return scores, evidence


# =========================================================
# 6. CLASSIFICATION FUNCTION
# =========================================================

def classify_record(row):
    text_raw = join_text(row, TEXT_COLS_FOR_CLASSIFICATION)
    text = normalize_text(text_raw)

    scores = {
        "main": defaultdict(int),
        "org_type": defaultdict(int),
    }

    health_signal_score = 0
    other_signal_score = 0
    evidence = []

    for item in KB:
        hits = phrase_hits(text, item["phrases"])
        if not hits:
            continue

        for hit in hits:
            health_signal_score += item.get("health", 0)
            other_signal_score += item.get("other", 0)

            for label, pts in item.get("main", {}).items():
                add_score(scores, evidence, "main", label, pts, item["name"], hit)

            for label, pts in item.get("org_type", {}).items():
                add_score(scores, evidence, "org_type", label, pts, item["name"], hit)

    scores, evidence = apply_combo_rules(text, scores, evidence)

    # Conservative Non-health / Other logic
    other_score = scores["main"].get("Non-health / Other", 0)
    if other_score >= OTHER_MIN_SCORE and health_signal_score < HEALTH_BLOCK_OTHER_SCORE:
        pred_main = "Non-health / Other"
        main_score = other_score
        main_second = ""
        main_second_score = 0
        pred_org_type = "Non-health Organization"
        org_score = scores["org_type"].get("Non-health Organization", other_score)
        confidence = "High" if other_score >= 10 else "Medium"

    else:
        # Do not let Other win unless conservative condition is met.
        scores["main"].pop("Non-health / Other", None)
        scores["org_type"].pop("Non-health Organization", None)

        pred_main, main_score, main_second, main_second_score = top_two(scores["main"])
        pred_org_type, org_score, org_second, org_second_score = top_two(scores["org_type"])

        confidence = confidence_from_scores(main_score, main_second_score)

        if pred_main == "Needs Review" or main_score == 0:
            pred_main = "Needs Review"
            pred_org_type = "Needs Review"
            confidence = "Manual Review"

        # If evidence is too weak, force review.
        # V1.1: this rule is less aggressive because combo rules now add stronger evidence.
        elif health_signal_score < HEALTH_BLOCK_OTHER_SCORE and pred_main != "Non-health / Other":
            # V1.2: if the rule evidence is strong enough, keep the classification.
            # This prevents strong activity evidence from being downgraded to Needs Review.
            if main_score < 7:
                pred_main = "Needs Review"
                pred_org_type = "Needs Review"
                confidence = "Manual Review"

    needs_review = (
        pred_main == "Needs Review"
        or pred_org_type == "Needs Review"
        or confidence in ["Low", "Manual Review"]
    )

    return pd.Series({
        "predicted_main_theme": pred_main,
        "predicted_organization_type": pred_org_type,
        "predicted_confidence": confidence,
        "needs_manual_review": "Yes" if needs_review else "No",
        "main_theme_score": main_score,
        "main_theme_second": main_second,
        "main_theme_second_score": main_second_score,
        "main_theme_score_gap": main_score - main_second_score,
        "health_signal_score": health_signal_score,
        "other_signal_score": other_signal_score,
        "main_theme_scores": scores_to_str(scores["main"]),
        "organization_type_scores": scores_to_str(scores["org_type"]),
        "matched_rule_evidence": evidence_to_str(evidence),
        "_classification_text_norm": text,
    })


# =========================================================
# 7. ENRICH WITH 2022 NAME
# =========================================================

def add_2022_name(df):
    if not REVENUE_2022_FILE.exists():
        print("WARNING: 2022_Revenue.xlsx not found. extract_name_2022 will be empty.")
        df["extract_name_2022"] = ""
        return df

    df_rev22 = pd.read_excel(REVENUE_2022_FILE, sheet_name=REVENUE_2022_SHEET)
    df_rev22.columns = [str(c).strip() for c in df_rev22.columns]

    if EIN_COL not in df_rev22.columns or "extract_name" not in df_rev22.columns:
        print("WARNING: Could not find extract_ein/extract_name in 2022_Revenue.xlsx.")
        df["extract_name_2022"] = ""
        return df

    df = df.copy()
    df["_ein_clean"] = df[EIN_COL].apply(clean_ein) if EIN_COL in df.columns else ""
    df_rev22["_ein_clean"] = df_rev22[EIN_COL].apply(clean_ein)

    name_lookup = (
        df_rev22[["_ein_clean", "extract_name"]]
        .dropna(subset=["_ein_clean"])
        .drop_duplicates(subset=["_ein_clean"])
        .rename(columns={"extract_name": "extract_name_2022"})
    )

    df = df.merge(name_lookup, on="_ein_clean", how="left")
    df["extract_name_2022"] = df["extract_name_2022"].fillna("")

    return df


# =========================================================
# 8. RUN CLASSIFIER
# =========================================================

def prepare_input(df):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    mission_col = get_col(df, MISSION_COL_CANDIDATES)
    revenue_col = get_col(df, REVENUE_COL_CANDIDATES, required=False)

    if revenue_col is None:
        df["_revenue_for_summary"] = 0
    else:
        df["_revenue_for_summary"] = safe_revenue(df[revenue_col])

    if EIN_COL in df.columns:
        df[EIN_COL] = df[EIN_COL].apply(clean_ein)

    if mission_col != "Org_mission":
        df["Org_mission"] = df[mission_col]

    return df


def classify_dataframe(df, label):
    df = prepare_input(df)
    df = add_2022_name(df)

    df["_mission_clean"] = df["Org_mission"].astype(str).str.strip()

    df_work = df[
        df["Org_mission"].notna()
        & (df["_mission_clean"] != "")
        & (df["_mission_clean"].str.lower() != "nan")
    ].copy()

    print(f"{label}: rows with non-empty mission:", len(df_work))

    predictions = df_work.apply(classify_record, axis=1)
    df_result = pd.concat(
        [df_work.reset_index(drop=True), predictions.reset_index(drop=True)],
        axis=1
    )

    df_result = df_result.sort_values("_revenue_for_summary", ascending=False).reset_index(drop=True)

    total_rev = df_result["_revenue_for_summary"].sum()
    df_result["revenue_rank"] = np.arange(1, len(df_result) + 1)
    df_result["cumulative_revenue"] = df_result["_revenue_for_summary"].cumsum()
    df_result["cumulative_revenue_percent"] = (
        df_result["cumulative_revenue"] / total_rev * 100
    ).round(2) if total_rev else 0

    return df_result


# =========================================================
# 9. VALIDATION METRICS
# =========================================================

def build_validation_metrics(df_pred):
    rows = []
    df = df_pred.copy()

    if "Main Theme" in df.columns:
        df["main_theme_exact_match"] = (
            df["Main Theme"].astype(str).str.strip()
            == df["predicted_main_theme"].astype(str).str.strip()
        )

    if "Organization Type" in df.columns:
        df["organization_type_exact_match"] = (
            df["Organization Type"].astype(str).str.strip()
            == df["predicted_organization_type"].astype(str).str.strip()
        )

    rev = safe_revenue(df["_revenue_for_summary"])
    total_rev = rev.sum()

    for col, label in [
        ("main_theme_exact_match", "main_theme"),
        ("organization_type_exact_match", "organization_type"),
    ]:
        if col in df.columns:
            org_acc = df[col].mean() * 100
            rows.append({
                "metric": f"{label}_exact_accuracy_org_weighted",
                "value": round(org_acc, 2)
            })

            if total_rev:
                rev_acc = rev[df[col] == True].sum() / total_rev * 100
                rows.append({
                    "metric": f"{label}_exact_accuracy_revenue_weighted",
                    "value": round(rev_acc, 2)
                })

    review_rate = (df["needs_manual_review"] == "Yes").mean() * 100
    rows.append({
        "metric": "manual_review_rate_org_weighted",
        "value": round(review_rate, 2)
    })

    if total_rev:
        review_rev_rate = rev[df["needs_manual_review"] == "Yes"].sum() / total_rev * 100
        rows.append({
            "metric": "manual_review_rate_revenue_weighted",
            "value": round(review_rev_rate, 2)
        })

    return df, pd.DataFrame(rows)


def summary_by(df, group_col):
    out = (
        df.groupby(group_col, dropna=False, as_index=False)
        .agg(
            org_count=("Org_mission", "count"),
            total_revenue=("_revenue_for_summary", "sum"),
            manual_review_count=("needs_manual_review", lambda x: (x == "Yes").sum()),
        )
        .sort_values("total_revenue", ascending=False)
    )
    total = out["total_revenue"].sum()
    out["revenue_share_percent"] = (out["total_revenue"] / total * 100).round(2) if total else 0
    return out


def confusion_table(df, manual_col, pred_col):
    if manual_col not in df.columns or pred_col not in df.columns:
        return pd.DataFrame()
    return pd.crosstab(
        df[manual_col].fillna(""),
        df[pred_col].fillna(""),
        rownames=[manual_col],
        colnames=[pred_col],
        dropna=False
    ).reset_index()


# =========================================================
# 10. MAIN
# =========================================================

print("Loading unmatched organizations...")
df_unmatched = pd.read_excel(UNMATCHED_FILE, sheet_name=UNMATCHED_SHEET)
print("Unmatched rows loaded:", len(df_unmatched))

print("Loading manual validation sample...")
df_manual = pd.read_excel(MANUAL_FILE, sheet_name=MANUAL_SHEET)
print("Manual sample rows loaded:", len(df_manual))

print("Classifying full unmatched organizations...")
df_classified = classify_dataframe(df_unmatched, "unmatched_orgs")

print("Classifying manual validation sample...")
df_manual_pred = classify_dataframe(df_manual, "manual_sample")

df_manual_comparison, validation_metrics = build_validation_metrics(df_manual_pred)

main_theme_summary = summary_by(df_classified, "predicted_main_theme")
org_type_summary = summary_by(df_classified, "predicted_organization_type")
confidence_summary = summary_by(df_classified, "predicted_confidence")

manual_main_theme_summary = summary_by(df_manual_comparison, "Main Theme") if "Main Theme" in df_manual_comparison.columns else pd.DataFrame()
manual_org_type_summary = summary_by(df_manual_comparison, "Organization Type") if "Organization Type" in df_manual_comparison.columns else pd.DataFrame()

main_confusion = confusion_table(df_manual_comparison, "Main Theme", "predicted_main_theme")
org_type_confusion = confusion_table(df_manual_comparison, "Organization Type", "predicted_organization_type")

kb_rows = []
for i, item in enumerate(KB, start=1):
    kb_rows.append({
        "rule_id": i,
        "phrases": "; ".join(item["phrases"]),
        "main_scores": "; ".join([f"{k}:{v}" for k, v in item.get("main", {}).items()]),
        "organization_type_scores": "; ".join([f"{k}:{v}" for k, v in item.get("org_type", {}).items()]),
        "health_score": item.get("health", 0),
        "other_score": item.get("other", 0),
        "rule_name": item.get("name", ""),
    })
df_rules = pd.DataFrame(kb_rows)

version_log = pd.DataFrame([
    {"version": "V1.2", "change": "Removed broad non-health trigger 'environment'", "reason": "Avoid false positives such as 'diverse environment' in research missions."},
    {"version": "V1.2", "change": "Strengthened Healthcare Services evidence", "reason": "Reduce false Needs Review and improve direct-care classification."},
    {"version": "V1.2", "change": "Strengthened Medical Research when discovery/disease-science is central", "reason": "Prevent secondary education/mentoring words from overriding research missions."},
    {"version": "V1.2", "change": "Added combo rules for direct care and research-vs-education", "reason": "Use phrase combinations rather than isolated keywords."},
    {"version": "V1.2", "change": "Relaxed forced Needs Review threshold from main_score < 8 to < 7", "reason": "Keep strong classifications even when health signal count is low."},
])

print("Saving output...")
drop_cols = ["_mission_clean", "_classification_text_norm"]

with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
    df_classified.drop(columns=drop_cols, errors="ignore").to_excel(writer, sheet_name="classification_results", index=False)
    df_classified.head(100).drop(columns=drop_cols, errors="ignore").to_excel(writer, sheet_name="pareto_top100", index=False)
    main_theme_summary.to_excel(writer, sheet_name="main_theme_summary", index=False)
    org_type_summary.to_excel(writer, sheet_name="organization_type_summary", index=False)
    confidence_summary.to_excel(writer, sheet_name="confidence_summary", index=False)
    df_manual_comparison.drop(columns=drop_cols, errors="ignore").to_excel(writer, sheet_name="validation_comparison", index=False)
    validation_metrics.to_excel(writer, sheet_name="validation_metrics", index=False)
    main_confusion.to_excel(writer, sheet_name="main_theme_confusion", index=False)
    org_type_confusion.to_excel(writer, sheet_name="org_type_confusion", index=False)
    manual_main_theme_summary.to_excel(writer, sheet_name="manual_main_theme_summary", index=False)
    manual_org_type_summary.to_excel(writer, sheet_name="manual_org_type_summary", index=False)
    df_rules.to_excel(writer, sheet_name="rules_reference", index=False)
    version_log.to_excel(writer, sheet_name="version_log", index=False)

print("DONE")
print("Output file:", OUTPUT_FILE)

print("\nValidation metrics:")
print(validation_metrics)

print("\nMain theme summary:")
print(main_theme_summary)

print("\nOrganization type summary:")
print(org_type_summary.head(20))
