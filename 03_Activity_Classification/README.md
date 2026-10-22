# Activity Classification and Validation

## 1. Research Objective

This stage focuses on the development and evaluation of a two-dimensional rule-based classifier for nonprofit organizations that could not be matched to specific diseases using the DALY–MeSH dictionary.

The classifier assigns two complementary dimensions:

- **Main Theme:** The organization's primary area of healthcare-related activity.
- **Organization Type:** The organization's functional or operational role.

The objective is to characterize previously unmatched nonprofit organizations, improve the interpretation of healthcare-related philanthropic funding, and evaluate the reliability of automated classification.

## 2. Research Methodology

The classification process consists of two separate stages.

### Stage A — Classifier Development (Pareto Sample)

A revenue-based Pareto sample of 110 organizations was used for manual classification and development of the rule-based classifier.

The development process included:

- Manual classification of nonprofit organizations.
- Definition and refinement of classification categories.
- Development of weighted keyword and phrase-matching rules.
- Improvement of classification decisions and confidence thresholds.
- Identification of ambiguous organizations requiring manual review.

### Stage B — Independent Random Validation

An independent random sample of 95 organizations was selected from the eligible unmatched population.

Organizations previously included in the development sample were excluded.

Sampling was performed without replacement using a fixed random seed (`random_state=42`) to support reproducibility.

The selected organizations were manually classified and compared against the automated predictions.

This separation distinguishes classifier development from independent performance evaluation.

## 3. Where to Find the Classification Categories

The predefined classification categories are documented in the Excel reference worksheets.

**Pareto development workbook:**

`Book3_classified_with_organization_type.xlsx`

- `Lists` — Reference lists for Main Theme and Organization Type.
- `Field_Definitions` — Definitions of classification fields.
- `Sheet2` — Additional worksheet in the reviewed Pareto workbook.

**Independent random validation workbook:**

`Random_classified_with_organization_type.xlsx`

- `Lists` — Reference lists for Main Theme and Organization Type.
- `Field_Definitions` — Definitions of classification fields.

These reference worksheets provide the classification framework used for manual review.

The `Lists` worksheets in both workbooks contain the same predefined category lists.

## 4. Where to Find the Manual Classifications

### Pareto Development Sample

**File:** `Book3_classified_with_organization_type.xlsx`

**Sheet:** `Classification_Template`

Contains the manually classified development sample of 110 organizations.

This dataset was used to develop and refine the classification rules.

### Independent Random Validation Sample

**File:** `Random_classified_with_organization_type.xlsx`

**Sheet:** `Classification_Template`

Contains the manually assigned classifications for the independent random validation sample.

This dataset is used to evaluate classifier performance on organizations excluded from the development sample.

## 5. Where to Find the Automated Classification Results

**File:** `Rule_Based_Activity_OrgType_Classifier__95_Random.xlsx`

**Sheet:** `classification_results`

Contains the automated classifications for the unmatched organization population with non-empty mission statements.

Key columns include:

- `predicted_main_theme`
- `predicted_organization_type`
- `predicted_confidence`
- `needs_manual_review`
- `main_theme_score`
- `matched_rule_evidence`

The classifier uses weighted rules and combinations of phrases to assign classifications.

## 6. Where to Find the Validation Results

**File:** `Rule_Based_Activity_OrgType_Classifier__95_Random.xlsx`

| Worksheet | Description |
|---|---|
| `validation_comparison` | Manual classifications compared with automated predictions |
| `validation_metrics` | Exact-match accuracy and manual review rates |
| `main_theme_confusion` | Main Theme confusion matrix |
| `org_type_confusion` | Organization Type confusion matrix |
| `manual_main_theme_summary` | Summary of manually assigned Main Theme categories |
| `manual_org_type_summary` | Summary of manually assigned Organization Type categories |

Validation metrics include organization-weighted and revenue-weighted exact-match rates.

**Methodological consideration:** Exact-match accuracy calculated across the entire validation sample includes predictions marked as `Needs Review`. These results should be distinguished from accuracy calculated only for organizations receiving a definitive automated classification.

## 7. Classification Rules and Documentation

**File:** `Rule_Based_Activity_OrgType_Classifier__95_Random.xlsx`

- `rules_reference` — Classification keywords, phrases, category assignments, and scoring weights.
- `version_log` — Documented changes to the classifier.

**Python scripts:**

- `Rule_Based_Activity_OrgType_Classifier_v1_2.py` — Rule-based classifier development and application.
- `Create_Independent_Random_Validation_Sample_95.py` — Independent random validation sample generation.
- `Evaluate_Classifier_Performance_95_Random.py` — Classification and performance evaluation using the independent validation sample.

## 8. Input and Output Files

| File | Role |
|---|---|
| `Org_Mission_Disease_Dictionary_Matches_v2_enriched.xlsx` | Input population of organizations unmatched by the disease dictionary |
| `Book3_classified_with_organization_type.xlsx` | Manually classified Pareto development sample |
| `Independent_Random_Validation_Sample_95.xlsx` | Generated independent random sample |
| `Random_classified_with_organization_type.xlsx` | Manually classified independent validation sample |
| `2022_Revenue.xlsx` | Optional organization name enrichment |
| `Rule_Based_Activity_OrgType_Classifier_v1_2_output.xlsx` | Classifier development output |
| `Rule_Based_Activity_OrgType_Classifier__95_Random.xlsx` | Automated classification and independent validation results |

## 9. Reproducibility and Interpretation

The classification process is designed to support transparent and reproducible analysis.

Important considerations:

- Classifier development and independent validation are separate stages.
- Classification rules and scoring weights are documented.
- Manual and automated classifications are retained for comparison.
- Uncertain classifications are flagged for manual review.
- Exact-match metrics should be interpreted together with manual review rates.
- Classification results represent organizational activities and functions, not necessarily direct allocations of revenue to specific diseases.

The resulting datasets support further analysis of the relationship between nonprofit healthcare funding and disease burden measured in DALYs.
