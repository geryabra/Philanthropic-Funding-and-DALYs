# Stage 01 — Initial DALY Mapping

## Purpose

This stage establishes the initial mapping between nonprofit healthcare organizations and disease categories from the Disability-Adjusted Life Years (DALYs) framework.

## Research Workflow

1. Load nonprofit revenue and organizational mission data.
2. Load the DALY disease burden reference dataset.
3. Apply direct disease matching and alias-based mapping.
4. Classify organizations into Groups A, B, and C.
5. Investigate unresolved organizations using entity extraction, NLP methods, and manual review.

## Organization Groups

- **Group A:** Organizations directly mapped to a disease or DALY category.
- **Group B:** Healthcare-related organizations without a specific disease assignment.
- **Group C:** Organizations without sufficient evidence for disease-level mapping.

## Main Files

| File | Purpose |
|---|---|
| `2022_Revenue.xlsx` | Nonprofit revenue and organization data |
| `DALYS_2022.xlsx` | DALY disease burden reference |
| `Org_mission_to_DALY_v8.py` | Initial disease mapping and A/B/C classification |
| `FINAL_DATASET_STAGE1.xlsx` | Initial mapping results |
| `B_entity_extraction_mapping_v1.py` | Entity extraction and candidate disease mapping for Group B |
| `B_proxy_review_workflow_v1.xlsx` | Review of proxy disease assignments for Group B |
| `C_entity_extraction_mapping_v1.py` | Entity extraction for Group C |
| `C_entity_extraction_mapping_v1.xlsx` | Group C entity extraction results |

## Downstream Research

Unresolved organizations are subsequently analyzed through the DALY–MeSH dictionary and activity classification workflows documented in Stages 02 and 03.
