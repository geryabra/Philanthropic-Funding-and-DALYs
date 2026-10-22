# DALY-Philanthropy-Mapping

This research project explores the relationship between philanthropic funding and global disease burden (DALYs).

The study maps nonprofit medical organizations to DALY disease categories using a combination of:

- Direct disease matching
- Alias-based mapping
- NLP analysis
- Entity extraction
- Semantic similarity
- Manual review and taxonomy refinement

---

## Research Workflow

1. Direct Disease Mapping
2. Alias-Based Matching
3. Classification into Groups A, B and C
4. NLP and Clustering Analysis
5. Entity Extraction
6. Semantic Similarity
7. Manual Review and Taxonomy Refinement

---

## Repository Structure

| File | Purpose |
|--------|---------|
| Org_mission_to_DALY_v8.py | Main research pipeline responsible for disease mapping, alias matching, and classification into Groups A, B and C |
| DALYS_2022.xlsx | WHO DALY reference dataset used as the disease burden taxonomy |
| 2022_Revenue.xlsx | Source nonprofit revenue dataset used throughout the study |
| FINAL_DATASET_STAGE1.xlsx | Stage 1 output containing organizations directly mapped to DALY diseases and categories |
| B_entity_extraction_mapping_v1.py | Entity extraction and DALY candidate mapping pipeline for unresolved Group B organizations |
| B_entity_extraction_mapping_v1.xlsx | Entity extraction analysis and candidate DALY mappings for Group B organizations |
| B_proxy_review_workflow_v1.xlsx | Manual review workflow used to validate candidate DALY assignments in Group B |
| C_entity_extraction_mapping_v1.py | Entity extraction and DALY candidate mapping pipeline for unresolved Group C organizations |
| C_entity_extraction_mapping_v1.xlsx | Entity extraction analysis and candidate DALY mappings for Group C organizations |

---

## Code-to-Output Relationships

| Code | Output |
|--------|---------|
| Org_mission_to_DALY_v8.py | FINAL_DATASET_STAGE1.xlsx |
| B_entity_extraction_mapping_v1.py | B_entity_extraction_mapping_v1.xlsx |
| C_entity_extraction_mapping_v1.py | C_entity_extraction_mapping_v1.xlsx |

---

## Dataset Groups

### Group A – Direct DALY Mapping

Organizations that could be directly mapped to a specific disease or DALY category using organization names, mission statements, disease keywords, and alias matching.

### Group B – Medical Organizations Without a Specific Disease Assignment

Organizations with a clear medical focus but without an explicit disease reference. These organizations were further analyzed using NLP, clustering, entity extraction, semantic similarity, and manual review.

### Group C – Organizations Without Sufficient DALY Mapping

Organizations that could not be confidently mapped to a disease, medical specialty, or DALY category using the existing mapping framework. Additional entity extraction and semantic analysis were performed to identify potential medical relationships.

---

## Documentation

This repository also contains an interim research report summarizing:

- Current findings
- Classification methodology
- Group A, B and C results
- Entity extraction and semantic similarity experiments
- Open methodological questions
- Proposed next research steps

---

## Research Objective

The primary objective of this research is to examine the relationship between philanthropic funding and global disease burden by mapping nonprofit medical organizations to DALY disease categories and comparing funding distribution with DALY-based disease burden estimates.

The project combines structured disease matching, NLP methods, semantic analysis, and manual review to improve disease attribution and identify opportunities for expanding DALY coverage across nonprofit medical organizations.
