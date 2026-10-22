# Philanthropic Funding and Global Disease Burden (DALYs)

## Research Overview

This repository documents an academic research project examining the relationship between philanthropic funding for healthcare and the global burden of disease, measured using Disability-Adjusted Life Years (DALYs).

The research investigates how nonprofit healthcare organizations can be linked to disease categories and medical areas, with the broader objective of comparing the distribution of philanthropic funding with disease burden.

A central methodological challenge is that nonprofit organizations do not always focus on a single disease. Their activities may address broader medical fields, healthcare services, research, or patient support.

The project therefore combines disease mapping, medical terminology dictionaries, and organization activity classification.

## Research Objectives

1. Identify relationships between nonprofit healthcare organizations and DALY disease categories.
2. Develop methods for identifying disease-related information in organizational names and mission statements.
3. Investigate organizations that cannot be directly assigned to a specific disease.
4. Evaluate the use of medical terminology and rule-based classification to improve the interpretation of organizational activities.
5. Establish a transparent and reproducible foundation for analyzing philanthropic funding in relation to global disease burden.

## Repository Structure

### 01 — Initial DALY Mapping

**Directory:** `01_Initial_DALY_Mapping/`

Establishes the initial mapping between nonprofit organizations and DALY disease categories.

The workflow includes direct disease matching, alias-based matching, classification into Groups A, B, and C, NLP analysis, entity extraction, semantic similarity, and manual review.

This stage includes the initial datasets, Python scripts, intermediate outputs, and methodological documentation.

### 02 — DALY–MeSH Dictionary

**Directory:** `02_MeSH_Dictionary/`

Investigates the use of Medical Subject Headings (MeSH) to expand and standardize the terminology associated with DALY disease categories.

The work includes disease terminology, synonyms, hierarchical relationships, and the examination of potential matches between organizational mission statements and medical concepts.

### 03 — Organization Activity Classification

**Directory:** `03_Activity_Classification/`

Develops and evaluates a rule-based approach to classifying nonprofit healthcare organizations along two dimensions:

- **Main Theme:** The primary area or focus of organizational activity.
- **Organization Type:** The type or functional role of the organization.

This stage includes classifier development, manually reviewed samples, and independent validation of classification performance.

## Research Methodology

The research combines several analytical approaches:

- Structured matching between organizational information and disease categories.
- Natural Language Processing (NLP) and entity extraction.
- Medical terminology mapping using MeSH.
- Rule-based organization classification.
- Manual review and validation.
- Quantitative analysis of philanthropic funding and disease burden.

The methodology is developed iteratively, with attention to mapping coverage, classification uncertainty, and the limitations of assigning broad organizational activities to individual diseases.

## Research Status

This repository documents ongoing academic research.

The datasets, scripts, classification methods, and intermediate results may be revised as the methodology develops.

Results and classifications should therefore be interpreted as research-stage outputs rather than final conclusions.

## Documentation

Each research-stage directory contains its own documentation describing the relevant methodology, files, and outputs.

The repository is organized to support transparency, reproducibility, and further methodological development.
