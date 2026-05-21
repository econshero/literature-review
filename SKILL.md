---
name: literature-review-extractor
description: Extract structured information from academic economics papers, PDFs, DOIs, journal links, article titles, working papers, or literature collections. Use when Codex needs to produce single-paper summaries, multi-paper literature reviews, methodology summaries, regression-table extraction, identification strategy analysis, robustness checks, heterogeneity analysis, variable construction, empirical findings synthesis, or substantive answers about an economics paper's content.
---

# Literature Review Extraction

Use this skill for academic economics papers and literature collections when the user asks for literature review, evidence synthesis, regression extraction, methodology comparison, identification strategy, robustness checks, heterogeneity analysis, empirical findings, variable construction, or substantive interpretation.

Author: Ms Ma (PhD, Economics). Version: 2.0.0. Workflow stage: literature.

## Core Principles

- Prioritise regression tables, appendix tables, robustness tables, heterogeneity tables, and specification tables over narrative discussion.
- If narrative text conflicts with regression tables, trust the tables.
- Extract coefficients directly from tables whenever possible.
- Preserve table number, column number, fixed effects, clustering level, weighting structure, and specification details.
- Do not summarise empirical findings without exact specification, table location, and statistical significance.
- Do not hallucinate unavailable information. Use `NA` for missing values.
- Preserve econometric terminology accurately.
- Use formal academic English with British spelling.
- Distinguish causal claims, correlational findings, and descriptive evidence.

## PDF Reading Priority

Use efficient extraction tools before sequential reading:

1. `pdfgrep` for keyword search, section location, regression tables, robustness, heterogeneity, appendix, model specifications, and identification language.
2. `pdfplumber` for regression table extraction, structured table parsing, appendix tables, coefficients, standard errors, and statistical results.
3. `pandoc` for conversion into markdown, hierarchy preservation, section indexing, and semantic parsing.

Useful search terms include: `robustness`, `heterogeneity`, `placebo`, `identification`, `parallel trends`, `fixed effects`, `clustered`, `standard errors`, `instrument`, and `event study`.

Avoid naive full-document sequential reading unless necessary.

## Single-Paper Extraction

For a single paper, structure the output as:

1. Basic Information
2. Research Question
3. Data and Sample
4. Sample Restrictions
5. Methodology
6. Identification Strategy
7. Variable Construction
8. Baseline Results
9. Robustness Checks
10. Heterogeneity Analysis
11. Mechanism Discussion
12. Main Contributions
13. Limitations

Extract basic information: publication year, title, authors, citation, journal, DOI, working paper series, and publication status.

Identify the research question, hypothesis, policy relevance, economic mechanism, and theoretical motivation.

Extract data information: country or region, dataset name, sample period, sample size, unit of observation, panel structure, weighting scheme, and data type, such as cross-sectional, panel, repeated cross-section, administrative, survey, experimental, census, longitudinal, RCT, or registry.

Extract major sample restrictions: gender, age, employment, marital status, fertility, region, sector, or balanced-panel requirements.

Extract methodology: empirical strategy, econometric models, identification strategy, OLS, IV, DiD, event study, RDD, matching, FE models, structural models, synthetic control, machine learning methods, and estimation assumptions.

## Identification Validity

For causal papers, explicitly extract identifying assumptions: exclusion restrictions, parallel trends, continuity, monotonicity, and exogeneity.

Record whether the paper provides balance tests, placebo tests, pre-trend tests, falsification tests, weak IV diagnostics, or sensitivity analysis.

Classify evidence strength as `strong causal`, `moderate causal`, `weak causal`, `correlational`, or `descriptive`. Justify the classification using identification strength, robustness evidence, specification quality, and validity tests.

## Variable Extraction

Summarise dependent variables by category, such as employment, labour supply, wages, fertility, education, health, consumption, productivity, migration, and firm outcomes.

Extract key independent variables: treatment variables, policy variables, instrumental variables, interaction terms, and exposure variables.

Summarise control variables by category: demographic, household, regional, macroeconomic, occupational, firm-level, fixed effects, and time trends.

## Empirical Results

When multiple specifications exist, prioritise:

1. Authors' preferred specification
2. Fully controlled specification
3. Specifications with fixed effects, clustered standard errors, and preferred identification strategy
4. Baseline specification only if no preferred model exists

Always record table number, column number, fixed effects, clustering level, coefficients, standard errors, confidence intervals, significance stars, specification details, coefficient direction, economic magnitude, and interpretation.

Whenever possible, standardise interpretation into percentage changes, percentage-point changes, log-point effects, standard-deviation effects, or elasticities. Explicitly distinguish level effects, log effects, marginal effects, odds ratios, and hazard ratios.

## Robustness And Heterogeneity

Extract robustness strategies: alternative variable definitions, alternative samples, placebo tests, lag structures, additional controls, clustering changes, weighting, trimming, winsorisation, alternative estimators, falsification tests, bandwidth sensitivity, and alternative fixed-effect structures.

Summarise whether coefficient signs remain stable, magnitudes change, significance changes, and conclusions remain valid. Identify checks that succeed, weaken conclusions, or fail.

Extract heterogeneity designs: gender, age, region, income, hukou, education, sector, occupation, industry, policy exposure, and baseline characteristics.

Summarise subgroup differences, implied mechanisms, statistically significant contrasts, and economically meaningful differences.

Classify mechanisms as:

1. Direct empirical mechanism evidence
2. Indirect suggestive evidence
3. Pure author interpretation or speculation

## Multi-Paper Literature Review

For multiple papers, organise the output as:

1. Research Theme
2. Literature Classification
3. Data Comparison
4. Method Comparison
5. Identification Strategy Comparison
6. Empirical Findings Comparison
7. Mechanism Comparison
8. Robustness Comparison
9. Research Gaps
10. Future Research Directions

Organise findings primarily by outcome category rather than paper chronology. Use one row per paper-by-outcome-variable combination. Separate employment, labour supply, wage, fertility, education, health, and other outcomes.

Avoid vague summaries such as "the paper finds significant effects." Specify direction, magnitude, significance, and specification context. Connect findings to theoretical mechanisms, identification assumptions, and policy implications where possible.

## Excel Database Workflow

After every successful extraction, append one row per outcome variable to `{LIT_REVIEW_DIR}/lit-review.xlsx` on the `literature` sheet.

Before appending, check whether `paper_id` already exists. If a duplicate exists, ask whether to overwrite, append additional outcomes, or skip.

Use `openpyxl` to open the workbook, locate the `literature` sheet, verify schema consistency, append rows, preserve frozen panes, filters, formatting, and column widths, then save safely.

Required schema columns:

```text
paper_id
year
authors
title
journal
doi
research_question
country_region
sample_period
dataset_name
data_type
sample_size
sample_restrictions
econometric_method
identification_strategy
fixed_effects
cluster_level
outcome_variable
outcome_category
outcome_definition
treatment_variable
control_variables
coefficient
std_error
significance
ci_low
ci_high
economic_magnitude
model_specification
robustness_summary
heterogeneity_summary
mechanism_classification
mechanism_interpretation
causal_claim
evidence_strength
main_contribution
limitations
pdf_path
notes
date_added
```

Use `NA` for missing values. Never invent values.

Reference append pattern:

```python
from openpyxl import load_workbook
from datetime import date

SCHEMA_COLUMNS = [
    "paper_id", "year", "authors", "title", "journal", "doi",
    "research_question", "country_region", "sample_period", "dataset_name",
    "data_type", "sample_size", "sample_restrictions", "econometric_method",
    "identification_strategy", "fixed_effects", "cluster_level",
    "outcome_variable", "outcome_category", "outcome_definition",
    "treatment_variable", "control_variables", "coefficient", "std_error",
    "significance", "ci_low", "ci_high", "economic_magnitude",
    "model_specification", "robustness_summary", "heterogeneity_summary",
    "mechanism_classification", "mechanism_interpretation", "causal_claim",
    "evidence_strength", "main_contribution", "limitations", "pdf_path",
    "notes", "date_added",
]

wb = load_workbook(xlsx_path)
ws = wb["literature"]

for row_dict in extracted_rows:
    ws.append([row_dict.get(col, "NA") for col in SCHEMA_COLUMNS])

wb.save(xlsx_path)
```
