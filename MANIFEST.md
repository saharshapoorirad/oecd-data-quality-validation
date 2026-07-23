# Repository Manifest

## Final scripts

- `R/01_nonfinancial_annual.R` — annual OECD retrieval, mapping, relabeling and QA
- `R/02_nonfinancial_quarterly.R` — quarterly OECD retrieval, mapping, General Government derivation, relabeling and QA
- `R/03_tfm_annual.R` — annual transaction-flow matrix construction and validation
- `R/04_tfm_quarterly.R` — quarterly transaction-flow matrix construction and validation

## Controlled input

- `data/FINAL_nonfinancial_accounts_input.xlsx`
  - `Readme`
  - `glossary`
  - `sectors`
  - `api_code_dict`
  - `mapping`
  - `source_reference`
  - `total_economy_labels`

## Validated outputs

- `outputs/ITA_stage1_nonfinancial_accounts_annual_output.xlsx`
- `outputs/ITA_stage1_nonfinancial_accounts_quarterly_output.xlsx`
- `outputs/ITA_stage2_TFM_annual_output.xlsx`
- `outputs/ITA_stage2_TFM_quarterly_output.xlsx`

## Portfolio and verification documents

- `docs/OECD_Data_Quality_Case_Study.pdf`
- `docs/case-study-cover.png`
- `VALIDATION.md`

## README figures

- `figures/01_closure_status.png`
- `figures/02_max_validation_differences.png`
- `figures/03_review_cases.png`

## Intentionally excluded

- `oecd_relabel_export_final_quarterly.R` and `oecd_relabel_inputs.xlsx` — superseded by the final Stage 1 quarterly script and controlled input workbook.
- `Financial_accounts.R` and `financial_accounts_input.xlsx` — separate experimental extension without a complete validated output in this case-study package.
- Numbered duplicate filenames and earlier adjustment-based versions.
