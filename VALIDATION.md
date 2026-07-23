# Validation Summary

This repository contains the final non-financial and transaction-flow files supplied for the project. The scripts, controlled input and output workbooks were cross-checked as one package in July 2026.

## Package checks

- The four public scripts match the final non-financial and TFM workflow.
- Script filenames and repository-relative input/output paths are aligned.
- Every required controlled-input sheet is present.
- The four output workbooks match the final supplied outputs.
- No Excel formula-error values were found in the supplied workbooks.
- No `KADJ` or other forced balancing row is inserted.

## Stage 1

| Check | Annual | Quarterly |
|---|---:|---:|
| P6/P7 source availability | Available | Available |
| Total Economy codes present | 90 | 47 |
| Total Economy codes unavailable from source | 0 | 43 |
| General Government source | Direct S13 | 2,340 documented S1-residual rows |
| Adjustment basis | Not applicable | Non-seasonally adjusted (`N`) |
| Key QA issue sheets | No issues found | No issues found |

For the quarterly workflow, direct S13 observations were unavailable for the relevant extraction. General Government values are therefore derived transparently as `S1 - S11 - S12 - S1M` and recorded in dedicated QA sheets.

## Stage 2

| Metric | Annual | Quarterly |
|---|---:|---:|
| Periods | 16 | 65 |
| Range | 2010–2025 | 2010 Q1–2026 Q1 |
| Observed closures marked OK | 188 | 865 |
| Constructed closures | 144 | 390 |
| Rounding cases | 51 | 105 |
| Review cases | 1 | 5 |
| Unavailable production closures | 0 | 195 |
| Maximum domestic net-lending error | €0.2m | €0.3m |
| Maximum Rest-of-World net-lending error | €3.0m | €2.9m |
| Maximum Total Economy difference | €3.0m | €0.1m |

The review rows are retained rather than adjusted:

- Annual: 2018, `D42`, residual −€3m.
- Quarterly: 2020 Q1 `D7` (+€1m), 2020 Q3 `D7` (+€2m), 2020 Q4 `D7` (−€3m), 2023 Q2 `D41` (−€1m), and 2024 Q4 `D41` (−€1m).

Quarterly production closures are unavailable where the source does not provide the required P1/P2 sector observations. These are labelled as unavailable rather than treated as zero.

## Public scope

The financial-accounts experiment is not included. No final financial output workbook was supplied, and the case study does not claim a complete financial stock-flow reconciliation.
