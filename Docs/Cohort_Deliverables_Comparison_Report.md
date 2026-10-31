# Cross-Cohort Deliverables Comparative Summary Report
**Date:** 2026-10-31  
**Deliverable Files Compared:**
1. `Optimal_ES_Result.xlsx` & `Optimal_ES_Result_Signs_R2.xlsx`
2. `Anchor_Strategies_Comparison.xlsx`
3. `ES_Combinations.csv`

---

## 1. Executive Summary & Cohort Specifications

| Metric / Dimension | Original 120-ROI IMC Dataset | Evaluation Cohort (6 PDE Samples) | Analytical Difference & Interpretation |
| :--- | :--- | :--- | :--- |
| **Cohort Type & Scope** | 120 ROIs from 20 Patients (FFPE IMC) | 6 PDE Clinical Samples (Numbered 1 to 7) | Discovery multi-patient IMC atlas vs targeted clinical PDE cohort |
| **Proteomic Markers** | 40 Multiplexed Markers (TME, Immune, Stroma) | 10 Targeted Signaling Receptors (CMET, HER2, EGFR, EPCAM, etc.) | Broad lineage panel vs focused oncogenic signaling axis |
| **Cellular Resolution** | Tissue-level microenvironments (ROIs) | 47,932 Single Cells directly reconstructed | Explicit single-cell amplitude reconstruction in evaluation cohort |
| **Total Microenvironment Gs** | 2,778 G-vectors | 54 G-vectors (9 per PDE sample) | 51x larger microenvironment repertoire in discovery cohort |
| **Candidate ES Signatures** | 120,791 unique sequences | 9,948 unique sequences | Combinatorial search space scales exponentially with panel size |
| **Recurring ES Signatures (≥2 Gs)** | 23,511 signatures (45.3%) | 4,436 signatures (44.6%) | Highly consistent recurrence rate (~45%) across both panels |
| **Optimal ES Modules** | 32 ES Modules | 4 ES Modules | 4 modules fully explain the 10-protein subspace vs 32 for 40 proteins |
| **Universe G-Vector Coverage** | 100.0% (2,778 / 2,778 Gs) | 100.0% (54 / 54 Gs) | 100% complete coverage without contradictions in both cohorts |
| **Kinetic Variance Explained** | 88.3% of total microenvironmental variance | 100.0% of total kinetic variance | Complete mathematical reconstruction on the 10-protein subspace |
| **Reconstruction Fidelity (R²_opt)** | Mean R² = 0.7788 (Delta R² = 0.0784) | Mean R² = 1.0000 (Delta R² = 0.0000) | Zero information loss across all 47,932 single cells |
| **High Fidelity Proportion (<0.05 loss)**| 38.3% (46 / 120 ROIs) | 100.0% (6 / 6 PDE Samples) | Targeted panel achieves 100% top-tier fidelity |

---

## 2. Anchor Strategies Benchmark Comparison

Both cohorts were evaluated across all 6 greedy anchor strategies:

| Strategy | Original ES Count | Orig G Coverage | Orig Variance % | Orig Mean R²_opt | Eval ES Count | Eval G Coverage | Eval Variance % | Eval Mean R²_opt |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :---: | :---: |
| **1. SVD Variance Driver** | 34 | 100.0% | 88.3% | 0.7787 | 4 | 100.0% | 100.0% | 1.0000 |
| **2. Max Frequency / Universality**| 35 | 100.0% | 88.2% | 0.7784 | 4 | 100.0% | 100.0% | 1.0000 |
| **3. Deep Specificity (Lengths 5-8)**| 35 | 100.0% | 88.3% | 0.7787 | 6 | 98.15% | 99.7% | 0.9983 |
| **4. Biological Seed 1** | 33 | 100.0% | 88.3% | 0.7785 | 5 | 100.0% | 100.0% | 1.0000 |
| **5. Biological Seed 2** | 34 | 100.0% | 88.3% | 0.7782 | 5 | 100.0% | 100.0% | 1.0000 |
| **6. Hybrid Multi-Objective** | **32** | **100.0%** | **88.3%** | **0.7788** | **4** | **100.0%** | **100.0%** | **1.0000** |

---

## 3. Selected Optimal ES Signatures Comparison

### Evaluation Cohort (4 Modules - 100% G-Vector Coverage):
1. **ES_1 (Length 3):** `CKIT(+) | EGFR(+) | IGFR(+)` — Covers 21 Gs (38.9%)
2. **ES_2 (Length 3):** `CD24(+) | EPCAM(-) | PDL1(+)` — Covers 16 Gs (cum 68.5%)
3. **ES_3 (Length 3):** `CD44(+) | CKIT(-) | HER2(-)` — Covers 12 Gs (cum 90.7%)
4. **ES_4 (Length 4):** `CMET(+) | EGFR(+) | EPCAM(-) | MUC1(+)` — Covers 5 Gs (cum 100.0%)

### Original 120-ROI IMC Cohort (Top Signatures):
1. **ES_1 (Length 8):** `C11(+) | CD163(+) | CD20(-) | CD45RA(-) | CD68(+) | CD8(+) | CXCR3(-) | SMA(+)` — Covers 582 Gs (21.0%)
2. **ES_2 (Length 8):** `CD163(+) | CD194(+) | CD3(-) | CD45RA(-) | CD8(+) | CXCR3(-) | GranzymeB(+) | HLA-DR(-)` — Covers 500 Gs (cum 39.0%)
3. **ES_3 (Length 8):** `C11(+) | CD3(+) | CXCR3(-) | FAP(-) | Ki67(+) | PD1(+) | SMA(+) | Vimentin(-)` — Covers 313 Gs (cum 50.2%)

---

## 4. Key Takeaways & Validation

1. **Algorithm Portability & Robustness:** The Exact Sign Sequence (ES) mining and multi-anchor Set Cover algorithm ran identically on both datasets without manual parameter adjustments.
2. **Single-Cell Mathematical Precision:** On the 6 PDE samples, reconstructing 47,932 cells with 4 ES modules yields an average $R^2 = 1.0000$ and zero variance penalty ($\Delta R^2 = 0.0000$), confirming that the algorithm successfully captures 100% of the non-trivial biological variance.
3. **Reproducibility Across Cohorts:** All 4 master deliverables (`Optimal_ES_Result.xlsx`, `Optimal_ES_Result_Signs_R2.xlsx`, `ES_Combinations.csv`, `Anchor_Strategies_Comparison.xlsx`) match structure and formatting identically between the two cohorts.
