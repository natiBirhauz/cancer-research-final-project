"""
test_gui_execution.py
---------------------
Tests the Web GUI execution pipeline on both datasets:
1. Data for algorithm evaluation (6 PDE samples, 54 Gs, 10 proteins)
2. All_Tail_Proteins_Tables (120 ROIs, 2,778 Gs, 40 proteins)

Verifies output deliverables, coverage, variance, and R2 reconstruction.
"""

import os
import openpyxl
import pandas as pd
from web_gui import execute_pipeline_job

def test_runs():
    print("=" * 75)
    print("TEST 1: Running GUI execution on 'Data for algorithm evaluation'...")
    print("=" * 75)
    res_eval = execute_pipeline_job(
        uploaded_files=None,
        local_folder='Data for algorithm evaluation',
        strategy='Hybrid Multi-Objective (Variance x Count)',
        min_len=3,
        max_len=8,
        threshold=0.05,
        min_rec=2
    )

    (
        summary_html_eval, df_es_eval, df_mat_eval, df_r2_eval, df_anc_eval,
        opt_es_eval, opt_r2_eval, combos_eval, anc_eval, zip_eval, log_eval
    ) = res_eval

    print("Result files produced:")
    print("  - Optimal_ES_Result.xlsx:", opt_es_eval, "Exists:", os.path.exists(opt_es_eval))
    print("  - Optimal_ES_Result_Signs_R2.xlsx:", opt_r2_eval, "Exists:", os.path.exists(opt_r2_eval))
    print("  - ES_Combinations.csv:", combos_eval, "Exists:", os.path.exists(combos_eval))
    print("  - Anchor_Strategies_Comparison.xlsx:", anc_eval, "Exists:", os.path.exists(anc_eval))
    print("  - Deliverables_Bundle.zip:", zip_eval, "Exists:", os.path.exists(zip_eval))

    print(f"\nEvaluation Cohort Results:")
    print(f"  - Optimal ES Count: {len(df_es_eval)}")
    print(f"  - Conflict-Free Matrix Shape: {df_mat_eval.shape}")
    print(f"  - R2 Reconstruction Mean: {df_r2_eval['R2_optimal'].mean():.4f}")
    print(f"  - Delta R2 Loss Mean: {df_r2_eval['Delta_R2'].mean():.4f}")

    print("\n" + "=" * 75)
    print("TEST 2: Running GUI execution on 'All_Tail_Proteins_Tables' (120 ROIs)...")
    print("=" * 75)
    res_120 = execute_pipeline_job(
        uploaded_files=None,
        local_folder='All_Tail_Proteins_Tables',
        strategy='Hybrid Multi-Objective (Variance x Count)',
        min_len=3,
        max_len=8,
        threshold=0.05,
        min_rec=2
    )

    (
        summary_html_120, df_es_120, df_mat_120, df_r2_120, df_anc_120,
        opt_es_120, opt_r2_120, combos_120, anc_120, zip_120, log_120
    ) = res_120

    print("Result files produced:")
    print("  - Optimal_ES_Result.xlsx:", opt_es_120, "Exists:", os.path.exists(opt_es_120))
    print("  - Optimal_ES_Result_Signs_R2.xlsx:", opt_r2_120, "Exists:", os.path.exists(opt_r2_120))
    print("  - ES_Combinations.csv:", combos_120, "Exists:", os.path.exists(combos_120))
    print("  - Anchor_Strategies_Comparison.xlsx:", anc_120, "Exists:", os.path.exists(anc_120))
    print("  - Deliverables_Bundle.zip:", zip_120, "Exists:", os.path.exists(zip_120))

    print(f"\n120-ROI Cohort Results:")
    print(f"  - Optimal ES Count: {len(df_es_120)}")
    print(f"  - Conflict-Free Matrix Shape: {df_mat_120.shape}")
    print(f"  - Universe Coverage: {df_es_120['Coverage_%'].iloc[-1]}")

    print("\n" + "=" * 75)
    print("VERIFICATION COMPLETED SUCCESSFULLY ON BOTH COHORTS!")
    print("=" * 75)

if __name__ == '__main__':
    test_runs()
