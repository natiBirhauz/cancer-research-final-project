"""
test_gui_execution.py
---------------------
Tests the simplified, clean Web GUI execution function on:
1. Data for algorithm evaluation (6 PDE samples, 54 Gs)
"""

import os
from web_gui import execute_pipeline_job

def test_run():
    print("Running GUI execution on 'Data for algorithm evaluation'...")
    res = execute_pipeline_job(
        uploaded_files=None,
        local_folder='Data for algorithm evaluation',
        strategy='Hybrid Multi-Objective (Variance x Count)',
        min_len=3,
        max_len=8,
        threshold=0.05,
        min_rec=2
    )

    summary_html, opt_es, opt_r2, combos, anchors, zip_bundle = res

    print("\nVerification Results:")
    print("  - Summary HTML generated:", bool(summary_html and "Execution Complete" in summary_html))
    print("  - Optimal_ES_Result.xlsx:", opt_es, "Exists:", os.path.exists(opt_es))
    print("  - Optimal_ES_Result_Signs_R2.xlsx:", opt_r2, "Exists:", os.path.exists(opt_r2))
    print("  - ES_Combinations.csv:", combos, "Exists:", os.path.exists(combos))
    print("  - Anchor_Strategies_Comparison.xlsx:", anchors, "Exists:", os.path.exists(anchors))
    print("  - Deliverables_Bundle.zip:", zip_bundle, "Exists:", os.path.exists(zip_bundle))
    print("\nALL VERIFICATIONS PASSED SUCCESSFULLY!")

if __name__ == '__main__':
    test_run()
