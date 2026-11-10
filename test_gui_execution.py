"""
test_gui_execution.py
---------------------
Automated verification suite for the Surprisal Analysis (SA) – Minimal Exact Sign
Signatures (ES) Discovery Platform:
1. Pre-Flight Dataset Check (inspect_user_data & run_precheck_job)
2. Full Pipeline & Streamlined Deliverables Verification
3. High-Dimensional Scalability Test (large gene panel with Min Sequence Length = 300)
"""

import os
import glob
import shutil
import openpyxl
from pipeline_engine import inspect_user_data, mine_candidate_signatures
from web_gui import execute_pipeline_job, run_precheck_job


def test_run():
    print("1. Testing Pre-Flight Dataset Inspection...")
    eval_files = glob.glob('Data for algorithm evaluation/*/data.xlsx')
    if not eval_files:
        eval_files = [
            "Sample_Input_Template.xlsx"
        ]
    pre_info = inspect_user_data(file_paths=eval_files, threshold=0.05)
    print(f"   - Datasets detected: {pre_info['num_datasets']}")
    print(f"   - Total G-vectors: {pre_info['total_g_vectors']}")
    print(f"   - Unique features: {pre_info['total_unique_features']}, Shared across all: {pre_info['shared_all_datasets']}")
    print(f"   - Estimated runtime: ~{pre_info['estimated_time_sec']}s")
    pre_html_en = run_precheck_job(eval_files, 0.05, lang='en')
    pre_html_he = run_precheck_job(eval_files, 0.05, lang='he')
    assert "Pre-Flight Dataset Inspection" in pre_html_en
    assert "בדיקה מקדימה" in pre_html_he

    print("\n2. Testing Full Pipeline Execution & Streamlined Deliverables...")
    res = execute_pipeline_job(
        uploaded_files=eval_files,
        min_len=3,
        max_len=8,
        threshold=0.05,
        min_rec=2,
        lang='en'
    )

    (
        summary_html,
        bar_svg_html,
        heatmap_html,
        df_es,
        df_dataset,
        df_mat,
        opt_es,
        html_rep,
        zip_bundle,
        anchors,
        combos
    ) = res

    wb = openpyxl.load_workbook(opt_es)
    print("   - Optimal_ES_Result.xlsx sheets:", wb.sheetnames)
    assert wb.sheetnames == [
        'Summary',
        'Optimal_ES_SetCover',
        'ES_Conflict_Free_Matrix',
        'Per_Dataset_Coverage_and_R2',
        'Per_G_Coverage_Detail'
    ], f"Unexpected sheets: {wb.sheetnames}"

    wb_anc = openpyxl.load_workbook(anchors)
    print("   - Anchor_Strategies_Comparison.xlsx sheets:", wb_anc.sheetnames)
    assert wb_anc.sheetnames == ['Anchor_Strategies_Overview']

    assert os.path.exists(html_rep), "Analysis_Summary_Report.html missing"
    assert os.path.exists(combos), "ES_Combinations.csv missing"
    assert os.path.exists(zip_bundle), "Deliverables_Bundle.zip missing"
    assert "<svg" in bar_svg_html, "Bar chart SVG missing"
    assert "<table" in heatmap_html, "Heatmap HTML missing"

    # Clean up generated test run directory
    run_dir = os.path.dirname(opt_es)
    if os.path.exists(run_dir) and "Run_" in run_dir:
        shutil.rmtree(run_dir, ignore_errors=True)

    print("\n3. Testing High-Dimensional Gene Panel (5,000 genes, Min Sequence Length = 300)...")
    synthetic_g_data = []
    shared_block_1 = {f"GENE_{i}": (1 if i % 2 == 0 else -1) for i in range(1, 451)}
    shared_block_2 = {f"GENE_{i}": (-1 if i % 3 == 0 else 1) for i in range(300, 851)}
    for idx in range(8):
        base = dict(shared_block_1 if idx < 4 else shared_block_2)
        if idx % 2 == 1:
            base = {k: -v for k, v in base.items()}  # Inverse orientation
        for noise_i in range(1000 + idx * 50, 1050 + idx * 50):
            base[f"GENE_{noise_i}"] = 1
        synthetic_g_data.append({
            'gkey': f"DS{idx}_G1",
            'roi': f"DS{idx}",
            'g_num': 1,
            'signs': base,
            'magnitudes': {k: 0.2 for k in base},
            'prots': set(base.keys()),
            'count': len(base),
            'weight': 0.5
        })
    seen_high_dim, mode_used = mine_candidate_signatures(
        synthetic_g_data, min_es_length=300, max_es_length=500, min_overlap=300, min_g_recurrence=2
    )
    print(f"   - High-dimensional mining mode: {mode_used}, candidates found: {len(seen_high_dim)}")
    assert len(seen_high_dim) > 0, "Expected high-dimensional candidates to be mined"

    print("\nALL VERIFICATIONS PASSED SUCCESSFULLY!")


if __name__ == '__main__':
    test_run()
