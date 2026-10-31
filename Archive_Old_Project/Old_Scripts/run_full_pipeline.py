"""
run_full_pipeline.py
--------------------
Orchestrates the entire cancer research sequence analysis pipeline from start to finish:

Stage 1: Centralize and copy tail_proteins.xlsx from all 120 patient ROI folders to All_Tail_Proteins_Tables/
Stage 2: Export unified TAILS table (Docs/All_Tails.xlsx, Docs/All_Tails.csv)
Stage 3: Mine and generate all unique ES combinations (Docs/ES_Combinations.csv)
Stage 4: Run optimal ES finder to select minimal ES set with maximum ROI coverage,
         conflict-free G assignment, and export rich 6-sheet Excel report (Docs/Optimal_ES_Result.xlsx)
"""

import sys
import time
import subprocess

def run_step(script_name: str, step_title: str):
    print("\n" + "=" * 75)
    print(f"STEP: {step_title}")
    print(f"Running: {script_name} ...")
    print("=" * 75)
    t0 = time.time()
    result = subprocess.run([sys.executable, script_name], capture_output=False)
    if result.returncode != 0:
        print(f"\n[ERROR] {script_name} failed with return code {result.returncode}!")
        sys.exit(result.returncode)
    print(f"Completed {script_name} in {time.time() - t0:.2f}s")

def main():
    t_start = time.time()
    print("*" * 75)
    print("STARTING COMPLETE PROTEIN SEQUENCE & ES OPTIMIZATION PIPELINE")
    print("*" * 75)

    # Stage 1: Copy tail tables
    run_step("copy_all_tail_tables.py", "Centralize Patient Tail Tables")

    # Stage 2: Export unified TAILS
    run_step("export_tails_table.py", "Export Unified All_Tails Table")

    # Stage 3: Generate ES combinations
    run_step("generate_es_combinations.py", "Mine All Unique ES Combinations (No Signs)")

    # Stage 4: Optimal ES Finder
    run_step("optimal_es_finder.py", "Optimal ES Set Cover & Excel Generation")

    total_time = time.time() - t_start
    print("\n" + "*" * 75)
    print(f"ALL PIPELINE STAGES COMPLETED SUCCESSFULLY IN {total_time:.2f} SECONDS!")
    print("*" * 75)

if __name__ == '__main__':
    main()
