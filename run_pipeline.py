"""
run_pipeline.py
---------------
Master Pipeline Orchestrator for Cancer Research TME Analysis.
Runs the complete workflow from raw patient files to final master deliverables:

  Step 1: 01_extract_tail_tables.py        -> Extracts tails to All_Tail_Proteins_Tables/
  Step 2: 02_cache_r2_data.py              -> Precomputes SVD matrices & variance cache (Docs/r2_cache.pkl)
  Step 3: 03_mine_all_es_combinations.py   -> Mines all 120,791 candidate combinations (lengths 3 to 8)
  Step 4: 04_compare_anchor_strategies.py  -> Evaluates 6 anchor strategies & selects winning Hybrid model
  Step 5: 05_run_winning_pipeline.py       -> Builds final master Excel & R2 reports (Docs/Optimal_ES_Result_Signs_R2.xlsx)

Usage:
  python run_pipeline.py             # Runs all 5 steps end-to-end
  python run_pipeline.py --step 5    # Runs only Step 5
  python run_pipeline.py --from 3    # Runs from Step 3 to Step 5
"""

import sys
import time
import argparse
import subprocess

STEPS = [
    (1, "01_extract_tail_tables.py", "Extract Tail Protein Tables from Patient Folders"),
    (2, "02_cache_r2_data.py", "Precompute Surprisal Analysis Matrices & Variance Cache"),
    (3, "03_mine_all_es_combinations.py", "Mine All Candidate ES Combinations (Lengths 3 to 8)"),
    (4, "04_compare_anchor_strategies.py", "Evaluate 6 Greedy Anchor Strategies & R2 Optimization"),
    (5, "05_run_winning_pipeline.py", "Build Master Excel Deliverables (Optimal_ES_Result_Signs_R2.xlsx)"),
]

def run_step(step_num, script_name, description):
    print(f"\n>>> Running [Step {step_num}/5]: {description} ({script_name})...", flush=True)
    t0 = time.time()
    res = subprocess.run([sys.executable, script_name], capture_output=False)
    if res.returncode != 0:
        print(f"\n[ERROR] Step {step_num} failed with return code {res.returncode}!", flush=True)
        return False
    print(f">>> [Step {step_num}/5] Finished successfully in {time.time()-t0:.2f}s.\n", flush=True)
    return True

def main():
    parser = argparse.ArgumentParser(description="Master Pipeline Runner for Cancer Research TME Project")
    parser.add_argument("--step", type=int, choices=[1, 2, 3, 4, 5], help="Run only a specific step")
    parser.add_argument("--from-step", type=int, choices=[1, 2, 3, 4, 5], default=1, help="Start from a specific step")
    args = parser.parse_args()

    t_start = time.time()
    print("=" * 75)
    print("  CANCER RESEARCH TME SURPRISAL ANALYSIS PIPELINE")
    print("  Exact Sign Sequences (ES), Multi-Anchor Set Cover & R2 Optimization")
    print("=" * 75)

    if args.step:
        steps_to_run = [s for s in STEPS if s[0] == args.step]
    else:
        steps_to_run = [s for s in STEPS if s[0] >= args.from_step]

    for step_num, script_name, desc in steps_to_run:
        ok = run_step(step_num, script_name, desc)
        if not ok:
            sys.exit(1)

    print("=" * 75)
    print(f"  ALL PIPELINE STEPS COMPLETED SUCCESSFULLY IN {time.time()-t_start:.2f}s!")
    print("  Deliverables:")
    print("  - Docs/Optimal_ES_Result_Signs_R2.xlsx (Master 8-sheet Excel deliverable)")
    print("  - Docs/Anchor_Strategies_Comparison.xlsx (6-anchor comparison + winning 32 ES)")
    print("  - Docs/R2 comparison per ROI.xlsx (Standalone R2 report)")
    print("  - Docs/Algorithm_Full_Summary.html (Interactive documentation)")
    print("=" * 75)

if __name__ == '__main__':
    main()
