"""
01_extract_tail_tables.py
-------------------------
Step 1 of the Cancer Research TME Pipeline:
Extracts all 'tail_proteins.xlsx' from nested patient directories
under 'all patients (without isotope)/' and centralizes them into
'All_Tail_Proteins_Tables/' with standard naming: P{X}_ROI_{Y}_tail_proteins.xlsx
"""

import os
import shutil
import glob
import time

SRC_ROOT = r'all patients (without isotope)'
DEST_DIR = r'All_Tail_Proteins_Tables'

def main():
    t0 = time.time()
    print("====================================================================")
    print(" [Step 1/5] Extracting Tail Protein Tables from Patient Folders")
    print("====================================================================")
    
    pattern = os.path.join(SRC_ROOT, '*', '*', 'surprisal analysis', 'tail_proteins.xlsx')
    files = sorted(glob.glob(pattern))

    if not files:
        print(f"ERROR: No files found with pattern:\n  {pattern}")
        print("Please ensure 'all patients (without isotope)' directory exists.")
        return 1

    os.makedirs(DEST_DIR, exist_ok=True)
    print(f"Found {len(files)} ROI tail tables. Copying to '{DEST_DIR}/'...\n")

    copied = 0
    for src in files:
        parts = os.path.normpath(src).split(os.sep)
        try:
            sa_idx = parts.index('surprisal analysis')
        except ValueError:
            print(f"  SKIP (unexpected path): {src}")
            continue
        
        roi_dir = parts[sa_idx - 1]   # e.g. ROI_26
        p_dir   = parts[sa_idx - 2]   # e.g. P10
        
        dest_name = f"{p_dir}_{roi_dir}_tail_proteins.xlsx"
        dest = os.path.join(DEST_DIR, dest_name)
        
        shutil.copy2(src, dest)
        copied += 1
        if copied <= 5 or copied % 25 == 0 or copied == len(files):
            print(f"  {copied:>3}. {p_dir}/{roi_dir} -> {dest_name}")

    print(f"\nStep 1 Complete: {copied} tail tables centralized in {time.time()-t0:.2f}s.\n")
    return 0

if __name__ == '__main__':
    main()
