"""
copy_all_tail_tables.py
------------------------
Copies all tail_proteins.xlsx from nested patient directories to
All_Tail_Proteins_Tables with renamed format: P{X}_ROI_{Y}_tail_proteins.xlsx
"""

import os
import shutil
import glob

SRC_ROOT = r'all patients (without isotope)'
DEST_DIR = r'All_Tail_Proteins_Tables'

# Pattern: all patients (without isotope)\P*\ROI_*\surprisal analysis\tail_proteins.xlsx
pattern = os.path.join(SRC_ROOT, '*', '*', 'surprisal analysis', 'tail_proteins.xlsx')
files = sorted(glob.glob(pattern))

if not files:
    print(f'ERROR: No files found with pattern:\n  {pattern}')
    print(f'\nMake sure the directory structure exists:')
    print(f'  {SRC_ROOT}\\P10\\ROI_26\\surprisal analysis\\tail_proteins.xlsx')
    exit(1)

os.makedirs(DEST_DIR, exist_ok=True)
print(f'Found {len(files)} file(s). Copying to {DEST_DIR}...\n')

copied = 0
for src in files:
    parts = os.path.normpath(src).split(os.sep)
    try:
        sa_idx = parts.index('surprisal analysis')
    except ValueError:
        print(f'  SKIP (unexpected path): {src}')
        continue
    
    roi_dir = parts[sa_idx - 1]   # e.g. ROI_44
    p_dir   = parts[sa_idx - 2]   # e.g. P10
    
    dest_name = f'{p_dir}_{roi_dir}_tail_proteins.xlsx'
    dest = os.path.join(DEST_DIR, dest_name)
    
    shutil.copy2(src, dest)
    copied += 1
    if copied <= 10 or copied % 20 == 0 or copied == len(files):
        print(f'  {copied:>3}. {p_dir}/{roi_dir}  ->  {dest_name}')

print(f'\nDone. {copied} files copied to {DEST_DIR}/')
