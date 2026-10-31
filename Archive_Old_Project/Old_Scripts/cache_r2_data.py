"""
cache_r2_data.py
----------------
High-speed parallel pre-computation and caching of surprisal analysis matrices
(X_log, G_all, L_all, R2_full, component weights) for all 120 ROIs using 12 CPU workers.
Outputs Docs/r2_cache.pkl.
"""

import os
import re
import glob
import time
import pickle
from concurrent.futures import ProcessPoolExecutor, as_completed
import numpy as np
import pandas as pd

BASE_PATH = 'all patients (without isotope)'
TAILS_DIR = 'All_Tail_Proteins_Tables'
CACHE_FILE = os.path.join('Docs', 'r2_cache.pkl')
EPSILON = 1e-6

def compute_r2_vectorized(X_log, G, L):
    X_hat = G @ L.T
    diff = X_log - X_hat
    ss_res = np.sum(diff ** 2, axis=0)
    mean_x = np.mean(X_log, axis=0, keepdims=True)
    ss_tot = np.sum((X_log - mean_x) ** 2, axis=0)
    valid = ss_tot > 0
    r2_cells = np.where(valid, 1.0 - ss_res / np.where(valid, ss_tot, 1.0), np.nan)
    return float(np.nanmean(r2_cells))

def process_single_roi(f):
    roi_id = os.path.basename(f).replace('_tail_proteins.xlsx', '')
    m_p = re.match(r'^(P\d+)', roi_id)
    pat = m_p.group(1) if m_p else roi_id
    roi_folder = roi_id.replace(f'{pat}_', '')
    
    sa_dir = os.path.join(BASE_PATH, pat, roi_folder, 'surprisal analysis')
    g_file = os.path.join(sa_dir, 'G.xlsx')
    l_file = os.path.join(sa_dir, 'L.xlsx')
    data_file = os.path.join(BASE_PATH, pat, roi_folder, f'{roi_id}.xlsx')
    
    df_g = pd.read_excel(g_file, index_col=0)
    df_l = pd.read_excel(l_file, index_col=0)
    df_raw = pd.read_excel(data_file)
    
    proteins = df_g.index.tolist()
    X_raw = df_raw[proteins].values
    X_log = np.log(X_raw + EPSILON).T
    G_all = df_g.loc[proteins].values.astype(np.float32)
    L_all = df_l.values.astype(np.float32)
    
    r2_full = compute_r2_vectorized(X_log, G_all, L_all)
    
    l_energies = np.sum(L_all ** 2, axis=0)
    dev_energy = float(np.sum(l_energies[1:])) if len(l_energies) > 1 else 1.0
    
    roi_comp_weights = {}
    g_weights = {}
    for c_idx, col in enumerate(df_l.columns):
        if col == 'L0': continue
        g_num = int(re.search(r'\d+', col).group(0))
        g_key = f"{roi_id}_G{g_num}"
        weight = float(l_energies[c_idx] / dev_energy) if dev_energy > 0 else 0.0
        g_weights[g_key] = weight
        roi_comp_weights[g_num] = weight
        
    roi_data = {
        'roi_id': roi_id,
        'patient': pat,
        'n_cells': int(X_log.shape[1]),
        'proteins': proteins,
        'X_log': X_log.astype(np.float32),
        'G_all': G_all,
        'L_all': L_all,
        'r2_full': r2_full,
        'comp_weights': roi_comp_weights
    }
    return roi_id, roi_data, g_weights

def build_cache():
    print("=== Parallel Building R2 Cache for all 120 ROIs (12 Workers) ===", flush=True)
    files = sorted(glob.glob(os.path.join(TAILS_DIR, '*.xlsx')))
    total_files = len(files)
    print(f"Target: {total_files} ROIs from {TAILS_DIR}", flush=True)
    
    t0 = time.time()
    cache_data = {}
    all_g_weights = {}
    
    with ProcessPoolExecutor(max_workers=12) as executor:
        futures = {executor.submit(process_single_roi, f): f for f in files}
        completed = 0
        for fut in as_completed(futures):
            roi_id, r_data, g_w = fut.result()
            cache_data[roi_id] = r_data
            all_g_weights.update(g_w)
            completed += 1
            if completed % 15 == 0 or completed == total_files:
                print(f"  [{completed:>3}/{total_files}] Completed in {time.time()-t0:.1f}s (Latest: {roi_id}, R2_full={r_data['r2_full']:.3f})", flush=True)

    os.makedirs(os.path.dirname(CACHE_FILE), exist_ok=True)
    with open(CACHE_FILE, 'wb') as fp:
        pickle.dump({'cache_data': cache_data, 'g_weights': all_g_weights}, fp, protocol=pickle.HIGHEST_PROTOCOL)
        
    print(f"\nCache successfully saved to {CACHE_FILE} (Size: {os.path.getsize(CACHE_FILE)/(1024*1024):.1f} MB)", flush=True)
    print(f"Total parallel execution time: {time.time()-t0:.2f}s.", flush=True)

if __name__ == '__main__':
    build_cache()
