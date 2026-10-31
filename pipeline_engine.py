"""
pipeline_engine.py
------------------
Unified execution engine for the Cancer Research TME Surprisal Analysis & 
Exact Sign Sequence (ES) Pipeline.

Supports:
1. Original 120-ROI IMC Dataset ('all patients (without isotope)' / 'All_Tail_Proteins_Tables')
2. Additional Evaluation Datasets ('Data for algorithm evaluation')
3. Custom user-selected dataset directories (auto-detects G/L or tail_proteins structure)
"""

import os
import re
import glob
import time
import pickle
from collections import defaultdict
from itertools import combinations
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def canon_dict(d):
    s = sorted(d.items(), key=lambda x: x[0])
    if s[0][1] == -1:
        return dict((p, -v) for p, v in s)
    return dict(s)

def tuple_to_str(d_or_tuple):
    if isinstance(d_or_tuple, dict):
        items = sorted(d_or_tuple.items())
    else:
        items = sorted(d_or_tuple)
    return " | ".join(f"{p}({'+' if s == 1 else '-'})" for p, s in items)

def match_es_g(es_dict, g_signs, g_prots, min_overlap=2):
    common = set(es_dict.keys()) & g_prots
    if len(common) < min_overlap:
        return None, 0, []
    if all(es_dict[p] == g_signs[p] for p in common):
        return '+', len(common), sorted(common)
    if all(es_dict[p] == -g_signs[p] for p in common):
        return '-', len(common), sorted(common)
    return None, 0, []

def compute_r2_vectorized(X_log, G, L, ss_tot_threshold=1e-3):
    X_hat = G @ L.T
    diff = X_log - X_hat
    ss_res = np.sum(diff ** 2, axis=0)
    mean_x = np.mean(X_log, axis=0, keepdims=True)
    ss_tot = np.sum((X_log - mean_x) ** 2, axis=0)
    valid = ss_tot > ss_tot_threshold
    if np.sum(valid) == 0:
        return 0.0
    r2_cells = 1.0 - (ss_res[valid] / ss_tot[valid])
    return float(np.mean(r2_cells))

def run_pipeline(
    data_source='evaluation',
    custom_path=None,
    min_es_length=3,
    max_es_length=8,
    threshold=0.05,
    min_overlap=2,
    min_g_recurrence=2,
    strategy='Hybrid',
    progress_callback=None
):
    """
    Executes the entire end-to-end pipeline.
    progress_callback(pct: float, message: str)
    """
    def log(msg, pct=None):
        if progress_callback:
            progress_callback(pct, msg)
        print(f"[{time.strftime('%H:%M:%S')}] {msg}")

    t0 = time.time()
    log("Initializing pipeline...", 0.05)

    g_data = []
    roi_info = {}
    
    # -------------------------------------------------------------
    # Step 1: Ingest Data
    # -------------------------------------------------------------
    if data_source == 'original':
        log("Loading 120-ROI IMC Original Dataset...", 0.10)
        from evaluate_algorithm import load_all_evaluation_datasets
        # Check cache or load tails
        tails_files = sorted(glob.glob('All_Tail_Proteins_Tables/*.xlsx'))
        cache_file = os.path.join('Docs', 'r2_cache.pkl')
        if not os.path.exists(cache_file):
            log("Cache not found, executing precomputation...", 0.15)
            os.system('python 02_cache_r2_data.py')
            
        with open(cache_file, 'rb') as fp:
            c = pickle.load(fp)
        cache_data = c['cache_data']
        g_weights = c['g_weights']
        
        for f in tails_files:
            roi_id = os.path.basename(f).replace('_tail_proteins.xlsx', '')
            pat = re.match(r'^(P\d+)', roi_id).group(1) if re.match(r'^(P\d+)', roi_id) else roi_id
            df = pd.read_excel(f)
            g_nums = sorted(set(int(re.search(r'G(\d+)', c).group(1)) for c in df.columns if re.match(r'G\d+[+-]$', c)))
            for n in g_nums:
                signs = {}
                for s, col in [('+', f'G{n}+'), ('-', f'G{n}-')]:
                    if col in df.columns:
                        for p in df[col].dropna():
                            p = str(p).strip()
                            if p:
                                signs[p] = 1 if s == '+' else -1
                if len(signs) >= min_overlap:
                    gkey = f"{roi_id}_G{n}"
                    w = g_weights.get(gkey, 0.0)
                    g_data.append({
                        'g_idx': len(g_data),
                        'gkey': gkey,
                        'roi': roi_id,
                        'patient': pat,
                        'g_num': n,
                        'signs': signs,
                        'prots': set(signs.keys()),
                        'count': len(signs),
                        'weight': w
                    })
        roi_info = cache_data
        log(f"Loaded {len(g_data):,} G-vectors across {len(cache_data)} ROIs.", 0.25)
        
    else:
        # Evaluation dataset or Custom folder
        target_dir = custom_path if (data_source == 'custom' and custom_path) else 'Data for algorithm evaluation'
        log(f"Scanning datasets in '{target_dir}'...", 0.10)
        
        dirs = sorted([d for d in os.listdir(target_dir) if os.path.isdir(os.path.join(target_dir, d))])
        for d in dirs:
            p_data = os.path.join(target_dir, d, 'data.xlsx')
            if not os.path.exists(p_data):
                continue
            raw_g = pd.read_excel(p_data, sheet_name='Gvalues', header=None)
            h_idx = 0
            for r in range(raw_g.shape[0]):
                row = [str(x).strip() for x in raw_g.iloc[r].tolist()]
                if 'G1' in row and 'G0' in row:
                    h_idx = r; break
                    
            prot_col = 0 if ('6 PDE' in d or '7 PDE' in d) else 11
            df_g = raw_g.iloc[h_idx+1:, :].copy()
            prots = df_g[prot_col].astype(str).str.strip().str.upper().tolist()
            
            g_cols = {
                str(raw_g.iloc[h_idx, c]).strip(): c 
                for c in range(raw_g.shape[1]) 
                if str(raw_g.iloc[h_idx, c]).strip().startswith('G')
            }
            
            n_comps = len(g_cols)
            G_full = np.zeros((len(prots), n_comps))
            for i in range(n_comps):
                if f'G{i}' in g_cols:
                    G_full[:, i] = df_g[g_cols[f'G{i}']].values.astype(float)
                
            raw_l = pd.read_excel(p_data, sheet_name='Lambavalues')
            l_cols = [f'L{i}' for i in range(n_comps) if f'L{i}' in raw_l.columns]
            L_full = raw_l[l_cols].values.astype(float)
            
            X_log = G_full @ L_full.T
            r2_full = compute_r2_vectorized(X_log, G_full, L_full)
            
            l_energies = np.sum(L_full ** 2, axis=0)
            dev_energy = float(np.sum(l_energies[1:])) if len(l_energies) > 1 else 1.0
            
            for g_num in range(1, n_comps):
                vals = G_full[:, g_num]
                signs = {}
                for p_name, val in zip(prots, vals):
                    if abs(val) >= threshold:
                        signs[p_name] = 1 if val > 0 else -1
                w = float(l_energies[g_num] / dev_energy) if dev_energy > 0 else 0.0
                gkey = f"{d}_G{g_num}"
                g_data.append({
                    'g_idx': len(g_data),
                    'gkey': gkey,
                    'roi': d,
                    'patient': d,
                    'g_num': g_num,
                    'signs': signs,
                    'prots': set(signs.keys()),
                    'count': len(signs),
                    'weight': w
                })
                
            roi_info[d] = {
                'roi_id': d,
                'patient': d,
                'n_cells': L_full.shape[0],
                'G_all': G_full,
                'L_all': L_full,
                'X_log': X_log,
                'r2_full': r2_full,
                'l_energies': l_energies,
                'dev_energy': dev_energy
            }
            
        log(f"Loaded {len(g_data):,} G-vectors from {len(roi_info)} datasets.", 0.30)

    # -------------------------------------------------------------
    # Step 2: Combinatorial ES Mining
    # -------------------------------------------------------------
    log(f"Mining canonical ES combinations (lengths {min_es_length} to {max_es_length})...", 0.40)
    combo_candidates = set()
    for g in g_data:
        items = list(g['signs'].items())
        n = len(items)
        for k in range(min_es_length, min(n + 1, max_es_length + 1)):
            for c in combinations(items, k):
                cd = canon_dict(dict(c))
                combo_candidates.add(tuple(sorted(cd.items())))
                
    log(f"Mined {len(combo_candidates):,} unique canonical candidate sequences.", 0.55)

    # -------------------------------------------------------------
    # Step 3: Greedy Set Cover
    # -------------------------------------------------------------
    log(f"Executing Greedy Set Cover with Strategy: '{strategy}'...", 0.65)
    covered = set()
    selected = []
    g_assignment = {}
    total_g = len(g_data)

    while len(covered) < total_g:
        best_cand = None
        best_score = -1.0
        best_gain = 0.0
        best_new_gs = []

        for cand in combo_candidates:
            cand_dict = dict(cand)
            cand_len = len(cand_dict)
            cand_gain = 0.0
            cand_new = []

            for g in g_data:
                if g['g_idx'] in covered:
                    continue
                sign, n_match, _ = match_es_g(cand_dict, g['signs'], g['prots'], min_overlap=min_overlap)
                if sign is not None:
                    cand_gain += g['weight']
                    cand_new.append((g['g_idx'], sign, n_match))

            if len(cand_new) >= min_g_recurrence:
                if strategy == 'SVD_Variance':
                    score = cand_gain
                elif strategy == 'Max_Frequency':
                    score = float(len(cand_new))
                elif strategy == 'Deep_Specificity':
                    score = cand_gain * (1.5 if cand_len >= 5 else 1.0)
                else:  # Hybrid
                    score = cand_gain * np.sqrt(len(cand_new))

                if score > best_score:
                    best_score = score
                    best_gain = cand_gain
                    best_cand = cand
                    best_new_gs = cand_new

        if best_cand is None or len(best_new_gs) < min_g_recurrence:
            break

        cand_dict = dict(best_cand)
        es_name = f"ES_{len(selected)+1:02d}"
        es_str = tuple_to_str(cand_dict)

        selected.append({
            'es_id': es_name,
            'dict': cand_dict,
            'str': es_str,
            'length': len(cand_dict),
            'new_gs': best_new_gs,
            'gain': best_gain,
            'n_new': len(best_new_gs)
        })

        for g_idx, sign, n_m in best_new_gs:
            covered.add(g_idx)
            g_assignment[g_idx] = {
                'es_id': es_name,
                'sign': sign,
                'n_match': n_m
            }

    cov_pct = (len(covered) / total_g) * 100.0 if total_g > 0 else 0.0
    tot_var_pct = sum(e['gain'] for e in selected) / len(roi_info) * 100.0 if roi_info else 0.0
    log(f"Greedy Cover Complete: {len(selected)} modules cover {len(covered):,}/{total_g:,} Gs ({cov_pct:.1f}%).", 0.75)

    # -------------------------------------------------------------
    # Step 4: Conflict-Free Allocation & R2 Calculation
    # -------------------------------------------------------------
    log("Constructing conflict-free allocation matrix & computing cell-level R2...", 0.85)
    matrix_rows = []
    r2_rows = []

    for r_id, info in roi_info.items():
        row_assign = {'ROI': r_id, 'Cells': info['n_cells']}
        selected_g_nums = [0]
        used_g_nums = set()

        for es in selected:
            es_id = es['es_id']
            cand_gs = [
                g for g in g_data
                if g['roi'] == r_id and g['g_idx'] in g_assignment and g_assignment[g['g_idx']]['es_id'] == es_id
            ]
            cand_gs = [g for g in cand_gs if g['g_num'] not in used_g_nums]
            if cand_gs:
                best_g = max(cand_gs, key=lambda x: (x['weight'], g_assignment[x['g_idx']]['n_match']))
                sign = g_assignment[best_g['g_idx']]['sign']
                row_assign[es_id] = f"G{best_g['g_num']}({sign})"
                selected_g_nums.append(best_g['g_num'])
                used_g_nums.add(best_g['g_num'])
            else:
                row_assign[es_id] = ""

        matrix_rows.append(row_assign)

        # R2 Calculation
        selected_g_nums = sorted(set(selected_g_nums))
        G_opt = info['G_all'][:, selected_g_nums]
        L_opt = info['L_all'][:, selected_g_nums]
        r2_opt = compute_r2_vectorized(info['X_log'], G_opt, L_opt)
        delta_r2 = info['r2_full'] - r2_opt

        status = 'Excellent' if delta_r2 < 0.05 else ('Acceptable' if delta_r2 <= 0.15 else 'Attention')
        r2_rows.append({
            'ROI': r_id,
            'Cells': info['n_cells'],
            'Selected_G_Count': len(selected_g_nums) - 1,
            'Selected_Gs': ", ".join(f"G{n}" for n in selected_g_nums if n > 0),
            'R2_full': info['r2_full'],
            'R2_optimal': r2_opt,
            'Delta_R2': delta_r2,
            'Status': status
        })

    df_matrix = pd.DataFrame(matrix_rows)
    df_r2 = pd.DataFrame(r2_rows)
    
    avg_r2_full = float(df_r2['R2_full'].mean())
    avg_r2_opt = float(df_r2['R2_optimal'].mean())
    avg_delta_r2 = float(df_r2['Delta_R2'].mean())
    
    # Format ES DataFrame
    es_records = []
    cum = 0
    for es in selected:
        cum += es['n_new']
        es_records.append({
            'ES_ID': es['es_id'],
            'Length': es['length'],
            'Exact_Sign_Sequence': es['str'],
            'G_Covered': es['n_new'],
            'Variance_Gain': round(es['gain'], 4),
            'Cum_Coverage_Pct': f"{(cum/total_g)*100.0:.1f}%"
        })
    df_es = pd.DataFrame(es_records)

    # -------------------------------------------------------------
    # Step 5: Export Master Deliverables
    # -------------------------------------------------------------
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    out_prefix = f"Run_{data_source}_{timestamp}"
    excel_path = os.path.join('Docs', f"{out_prefix}_Report.xlsx")
    
    log(f"Exporting Excel deliverable to '{excel_path}'...", 0.95)
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df_es.to_excel(writer, sheet_name='ES_Modules', index=False)
        df_matrix.to_excel(writer, sheet_name='Allocation_Matrix', index=False)
        df_r2.to_excel(writer, sheet_name='R2_Reconstruction', index=False)

    summary = {
        'data_source': data_source,
        'total_rois': len(roi_info),
        'total_g': total_g,
        'es_count': len(selected),
        'g_covered': len(covered),
        'coverage_pct': round(cov_pct, 2),
        'variance_explained_pct': round(tot_var_pct, 2),
        'avg_r2_full': round(avg_r2_full, 4),
        'avg_r2_optimal': round(avg_r2_opt, 4),
        'avg_delta_r2': round(avg_delta_r2, 4),
        'excel_path': excel_path,
        'duration_sec': round(time.time() - t0, 2)
    }

    log(f"Pipeline Complete in {summary['duration_sec']}s! All deliverables saved.", 1.00)
    return summary, df_es, df_matrix, df_r2
