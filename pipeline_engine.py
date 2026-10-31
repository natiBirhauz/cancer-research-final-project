"""
pipeline_engine.py
------------------
Unified execution engine for the Cancer Research TME Surprisal Analysis & 
Exact Sign Sequence (ES) Pipeline.

Processes user-provided datasets without preloaded data:
- Uploaded Excel file (.xlsx) with 'Gvalues' and 'Lambavalues' sheets
- Multiple uploaded Excel files (one per sample/ROI)
- Uploaded ZIP archive containing dataset folders
- User-specified local directory path

Computes:
1. Surprisal Analysis matrix ingestion (G basis, L amplitudes)
2. Kinetic variance energy weights w(g)
3. Combinatorial canonical Exact Sign Sequence (ES) mining
4. Greedy Set Cover optimization under selectable anchor strategies
5. Conflict-free disjoint allocation matrix (zero duplicate allocations)
6. Cell-level R2 and compression penalty Delta R2 reconstruction
"""

import os
import re
import glob
import time
import zipfile
import tempfile
import shutil
from collections import defaultdict
from itertools import combinations
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def canon_dict(d):
    """Canonical form: alphabetical order; invert signs if first protein sign is negative."""
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
    """
    Checks whether an ES matches a G-vector with zero contradictions.
    Returns: sign ('+' or '-'), overlap_count, sorted_common_proteins.
    """
    common = set(es_dict.keys()) & g_prots
    if len(common) < min_overlap:
        return None, 0, []
    if all(es_dict[p] == g_signs[p] for p in common):
        return '+', len(common), sorted(common)
    if all(es_dict[p] == -g_signs[p] for p in common):
        return '-', len(common), sorted(common)
    return None, 0, []

def compute_r2_vectorized(X_log, G, L, ss_tot_threshold=1e-3):
    """
    Computes cell-level R2 across proteins.
    Excludes cells with near-zero total sum of squares (constant baseline cells).
    """
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

def parse_excel_sample(file_path, sample_name, threshold=0.05, min_overlap=2):
    """
    Parses a single Excel file containing either:
    1. 'Gvalues' and 'Lambavalues' sheets
    2. A tail_proteins sheet (columns with G*+ and G*-)
    """
    xl = pd.ExcelFile(file_path)
    sheet_names_lower = [s.lower() for s in xl.sheet_names]

    # Format 1: Gvalues & Lambavalues
    has_g = any('gvalue' in s or s == 'g' for s in sheet_names_lower)
    has_l = any('lambavalue' in s or s == 'l' or 'lambda' in s for s in sheet_names_lower)

    if has_g and has_l:
        g_sheet = [s for s in xl.sheet_names if 'gvalue' in s.lower() or s.lower() == 'g'][0]
        l_sheet = [s for s in xl.sheet_names if 'lambavalue' in s.lower() or s.lower() == 'l' or 'lambda' in s.lower()][0]

        raw_g = xl.parse(g_sheet, header=None)
        h_idx = None
        for r in range(min(10, raw_g.shape[0])):
            row = [str(x).strip() for x in raw_g.iloc[r].tolist()]
            if any(re.match(r'^G\d+$', x) for x in row):
                h_idx = r
                break
        if h_idx is None:
            h_idx = 0

        # Find protein marker column
        prot_col = None
        for c in range(raw_g.shape[1]):
            col_vals = [str(x).strip().upper() for x in raw_g.iloc[h_idx+1:, c].dropna()]
            if any(len(v) > 1 and not v.replace('.', '', 1).replace('-', '', 1).isdigit() for v in col_vals):
                prot_col = c
                break
        if prot_col is None:
            prot_col = 0

        df_g_data = raw_g.iloc[h_idx+1:, :].copy()
        prots = df_g_data[prot_col].astype(str).str.strip().str.upper().tolist()

        g_cols = {}
        for c in range(raw_g.shape[1]):
            hdr = str(raw_g.iloc[h_idx, c]).strip()
            if re.match(r'^G\d+$', hdr):
                g_cols[hdr] = c

        n_comps = len(g_cols)
        G_full = np.zeros((len(prots), n_comps))
        for i in range(n_comps):
            g_key = f"G{i}"
            if g_key in g_cols:
                G_full[:, i] = df_g_data[g_cols[g_key]].values.astype(float)

        raw_l = xl.parse(l_sheet)
        l_cols = [c for c in raw_l.columns if re.match(r'^L\d+$', str(c).strip())]
        l_cols = sorted(l_cols, key=lambda x: int(re.search(r'\d+', str(x)).group(0)))
        L_full = raw_l[l_cols].values.astype(float)

        X_log = G_full @ L_full.T
        r2_full = compute_r2_vectorized(X_log, G_full, L_full)

        l_energies = np.sum(L_full ** 2, axis=0)
        dev_energy = float(np.sum(l_energies[1:])) if len(l_energies) > 1 else 1.0

        sample_g_vectors = []
        for g_num in range(1, n_comps):
            vals = G_full[:, g_num]
            signs = {}
            for p_name, val in zip(prots, vals):
                if abs(val) >= threshold:
                    signs[p_name] = 1 if val > 0 else -1
            w = float(l_energies[g_num] / dev_energy) if dev_energy > 0 else 0.0
            sample_g_vectors.append({
                'gkey': f"{sample_name}_G{g_num}",
                'roi': sample_name,
                'g_num': g_num,
                'signs': signs,
                'prots': set(signs.keys()),
                'count': len(signs),
                'weight': w
            })

        sample_info = {
            'roi_id': sample_name,
            'proteins': prots,
            'n_cells': L_full.shape[0],
            'G_all': G_full,
            'L_all': L_full,
            'X_log': X_log,
            'r2_full': r2_full,
            'l_energies': l_energies,
            'dev_energy': dev_energy
        }
        return sample_info, sample_g_vectors

    # Format 2: Tail proteins format
    df_tail = xl.parse(0)
    g_cols = [c for c in df_tail.columns if re.match(r'^G\d+[+-]$', str(c).strip())]
    g_nums = sorted(set(int(re.search(r'G(\d+)', c).group(1)) for c in g_cols))

    sample_g_vectors = []
    all_prots = set()
    for n in g_nums:
        signs = {}
        for s, col in [('+', f'G{n}+'), ('-', f'G{n}-')]:
            if col in df_tail.columns:
                for p in df_tail[col].dropna():
                    p = str(p).strip().upper()
                    if p and p != 'NAN':
                        signs[p] = 1 if s == '+' else -1
                        all_prots.add(p)
        if len(signs) >= min_overlap:
            sample_g_vectors.append({
                'gkey': f"{sample_name}_G{n}",
                'roi': sample_name,
                'g_num': n,
                'signs': signs,
                'prots': set(signs.keys()),
                'count': len(signs),
                'weight': 1.0 / len(g_nums)
            })

    sample_info = {
        'roi_id': sample_name,
        'proteins': sorted(list(all_prots)),
        'n_cells': 0,
        'G_all': None,
        'L_all': None,
        'X_log': None,
        'r2_full': 1.0,
        'l_energies': None,
        'dev_energy': 1.0
    }
    return sample_info, sample_g_vectors

def ingest_user_data(file_paths, folder_path=None, threshold=0.05, min_overlap=2):
    """
    Ingests user-supplied files, archives, or directories.
    """
    all_g_data = []
    all_roi_info = {}
    temp_dirs_to_clean = []

    files_to_process = []
    if file_paths:
        if isinstance(file_paths, str):
            file_paths = [file_paths]
        for fp in file_paths:
            if fp.lower().endswith('.zip'):
                td = tempfile.mkdtemp(prefix='cancer_es_upload_')
                temp_dirs_to_clean.append(td)
                with zipfile.ZipFile(fp, 'r') as zf:
                    zf.extractall(td)
                for root, _, files in os.walk(td):
                    for f in files:
                        if f.lower().endswith(('.xlsx', '.xls')):
                            files_to_process.append(os.path.join(root, f))
            elif fp.lower().endswith(('.xlsx', '.xls')):
                files_to_process.append(fp)

    if folder_path and os.path.isdir(folder_path):
        for root, _, files in os.walk(folder_path):
            for f in files:
                if f.lower().endswith(('.xlsx', '.xls')):
                    files_to_process.append(os.path.join(root, f))

    if not files_to_process:
        raise ValueError("No valid Excel spreadsheet files (.xlsx / .xls) found in the user input.")

    for fpath in files_to_process:
        parent_dir = os.path.basename(os.path.dirname(fpath))
        base_name = os.path.splitext(os.path.basename(fpath))[0]
        sample_name = parent_dir if (base_name.lower() in ('data', 'dataset') and parent_dir) else base_name

        try:
            s_info, s_g_vecs = parse_excel_sample(fpath, sample_name, threshold=threshold, min_overlap=min_overlap)
            if s_g_vecs:
                all_roi_info[sample_name] = s_info
                for g in s_g_vecs:
                    g['g_idx'] = len(all_g_data)
                    all_g_data.append(g)
        except Exception as e:
            print(f"Warning: Failed parsing file {fpath}: {e}")

    for td in temp_dirs_to_clean:
        try:
            shutil.rmtree(td, ignore_errors=True)
        except Exception:
            pass

    if not all_g_data:
        raise ValueError("Could not extract any valid biological G-vectors from the provided files. Please ensure the Excel sheets conform to the documented structure.")

    return all_roi_info, all_g_data

def run_user_pipeline(
    file_paths=None,
    folder_path=None,
    min_es_length=3,
    max_es_length=8,
    threshold=0.05,
    min_overlap=2,
    min_g_recurrence=2,
    strategy='Hybrid',
    progress_callback=None
):
    """
    Executes the complete pipeline strictly on user-supplied data.
    """
    def log(msg, pct=None):
        if progress_callback:
            progress_callback(pct, msg)
        print(f"[{time.strftime('%H:%M:%S')}] {msg}")

    t0 = time.time()
    log("Ingesting user-supplied dataset...", 0.10)

    roi_info, g_data = ingest_user_data(
        file_paths=file_paths,
        folder_path=folder_path,
        threshold=threshold,
        min_overlap=min_overlap
    )

    total_g = len(g_data)
    total_rois = len(roi_info)
    log(f"Successfully loaded {total_g:,} G-vectors across {total_rois} sample(s).", 0.30)

    # Combinatorial mining
    log(f"Mining canonical Exact Sign Sequences (lengths {min_es_length} to {max_es_length})...", 0.45)
    combo_candidates = set()
    for g in g_data:
        items = list(g['signs'].items())
        n = len(items)
        for k in range(min_es_length, min(n + 1, max_es_length + 1)):
            for c in combinations(items, k):
                cd = canon_dict(dict(c))
                combo_candidates.add(tuple(sorted(cd.items())))

    log(f"Mined {len(combo_candidates):,} unique candidate sequences.", 0.60)

    # Greedy set cover
    log(f"Running Greedy Optimization with Strategy: '{strategy}'...", 0.70)
    covered = set()
    selected = []
    g_assignment = {}

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
                else:  # Hybrid Multi-Objective
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

    cov_pct = (len(covered) / total_g * 100.0) if total_g > 0 else 0.0
    tot_var_pct = (sum(e['gain'] for e in selected) / total_rois * 100.0) if total_rois > 0 else 0.0
    log(f"Greedy Cover Complete: {len(selected)} ES modules cover {len(covered):,}/{total_g:,} Gs ({cov_pct:.1f}%).", 0.85)

    # Conflict-Free Matrix & R2 Reconstruction
    log("Constructing conflict-free matrix and computing cell-level R2...", 0.90)
    matrix_rows = []
    r2_rows = []

    for r_id, info in roi_info.items():
        row_assign = {'Sample': r_id, 'Cells': info['n_cells']}
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

        # R2 Calculation if matrices are present
        if info['G_all'] is not None and info['L_all'] is not None and info['X_log'] is not None:
            selected_g_nums = sorted(set(selected_g_nums))
            G_opt = info['G_all'][:, selected_g_nums]
            L_opt = info['L_all'][:, selected_g_nums]
            r2_opt = compute_r2_vectorized(info['X_log'], G_opt, L_opt)
            delta_r2 = info['r2_full'] - r2_opt
            status = 'High Fidelity (<0.05)' if delta_r2 < 0.05 else ('Moderate Loss (0.05-0.15)' if delta_r2 <= 0.15 else 'Noticeable Loss (>0.15)')
            r2_full_val = info['r2_full']
        else:
            r2_opt = 1.0
            delta_r2 = 0.0
            r2_full_val = 1.0
            status = 'Tail Table (Unspecified)'

        r2_rows.append({
            'Sample': r_id,
            'Cells': info['n_cells'],
            'Selected_G_Count': max(0, len(selected_g_nums) - 1),
            'Selected_Gs': ", ".join(f"G{n}" for n in selected_g_nums if n > 0),
            'R2_full': r2_full_val,
            'R2_optimal': r2_opt,
            'Delta_R2': delta_r2,
            'Status': status
        })

    df_matrix = pd.DataFrame(matrix_rows)
    df_r2 = pd.DataFrame(r2_rows)

    # Format ES table
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

    # Export master deliverable Excel
    os.makedirs('Docs', exist_ok=True)
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    excel_path = os.path.join('Docs', f"User_Run_{timestamp}_Report.xlsx")

    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df_es.to_excel(writer, sheet_name='Selected_ES_Modules', index=False)
        df_matrix.to_excel(writer, sheet_name='Allocation_Matrix', index=False)
        df_r2.to_excel(writer, sheet_name='R2_Reconstruction', index=False)

    summary = {
        'total_samples': total_rois,
        'total_g': total_g,
        'es_count': len(selected),
        'g_covered': len(covered),
        'coverage_pct': round(cov_pct, 2),
        'variance_explained_pct': round(tot_var_pct, 2),
        'avg_r2_optimal': round(float(df_r2['R2_optimal'].mean()), 4),
        'avg_delta_r2': round(float(df_r2['Delta_R2'].mean()), 4),
        'excel_path': excel_path,
        'duration_sec': round(time.time() - t0, 2)
    }

    log(f"Execution finished successfully in {summary['duration_sec']}s! Report generated.", 1.00)
    return summary, df_es, df_matrix, df_r2
