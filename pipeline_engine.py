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

Generates the 4 Master Deliverable Files:
1. Optimal_ES_Result.xlsx
2. Optimal_ES_Result_Signs_R2.xlsx
3. ES_Combinations.csv
4. Anchor_Strategies_Comparison.xlsx
Plus Deliverables_Bundle.zip containing all 4 files.
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
    1. 'Gvalues' and 'Lambavalues' sheets (Surprisal Analysis Matrix)
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
                'weight': 1.0 / len(g_nums) if g_nums else 0.0
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

def ingest_user_data(file_paths=None, folder_path=None, threshold=0.05, min_overlap=2):
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
                    for f in sorted(files):
                        if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$'):
                            files_to_process.append(os.path.join(root, f))
            elif fp.lower().endswith(('.xlsx', '.xls')) and not os.path.basename(fp).startswith('~$'):
                files_to_process.append(fp)

    if folder_path and os.path.isdir(folder_path):
        for root, _, files in os.walk(folder_path):
            for f in sorted(files):
                if f.lower().endswith(('.xlsx', '.xls')) and not f.startswith('~$'):
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
    output_dir='output',
    progress_callback=None
):
    """
    Executes the complete pipeline on user-supplied data and produces the 4 master deliverables:
    1. Optimal_ES_Result.xlsx
    2. Optimal_ES_Result_Signs_R2.xlsx
    3. ES_Combinations.csv
    4. Anchor_Strategies_Comparison.xlsx
    Plus Deliverables_Bundle.zip.
    """
    def log(msg, pct=None):
        if progress_callback:
            progress_callback(pct, msg)
        print(f"[{time.strftime('%H:%M:%S')}] {msg}")

    t0 = time.time()
    os.makedirs(output_dir, exist_ok=True)
    log("Ingesting user-supplied dataset...", 0.05)

    roi_info, g_data = ingest_user_data(
        file_paths=file_paths,
        folder_path=folder_path,
        threshold=threshold,
        min_overlap=min_overlap
    )

    total_g = len(g_data)
    total_rois = len(roi_info)
    log(f"Successfully loaded {total_g:,} G-vectors across {total_rois} sample(s).", 0.15)

    # Combinatorial mining
    log(f"Mining canonical Exact Sign Sequences (lengths {min_es_length} to {max_es_length})...", 0.25)
    seen_combos = defaultdict(list)
    for g in g_data:
        items = list(g['signs'].items())
        n = len(items)
        for k in range(min_es_length, min(n + 1, max_es_length + 1)):
            for c in combinations(items, k):
                cd = canon_dict(dict(c))
                ct = tuple(sorted(cd.items()))
                orient, _, _ = match_es_g(cd, g['signs'], g['prots'], min_overlap=min_overlap)
                if orient is not None:
                    seen_combos[ct].append(g['gkey'])

    total_mined = len(seen_combos)
    log(f"Mined {total_mined:,} unique candidate signatures.", 0.40)

    # 1. Export ES_Combinations.csv
    combo_records = []
    for ct, gkeys in seen_combos.items():
        unique_gkeys = sorted(set(gkeys))
        cd = dict(ct)
        combo_records.append({
            'Length': len(cd),
            'Frequency_Count': len(unique_gkeys),
            'In_2+_Gs': len(unique_gkeys) >= 2,
            'Exact_Sign_Sequence': tuple_to_str(cd),
            'Covered_G_Vectors': "; ".join(unique_gkeys)
        })

    df_combos = pd.DataFrame(combo_records)
    df_combos.sort_values(by=['In_2+_Gs', 'Frequency_Count', 'Length'], ascending=[False, False, False], inplace=True)
    es_combinations_csv_path = os.path.join(output_dir, "ES_Combinations.csv")
    df_combos.to_csv(es_combinations_csv_path, index=False)
    log("Generated deliverable: ES_Combinations.csv", 0.45)

    # Select candidate modules for Set Cover optimization
    log("Selecting candidate modules for Set Cover optimization...", 0.50)
    cands_sorted = sorted(seen_combos.items(), key=lambda x: len(set(x[1])), reverse=True)
    if len(cands_sorted) > 4000:
        top_cands = [ct for ct, glist in cands_sorted if len(set(glist)) >= min_g_recurrence][:3500]
        # Include full G-vector signatures
        seen_ct = set(top_cands)
        for g in g_data:
            if len(g['signs']) >= min_es_length:
                ct = tuple(sorted(canon_dict(g['signs']).items()))
                if ct not in seen_ct:
                    seen_ct.add(ct)
                    top_cands.append(ct)
    else:
        top_cands = [ct for ct, glist in cands_sorted if len(set(glist)) >= min_g_recurrence]

    eligible_modules = []
    for ct in top_cands:
        mod_dict = dict(ct)
        hits = {}
        tot_var = 0.0
        for g in g_data:
            orient, n_common, matched = match_es_g(mod_dict, g['signs'], g['prots'], min_overlap=min_overlap)
            if orient is not None:
                hits[g['g_idx']] = (orient, n_common, matched)
                tot_var += g['weight']
        if len(hits) >= min_g_recurrence:
            eligible_modules.append({
                'ct': ct,
                'mod': mod_dict,
                'len': len(mod_dict),
                'hits': hits,
                'initial_count': len(hits),
                'initial_var': tot_var
            })

    # 2. Evaluate Anchor Strategies
    log("Evaluating Anchor Strategies & Set Cover models...", 0.55)
    def evaluate_strategy(name, seed_filter, scoring_func):
        t_strat = time.time()
        covered = set()
        selected = []
        g_assignment = {}

        if seed_filter:
            cands = [m for m in eligible_modules if seed_filter(m)]
            if cands:
                best_seed = max(cands, key=lambda m: scoring_func(m, set(m['hits'].keys())))
                best_new = set(best_seed['hits'].keys())
                covered.update(best_new)
                lbl = f"ES_{len(selected)+1}"
                for gi in best_new:
                    g_assignment[gi] = (lbl, best_seed['hits'][gi])
                selected.append({
                    'lbl': lbl,
                    'mod': best_seed['mod'],
                    'new_g': best_new,
                    'len': best_seed['len'],
                    'var_gain': sum(g_data[gi]['weight'] for gi in best_new)
                })

        while len(covered) < total_g:
            best_mod = None
            best_new = set()
            best_score = -1e9

            for m in eligible_modules:
                new_g = set(m['hits'].keys()) - covered
                if len(new_g) < min_g_recurrence:
                    continue
                score = scoring_func(m, new_g)
                if score > best_score:
                    best_score = score
                    best_new = new_g
                    best_mod = m

            if not best_mod or len(best_new) < min_g_recurrence:
                break

            covered.update(best_new)
            lbl = f"ES_{len(selected)+1}"
            for gi in best_new:
                g_assignment[gi] = (lbl, best_mod['hits'][gi])
            selected.append({
                'lbl': lbl,
                'mod': best_mod['mod'],
                'new_g': best_new,
                'len': best_mod['len'],
                'var_gain': sum(g_data[gi]['weight'] for gi in best_new)
            })

        # Compute R2 per sample
        roi_g_assigned = defaultdict(set)
        for gi in covered:
            g = g_data[gi]
            roi_g_assigned[g['roi']].add(g['g_num'])

        r2_diffs, r2_opts, r2_fulls = [], [], []
        for r_id, rdata in roi_info.items():
            if rdata['G_all'] is not None and rdata['L_all'] is not None and rdata['X_log'] is not None:
                X_log = rdata['X_log']
                G_all = rdata['G_all']
                L_all = rdata['L_all']
                r2_f = rdata['r2_full']
                sel_nums = sorted(roi_g_assigned.get(r_id, set()))
                sel_cols = [0] + sel_nums
                r2_o = compute_r2_vectorized(X_log, G_all[:, sel_cols], L_all[:, sel_cols])
                r2_fulls.append(r2_f)
                r2_opts.append(r2_o)
                r2_diffs.append(r2_f - r2_o)
            else:
                r2_fulls.append(1.0)
                r2_opts.append(1.0)
                r2_diffs.append(0.0)

        tot_var = sum(g_data[gi]['weight'] for gi in covered)

        return {
            'Anchor_Strategy': name,
            'ES_Count': len(selected),
            'Covered_Gs': len(covered),
            'Coverage_%': round(len(covered) / total_g * 100, 2) if total_g > 0 else 0,
            'Variance_%': round(tot_var / total_rois * 100, 2) if total_rois > 0 else 0,
            'Mean_R2_full': round(float(np.mean(r2_fulls)), 4),
            'Mean_R2_opt': round(float(np.mean(r2_opts)), 4),
            'Mean_Delta_R2': round(float(np.mean(r2_diffs)), 4),
            'Excellent_%': round(float(np.mean(np.array(r2_diffs) < 0.05) * 100), 1),
            'Acceptable_%': round(float(np.mean((np.array(r2_diffs) >= 0.05) & (np.array(r2_diffs) <= 0.15)) * 100), 1),
            'Alert_%': round(float(np.mean(np.array(r2_diffs) > 0.15) * 100), 1),
            'Exec_Time_s': round(time.time() - t_strat, 2),
            'selected': selected,
            'g_assignment': g_assignment,
            'r2_per_sample': {r_id: (r2_fulls[i], r2_opts[i], r2_diffs[i], sorted(roi_g_assigned.get(r_id, set()))) for i, r_id in enumerate(roi_info.keys())}
        }

    # Evaluate strategies
    strategies_eval = []
    # 1. Pure SVD Variance Driver
    strategies_eval.append(evaluate_strategy(
        "1. SVD Variance Driver", None,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g)
    ))
    # 2. Max Frequency / Universality
    strategies_eval.append(evaluate_strategy(
        "2. Max Frequency / Universality", None,
        lambda m, new_g: float(len(new_g))
    ))
    # 3. Deep Specificity
    strategies_eval.append(evaluate_strategy(
        "3. Deep Specificity", None,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g) * (1.5 if m['len'] >= 5 else 1.0)
    ))
    # 4. Biological Hallmark / Seed
    pair_counts = defaultdict(int)
    for g in g_data:
        for p1, p2 in combinations(sorted(g['prots']), 2):
            if g['signs'][p1] == g['signs'][p2]:
                pair_counts[(p1, p2)] += 1
    top_pair = max(pair_counts.items(), key=lambda x: x[1])[0] if pair_counts else None
    seed_func = (lambda m: top_pair[0] in m['mod'] and top_pair[1] in m['mod']) if top_pair else None
    strategies_eval.append(evaluate_strategy(
        f"4. Biological Hallmark ({top_pair[0] if top_pair else 'Seed'} Co-expression)", seed_func,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g) * np.sqrt(len(new_g))
    ))
    # 5. Hybrid Multi-Objective (Variance x Count)
    hybrid_res = evaluate_strategy(
        "5. Hybrid Multi-Objective (Variance x Count)", None,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g) * np.sqrt(len(new_g))
    )
    strategies_eval.append(hybrid_res)

    # Select target strategy based on user preference
    target_strat = hybrid_res
    if 'SVD' in strategy:
        target_strat = strategies_eval[0]
    elif 'Frequency' in strategy:
        target_strat = strategies_eval[1]
    elif 'Deep' in strategy:
        target_strat = strategies_eval[2]
    elif 'Bio' in strategy:
        target_strat = strategies_eval[3]

    # Export Anchor_Strategies_Comparison.xlsx
    anchor_xlsx_path = os.path.join(output_dir, "Anchor_Strategies_Comparison.xlsx")
    wb_anc = openpyxl.Workbook()
    ws_ov = wb_anc.active
    ws_ov.title = "Anchor_Strategies_Overview"
    headers_ov = ['Anchor_Strategy', 'ES_Count', 'Covered_Gs', 'Coverage_%', 'Variance_%', 'Mean_R2_full', 'Mean_R2_opt', 'Mean_Delta_R2', 'Excellent_%', 'Acceptable_%', 'Alert_%', 'Exec_Time_s']
    ws_ov.append(headers_ov)
    for s in strategies_eval:
        ws_ov.append([
            s['Anchor_Strategy'], s['ES_Count'], s['Covered_Gs'], s['Coverage_%'],
            s['Variance_%'], s['Mean_R2_full'], s['Mean_R2_opt'], s['Mean_Delta_R2'],
            s['Excellent_%'], s['Acceptable_%'], s['Alert_%'], s['Exec_Time_s']
        ])

    # Add detailed breakdown sheets for ALL evaluated strategies
    strat_sheet_titles = [
        "Strat_1_SVD_Variance",
        "Strat_2_Max_Frequency",
        "Strat_3_Deep_Specificity",
        "Strat_4_Biological_Seed",
        "Strat_5_Hybrid_MultiObjective"
    ]
    headers_es = ['Rank', 'ES_Label', 'Sequence_Length', 'New_G_Covered', 'Direct_Matches_(+)', 'Inverse_Matches_(-)', 'Cumulative_G_Covered', 'Coverage_%', 'New_Variance_Weight', 'Cumulative_Variance_%', 'ROIs_Covered_Count', 'Exact_Sign_Sequence']

    for idx, strat in enumerate(strategies_eval):
        title = strat_sheet_titles[idx] if idx < len(strat_sheet_titles) else f"Strat_{idx+1}_Detail"
        ws_strat = wb_anc.create_sheet(title=title)
        ws_strat.append(headers_es)
        cum_g = 0
        cum_var = 0.0
        for rk, es in enumerate(strat['selected'], 1):
            cum_g += len(es['new_g'])
            cum_var += es['var_gain']
            d_cnt = sum(1 for gi in es['new_g'] if strat['g_assignment'][gi][1][0] == '+')
            i_cnt = sum(1 for gi in es['new_g'] if strat['g_assignment'][gi][1][0] == '-')
            rois_cnt = len(set(g_data[gi]['roi'] for gi in es['new_g']))
            ws_strat.append([
                rk, es['lbl'], es['len'], len(es['new_g']), d_cnt, i_cnt,
                cum_g, round(cum_g / total_g * 100, 2) if total_g > 0 else 0,
                round(es['var_gain'], 4), round(cum_var / total_rois * 100, 2) if total_rois > 0 else 0,
                rois_cnt, tuple_to_str(es['mod'])
            ])

    wb_anc.save(anchor_xlsx_path)
    log("Generated deliverable: Anchor_Strategies_Comparison.xlsx", 0.70)

    # 3 & 4: Build Optimal_ES_Result_Signs_R2.xlsx and Optimal_ES_Result.xlsx
    log("Constructing Master Excel Deliverables & Reconstruction Matrices...", 0.80)
    def build_master_workbook(file_name):
        wb = openpyxl.Workbook()
        # Sheet 1: Summary
        ws_sum = wb.active
        ws_sum.title = 'Summary'
        ws_sum.append(['OPTIMAL EXACT SIGN SEQUENCES (ES) & R2 VARIANCE REPORT'])
        ws_sum.append(['Surprisal Analysis: Exact Sign Sequences, G Orientations & Reconstruction Quality'])
        ws_sum.append([])
        ws_sum.append(['Metric', 'Value'])
        ws_sum.append(['Primary Strategy Applied', target_strat['Anchor_Strategy']])
        ws_sum.append(['Total Input Samples / ROIs', total_rois])
        ws_sum.append(['Total Cellular Microenvironment Vectors (G)', total_g])
        ws_sum.append(['Number of Optimal ES Modules', target_strat['ES_Count']])
        ws_sum.append(['Total G-Vectors Covered', target_strat['Covered_Gs']])
        ws_sum.append(['Universe G Coverage %', f"{target_strat['Coverage_%']}%"])
        ws_sum.append(['Total Kinetic Variance Explained', f"{target_strat['Variance_%']}%"])
        ws_sum.append(['Mean R2 Full Reconstruction', target_strat['Mean_R2_full']])
        ws_sum.append(['Mean R2 Optimal Reconstruction', target_strat['Mean_R2_opt']])
        ws_sum.append(['Mean Delta R2 Compression Penalty', target_strat['Mean_Delta_R2']])
        ws_sum.append(['High Fidelity Samples (<0.05 Delta R2)', f"{target_strat['Excellent_%']}%"])
        ws_sum.append(['Acceptable Fidelity Samples (0.05-0.15)', f"{target_strat['Acceptable_%']}%"])
        ws_sum.append(['Noticeable Loss Samples (>0.15)', f"{target_strat['Alert_%']}%"])

        # Sheet 2: Optimal_ES_SetCover
        ws_opt = wb.create_sheet(title='Optimal_ES_SetCover')
        ws_opt.append(headers_es)
        cum_g_local = 0
        cum_var_local = 0.0
        for rk, es in enumerate(target_strat['selected'], 1):
            cum_g_local += len(es['new_g'])
            cum_var_local += es['var_gain']
            d_cnt = sum(1 for gi in es['new_g'] if target_strat['g_assignment'][gi][1][0] == '+')
            i_cnt = sum(1 for gi in es['new_g'] if target_strat['g_assignment'][gi][1][0] == '-')
            rois_cnt = len(set(g_data[gi]['roi'] for gi in es['new_g']))
            ws_opt.append([
                rk, es['lbl'], es['len'], len(es['new_g']), d_cnt, i_cnt,
                cum_g_local, round(cum_g_local / total_g * 100, 2) if total_g > 0 else 0,
                round(es['var_gain'], 4), round(cum_var_local / total_rois * 100, 2) if total_rois > 0 else 0,
                rois_cnt, tuple_to_str(es['mod'])
            ])

        # Sheet 3: ES_Conflict_Free_Matrix
        ws_mat = wb.create_sheet(title='ES_Conflict_Free_Matrix')
        mat_cols = ['Sample', 'Cells', 'Total_G'] + [es['lbl'] for es in target_strat['selected']]
        ws_mat.append(mat_cols)
        for r_id, s in roi_info.items():
            row = [r_id, s['n_cells'], len([g for g in g_data if g['roi'] == r_id])]
            used_g = set()
            for es in target_strat['selected']:
                lbl = es['lbl']
                cands = [gi for gi in es['new_g'] if g_data[gi]['roi'] == r_id and g_data[gi]['g_num'] not in used_g]
                if cands:
                    best_gi = max(cands, key=lambda gi: g_data[gi]['weight'])
                    g_item = g_data[best_gi]
                    orient = target_strat['g_assignment'][best_gi][1][0]
                    row.append(f"G{g_item['g_num']}({orient})")
                    used_g.add(g_item['g_num'])
                else:
                    row.append("")
            ws_mat.append(row)

        # Sheet 4: Best_Combination_Details
        ws_det = wb.create_sheet(title='Best_Combination_Details')
        ws_det.append(['Sample', 'Cells', 'Total_G_In_Sample', 'Covered_G_Count', 'Sample_Coverage_%', 'Assigned_ES_Labels', 'Consolidated_G_Assignments'])
        for r_id, s in roi_info.items():
            sample_gs = [g for g in g_data if g['roi'] == r_id]
            cov_gs = [g for g in sample_gs if g['g_idx'] in target_strat['g_assignment']]
            assigned_es = sorted(set(target_strat['g_assignment'][g['g_idx']][0] for g in cov_gs))
            assignments_str = []
            for g in sorted(cov_gs, key=lambda x: x['g_num']):
                lbl, hit = target_strat['g_assignment'][g['g_idx']]
                orient = hit[0]
                pct = (hit[1] / g['count'] * 100) if g['count'] > 0 else 0
                assignments_str.append(f"G{g['g_num']}:{lbl}({orient})[{pct:.1f}%]")
            ws_det.append([
                r_id, s['n_cells'], len(sample_gs), len(cov_gs),
                round(len(cov_gs) / len(sample_gs) * 100, 2) if sample_gs else 0,
                "; ".join(assigned_es), ", ".join(assignments_str)
            ])

        # Sheet 5: Per_G_Coverage_Detail
        ws_perg = wb.create_sheet(title='Per_G_Coverage_Detail')
        ws_perg.append(['Sample', 'G_Key', 'G_Num', 'Total_Prots_In_G', 'Assigned_ES', 'Orientation', 'Matched_Prots_Count', 'Overlap_%', 'Variance_Weight', 'Status'])
        for g in g_data:
            if g['g_idx'] in target_strat['g_assignment']:
                lbl, hit = target_strat['g_assignment'][g['g_idx']]
                orient, n_m, _ = hit
                pct = round(n_m / g['count'] * 100, 1) if g['count'] > 0 else 0
                ws_perg.append([
                    g['roi'], g['gkey'], g['g_num'], g['count'], lbl, orient, n_m, pct, round(g['weight'], 4), 'Covered'
                ])
            else:
                ws_perg.append([
                    g['roi'], g['gkey'], g['g_num'], g['count'], 'UNASSIGNED', 'None', 0, 0.0, round(g['weight'], 4), 'Uncovered'
                ])

        # Sheet 6: R2_Comparison_Per_ROI
        ws_r2 = wb.create_sheet(title='R2_Comparison_Per_ROI')
        ws_r2.append(['Sample', 'N_Cells', 'R2_Full_All_Components', 'R2_Optimal_Selected_Components', 'Delta_R2_Loss', 'N_Components_Used', 'Selected_G_List', 'Interpretation'])
        for r_id, (r2_f, r2_o, delta, sel_gs) in target_strat['r2_per_sample'].items():
            s = roi_info[r_id]
            interp = 'High Fidelity (<0.05 loss)' if delta < 0.05 else ('Moderate Loss (0.05-0.15)' if delta <= 0.15 else 'Noticeable Loss (>0.15)')
            ws_r2.append([
                r_id, s['n_cells'], round(r2_f, 4), round(r2_o, 4), round(delta, 4),
                len(sel_gs), ", ".join(f"G{n}" for n in sel_gs), interp
            ])

        # Sheet 7: Per_ROI_Coverage
        ws_rois = wb.create_sheet(title='Per_ROI_Coverage')
        ws_rois.append(['Sample', 'Cells', 'Total_G', 'Covered_G', 'Coverage_%', 'Variance_Captured_%'])
        for r_id, s in roi_info.items():
            sample_gs = [g for g in g_data if g['roi'] == r_id]
            cov_sample_gs = [g for g in sample_gs if g['g_idx'] in target_strat['g_assignment']]
            var_cap = sum(g['weight'] for g in cov_sample_gs)
            ws_rois.append([
                r_id, s['n_cells'], len(sample_gs), len(cov_sample_gs),
                round(len(cov_sample_gs) / len(sample_gs) * 100, 2) if sample_gs else 0,
                round(var_cap * 100, 2)
            ])

        # Sheet 8: Coverage_Curve
        ws_curv = wb.create_sheet(title='Coverage_Curve')
        ws_curv.append(['Step', 'ES_Label', 'Cumulative_G_Covered', 'Cumulative_Coverage_%', 'Cumulative_Variance_%'])
        cum_g_local = 0
        cum_var_local = 0.0
        for rk, es in enumerate(target_strat['selected'], 1):
            cum_g_local += len(es['new_g'])
            cum_var_local += es['var_gain']
            ws_curv.append([
                rk, es['lbl'], cum_g_local,
                round(cum_g_local / total_g * 100, 2) if total_g > 0 else 0,
                round(cum_var_local / total_rois * 100, 2) if total_rois > 0 else 0
            ])

        p = os.path.join(output_dir, file_name)
        wb.save(p)
        return p

    opt_signs_r2_path = build_master_workbook("Optimal_ES_Result_Signs_R2.xlsx")
    opt_es_result_path = build_master_workbook("Optimal_ES_Result.xlsx")
    log("Generated deliverables: Optimal_ES_Result.xlsx & Optimal_ES_Result_Signs_R2.xlsx", 0.90)

    # 5. Create ZIP Bundle containing all 4 deliverables
    zip_bundle_path = os.path.join(output_dir, "Deliverables_Bundle.zip")
    with zipfile.ZipFile(zip_bundle_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(opt_es_result_path, arcname="Optimal_ES_Result.xlsx")
        zf.write(opt_signs_r2_path, arcname="Optimal_ES_Result_Signs_R2.xlsx")
        zf.write(es_combinations_csv_path, arcname="ES_Combinations.csv")
        zf.write(anchor_xlsx_path, arcname="Anchor_Strategies_Comparison.xlsx")
    log("Created compressed deliverable bundle: Deliverables_Bundle.zip", 0.95)

    # Prepare DataFrames for GUI preview
    # 1. ES modules DataFrame
    es_preview_rows = []
    cum_g = 0
    cum_var = 0.0
    for rk, es in enumerate(target_strat['selected'], 1):
        cum_g += len(es['new_g'])
        cum_var += es['var_gain']
        d_cnt = sum(1 for gi in es['new_g'] if target_strat['g_assignment'][gi][1][0] == '+')
        i_cnt = sum(1 for gi in es['new_g'] if target_strat['g_assignment'][gi][1][0] == '-')
        es_preview_rows.append({
            'Rank': rk,
            'ES_ID': es['lbl'],
            'Length': es['len'],
            'New_Gs_Covered': len(es['new_g']),
            'Direct (+)': d_cnt,
            'Inverse (-)': i_cnt,
            'Cumulative_Gs': cum_g,
            'Coverage_%': f"{(cum_g / total_g * 100.0):.1f}%",
            'Exact_Sign_Sequence': tuple_to_str(es['mod'])
        })
    df_es = pd.DataFrame(es_preview_rows)

    # 2. Conflict-free matrix DataFrame
    mat_rows = []
    for r_id, s in roi_info.items():
        row_dict = {'Sample': r_id, 'Cells': s['n_cells']}
        used_g = set()
        for es in target_strat['selected']:
            lbl = es['lbl']
            cands = [gi for gi in es['new_g'] if g_data[gi]['roi'] == r_id and g_data[gi]['g_num'] not in used_g]
            if cands:
                best_gi = max(cands, key=lambda gi: g_data[gi]['weight'])
                g_item = g_data[best_gi]
                orient = target_strat['g_assignment'][best_gi][1][0]
                row_dict[lbl] = f"G{g_item['g_num']}({orient})"
                used_g.add(g_item['g_num'])
            else:
                row_dict[lbl] = ""
        mat_rows.append(row_dict)
    df_mat = pd.DataFrame(mat_rows)

    # 3. R2 per sample DataFrame
    r2_preview_rows = []
    for r_id, (r2_f, r2_o, delta, sel_gs) in target_strat['r2_per_sample'].items():
        interp = 'High Fidelity (<0.05 loss)' if delta < 0.05 else ('Moderate Loss (0.05-0.15)' if delta <= 0.15 else 'Noticeable Loss (>0.15)')
        r2_preview_rows.append({
            'Sample': r_id,
            'Cells': roi_info[r_id]['n_cells'],
            'Components_Used': len(sel_gs),
            'R2_full': round(r2_f, 4),
            'R2_optimal': round(r2_o, 4),
            'Delta_R2': round(delta, 4),
            'Fidelity_Tier': interp
        })
    df_r2 = pd.DataFrame(r2_preview_rows)

    # 4. Anchor strategies overview DataFrame
    anc_rows = []
    for s in strategies_eval:
        anc_rows.append({
            'Anchor_Strategy': s['Anchor_Strategy'],
            'ES_Count': s['ES_Count'],
            'Covered_Gs': f"{s['Covered_Gs']}/{total_g}",
            'Coverage_%': f"{s['Coverage_%']}%",
            'Variance_%': f"{s['Variance_%']}%",
            'Mean_R2_opt': s['Mean_R2_opt'],
            'Mean_Delta_R2': s['Mean_Delta_R2'],
            'High_Fidelity_%': f"{s['Excellent_%']}%"
        })
    df_anc = pd.DataFrame(anc_rows)

    summary = {
        'total_samples': total_rois,
        'total_g': total_g,
        'es_count': target_strat['ES_Count'],
        'g_covered': target_strat['Covered_Gs'],
        'coverage_pct': target_strat['Coverage_%'],
        'variance_explained_pct': target_strat['Variance_%'],
        'avg_r2_optimal': target_strat['Mean_R2_opt'],
        'avg_delta_r2': target_strat['Mean_Delta_R2'],
        'high_fidelity_pct': target_strat['Excellent_%'],
        'duration_sec': round(time.time() - t0, 2),
        'opt_es_result_path': opt_es_result_path,
        'opt_signs_r2_path': opt_signs_r2_path,
        'es_combinations_csv_path': es_combinations_csv_path,
        'anchor_xlsx_path': anchor_xlsx_path,
        'zip_bundle_path': zip_bundle_path
    }

    log(f"All 4 deliverables generated in {summary['duration_sec']}s!", 1.00)
    return summary, df_es, df_mat, df_r2, df_anc
