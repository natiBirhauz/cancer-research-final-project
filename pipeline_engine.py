"""
pipeline_engine.py
------------------
Unified execution engine for the Surprisal Analysis (SA) – Minimal Exact Sign
Signatures (ES) Discovery Platform.

Accepts any collection of datasets analyzed via Surprisal Analysis (SA) sharing
common proteins or genes:
- Uploaded Excel file(s) (.xlsx / .xls) with 'Gvalues' and 'Lambavalues' sheets
- Uploaded Excel file(s) with pre-extracted Tail tables (G1+, G1-, G2+, G2-, ...)
- Uploaded ZIP archive containing multiple dataset folders/files

Capabilities:
1. Pre-Flight Dataset Inspection (inspect_user_data):
   - Detects number of datasets, total G-vectors, total unique genes/proteins,
     shared genes/proteins across datasets, tail size distribution, and runtime estimate.
2. Dual-Mode Scalable ES Mining Engine:
   - Exact Combinatorial Subset Enumeration for compact panels (lengths 2..12).
   - High-Dimensional Sign-Consistent Pairwise/Cluster Intersection Mining for
     large omics/transcriptomic datasets (up to 20,000+ genes and sequence lengths 1,000+).
3. Multi-Objective Greedy Set Cover & Single-Cell/Sample R² Reconstruction.
4. Streamlined Deliverables & Visual HTML Summary Report:
   - Optimal_ES_Result.xlsx (5 focused, non-redundant sheets)
   - Analysis_Summary_Report.html (Standalone visual report with charts, heatmap & tables)
   - Anchor_Strategies_Comparison.xlsx (Overview benchmark of 5 greedy anchor strategies)
   - ES_Combinations.csv (Filtered recurring candidate combinations, 2+ Gs)
   - Deliverables_Bundle.zip (Consolidated archive)
"""

import os
import re
import glob
import time
import html
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
    """Canonical form: alphabetical order; invert signs if first feature sign is negative."""
    s = sorted(d.items(), key=lambda x: x[0])
    if not s:
        return {}
    if s[0][1] == -1:
        return dict((p, -v) for p, v in s)
    return dict(s)


def tuple_to_str(d_or_tuple, max_items=None):
    if isinstance(d_or_tuple, dict):
        items = sorted(d_or_tuple.items())
    else:
        items = sorted(d_or_tuple)
    if max_items and len(items) > max_items:
        shown = " | ".join(f"{p}({'+' if s == 1 else '-'})" for p, s in items[:max_items])
        return f"{shown} ... (+{len(items) - max_items} more)"
    return " | ".join(f"{p}({'+' if s == 1 else '-'})" for p, s in items)


def match_es_g(es_dict, g_signs, g_prots, min_overlap=2):
    """
    Checks whether an ES matches a G-vector with zero contradictions.
    Returns: sign ('+' or '-'), overlap_count, sorted_common_features.
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
    Computes cell/sample-level R2 across features (genes/proteins).
    Excludes columns with near-zero total sum of squares.
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

        # Find feature (protein/gene) marker column
        prot_col = None
        for c in range(raw_g.shape[1]):
            col_vals = [str(x).strip().upper() for x in raw_g.iloc[h_idx+1:, c].dropna()]
            if any(len(v) > 1 and not v.replace('.', '', 1).replace('-', '', 1).replace('E', '', 1).isdigit() for v in col_vals[:50]):
                prot_col = c
                break
        if prot_col is None:
            prot_col = 0

        df_g_data = raw_g.iloc[h_idx+1:, :].copy()
        prots = [p for p in df_g_data[prot_col].astype(str).str.strip().str.upper().tolist() if p and p != 'NAN']
        df_g_data = df_g_data.iloc[:len(prots), :]

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
                G_full[:, i] = pd.to_numeric(df_g_data[g_cols[g_key]], errors='coerce').fillna(0.0).values

        raw_l = xl.parse(l_sheet)
        l_cols = [c for c in raw_l.columns if re.match(r'^L\d+$', str(c).strip())]
        l_cols = sorted(l_cols, key=lambda x: int(re.search(r'\d+', str(x)).group(0)))
        L_full = raw_l[l_cols].apply(pd.to_numeric, errors='coerce').fillna(0.0).values

        X_log = G_full @ L_full.T
        r2_full = compute_r2_vectorized(X_log, G_full, L_full)

        l_energies = np.sum(L_full ** 2, axis=0)
        dev_energy = float(np.sum(l_energies[1:])) if len(l_energies) > 1 else 1.0

        sample_g_vectors = []
        for g_num in range(1, n_comps):
            vals = G_full[:, g_num]
            signs = {}
            magnitudes = {}
            for p_name, val in zip(prots, vals):
                if abs(val) >= threshold:
                    signs[p_name] = 1 if val > 0 else -1
                    magnitudes[p_name] = float(abs(val))
            w = float(l_energies[g_num] / dev_energy) if dev_energy > 0 else 0.0
            sample_g_vectors.append({
                'gkey': f"{sample_name}_G{g_num}",
                'roi': sample_name,
                'g_num': g_num,
                'signs': signs,
                'magnitudes': magnitudes,
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

    # Format 2: Tail proteins/genes format
    df_tail = xl.parse(0)
    g_cols = [c for c in df_tail.columns if re.match(r'^G\d+[+-]$', str(c).strip())]
    g_nums = sorted(set(int(re.search(r'G(\d+)', c).group(1)) for c in g_cols))

    sample_g_vectors = []
    all_prots = set()
    for n in g_nums:
        signs = {}
        magnitudes = {}
        for s, col in [('+', f'G{n}+'), ('-', f'G{n}-')]:
            if col in df_tail.columns:
                col_series = df_tail[col].dropna().tolist()
                total_in_col = max(1, len(col_series))
                for rank_idx, p in enumerate(col_series):
                    p = str(p).strip().upper()
                    if p and p != 'NAN':
                        signs[p] = 1 if s == '+' else -1
                        # Higher rank in tail table = slightly higher priority weight
                        magnitudes[p] = 1.0 + (total_in_col - rank_idx) / total_in_col
                        all_prots.add(p)
        if len(signs) >= max(1, min_overlap):
            sample_g_vectors.append({
                'gkey': f"{sample_name}_G{n}",
                'roi': sample_name,
                'g_num': n,
                'signs': signs,
                'magnitudes': magnitudes,
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
    Ingests user-supplied Excel files, ZIP archives, or directories.
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
                td = tempfile.mkdtemp(prefix='sa_es_upload_')
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
        raise ValueError("No valid Excel spreadsheet files (.xlsx / .xls) found in the uploaded input.")

    for fpath in files_to_process:
        parent_dir = os.path.basename(os.path.dirname(fpath))
        base_name = os.path.splitext(os.path.basename(fpath))[0]
        sample_name = parent_dir if (base_name.lower() in ('data', 'dataset', 'tail_proteins') and parent_dir and not parent_dir.startswith('sa_es_upload_')) else base_name
        # Ensure unique sample_name if multiple files share base_name
        if sample_name in all_roi_info and parent_dir:
            sample_name = f"{parent_dir}_{base_name}"

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
        raise ValueError("Could not extract any valid G-vectors from the provided files. Please verify that the Excel sheets match Format A (Gvalues/Lambavalues) or Format B (G1+/G1- Tails).")

    return all_roi_info, all_g_data


def inspect_user_data(file_paths=None, threshold=0.05):
    """
    Pre-Flight Dataset Inspection before running the full greedy algorithm.
    Returns counts of datasets, G-vectors, unique & shared genes/proteins,
    tail length distribution, suggested sequence length bounds, and estimated runtime.
    """
    roi_info, g_data = ingest_user_data(file_paths=file_paths, threshold=threshold, min_overlap=1)
    num_datasets = len(roi_info)
    total_g = len(g_data)

    # Feature sets per dataset (all features present in dataset)
    dataset_feature_sets = []
    for r_id, info in roi_info.items():
        pset = set(info.get('proteins', []))
        if not pset:
            pset = set().union(*(g['prots'] for g in g_data if g['roi'] == r_id))
        dataset_feature_sets.append(pset)

    all_features = set().union(*dataset_feature_sets) if dataset_feature_sets else set()
    shared_all_datasets = set.intersection(*dataset_feature_sets) if dataset_feature_sets else set()

    feat_dataset_counts = defaultdict(int)
    for pset in dataset_feature_sets:
        for p in pset:
            feat_dataset_counts[p] += 1
    shared_2plus_datasets = {p for p, c in feat_dataset_counts.items() if c >= 2} if num_datasets > 1 else all_features

    # Active tail features (features with |G| >= threshold in at least one G-vector)
    active_tail_features = set().union(*(g['prots'] for g in g_data)) if g_data else set()

    tail_lengths = [g['count'] for g in g_data if g['count'] > 0]
    min_tail = int(min(tail_lengths)) if tail_lengths else 0
    med_tail = int(np.median(tail_lengths)) if tail_lengths else 0
    max_tail = int(max(tail_lengths)) if tail_lengths else 0

    # Suggest sensible sequence lengths based on median/max tail size
    if med_tail <= 12:
        sug_min = 3
        sug_max = max(4, min(10, max_tail))
    else:
        sug_min = max(3, int(round(med_tail * 0.15)))
        sug_max = max(sug_min + 5, int(round(med_tail * 0.60)))

    # Estimate runtime in seconds
    est_sec = round(1.0 + 0.0025 * total_g + 0.04 * num_datasets + (0.0002 * len(all_features)), 1)
    est_sec = max(1.0, min(est_sec, 120.0))

    return {
        'num_datasets': num_datasets,
        'total_g_vectors': total_g,
        'total_unique_features': len(all_features),
        'active_tail_features': len(active_tail_features),
        'shared_all_datasets': len(shared_all_datasets),
        'shared_2plus_datasets': len(shared_2plus_datasets),
        'min_tail_len': min_tail,
        'median_tail_len': med_tail,
        'max_tail_len': max_tail,
        'suggested_min_len': sug_min,
        'suggested_max_len': sug_max,
        'estimated_time_sec': est_sec,
        'dataset_names': list(roi_info.keys())
    }


def _truncate_signature(sign_dict, mag_dict, target_len):
    """Selects the top `target_len` features ordered by magnitude (then alphabetically)."""
    if len(sign_dict) <= target_len:
        return dict(sign_dict)
    ranked = sorted(sign_dict.keys(), key=lambda p: (-mag_dict.get(p, 1.0), p))[:target_len]
    return {p: sign_dict[p] for p in ranked}


def mine_candidate_signatures(g_data, min_es_length, max_es_length, min_overlap, min_g_recurrence):
    """
    Dual-Mode Candidate ES Mining Engine:
    - Mode 1 (Exact Combinatorial Subset Enumeration):
      Used when G-vector tails are compact (max tail <= 14 and max_es_length <= 10).
      Enumerates all exact combinations C(|G|, k) for k in [min_es_length, max_es_length].
    - Mode 2 (High-Dimensional Sign-Consistent Intersection Mining):
      Automatically engages when datasets contain large gene/protein panels (e.g. 100 to 20,000+ genes,
      and min_es_length up to 1,000+) where naive C(|G|, k) would combinatorially explode.
      Computes all direct and inverse sign-consistent intersections across G-vector pairs & clusters.
    """
    max_g_tail = max((g['count'] for g in g_data), default=0)

    # Estimate whether exact enumeration is tractable
    use_exact_combinatorics = (max_g_tail <= 14 and max_es_length <= 10 and min_es_length <= 10)
    seen_combos = defaultdict(list)

    if use_exact_combinatorics:
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
        return seen_combos, "Exact Combinatorial Enumeration"

    # Mode 2: High-Dimensional / Large-Length Intersection & Subset Mining
    candidate_cts = set()

    # A. Add bounded full G-vector tails for vectors with count >= min_es_length
    for g in g_data:
        if g['count'] >= min_es_length:
            for target_l in set([min(g['count'], max_es_length), min_es_length]):
                if min_es_length <= target_l <= max_es_length:
                    sub_d = _truncate_signature(g['signs'], g.get('magnitudes', {}), target_l)
                    cd = canon_dict(sub_d)
                    if len(cd) >= min_es_length:
                        candidate_cts.add(tuple(sorted(cd.items())))

    # B. Pairwise Direct & Inverse Sign-Consistent Intersections across active G-vectors
    valid_gs = [g for g in g_data if g['count'] >= min_es_length]
    # If very large number of G-vectors, prioritize top weighted G-vectors for pairwise intersections
    if len(valid_gs) > 600:
        valid_gs = sorted(valid_gs, key=lambda x: x['weight'], reverse=True)[:600]

    n_v = len(valid_gs)
    for i in range(n_v):
        ga = valid_gs[i]
        prots_a = ga['prots']
        signs_a = ga['signs']
        mags_a = ga.get('magnitudes', {})
        for j in range(i + 1, n_v):
            gb = valid_gs[j]
            common = prots_a & gb['prots']
            if len(common) < min_es_length:
                continue
            signs_b = gb['signs']
            mags_b = gb.get('magnitudes', {})

            # Direct concordant genes/proteins
            same_p = [p for p in common if signs_a[p] == signs_b[p]]
            # Inverse concordant genes/proteins
            opp_p = [p for p in common if signs_a[p] == -signs_b[p]]

            for concordant_list in (same_p, opp_p):
                c_len = len(concordant_list)
                if c_len >= min_es_length:
                    comb_signs = {p: signs_a[p] for p in concordant_list}
                    comb_mags = {p: (mags_a.get(p, 1.0) + mags_b.get(p, 1.0)) * 0.5 for p in concordant_list}
                    # Add max allowed length signature and min_es_length core signature
                    lengths_to_extract = {min(c_len, max_es_length), min_es_length}
                    if c_len >= (min_es_length + max_es_length) // 2:
                        lengths_to_extract.add(min(c_len, (min_es_length + max_es_length) // 2))
                    for tl in lengths_to_extract:
                        if min_es_length <= tl <= max_es_length:
                            sub_d = _truncate_signature(comb_signs, comb_mags, tl)
                            cd = canon_dict(sub_d)
                            if len(cd) >= min_es_length:
                                candidate_cts.add(tuple(sorted(cd.items())))

    # Also if min_es_length <= 8, supplement with top-ranked feature subsets per G-vector
    if min_es_length <= 8:
        for g in g_data:
            if g['count'] >= min_es_length:
                top_sub = _truncate_signature(g['signs'], g.get('magnitudes', {}), min(10, g['count']))
                items = list(top_sub.items())
                for k in range(min_es_length, min(len(items) + 1, min(8, max_es_length) + 1)):
                    for c in combinations(items, k):
                        cd = canon_dict(dict(c))
                        candidate_cts.add(tuple(sorted(cd.items())))

    # Evaluate candidate coverage across all G-vectors
    eff_overlap = max(min_overlap, min_es_length) if min_es_length > 8 else min_overlap
    for ct in candidate_cts:
        cd = dict(ct)
        for g in g_data:
            orient, _, _ = match_es_g(cd, g['signs'], g['prots'], min_overlap=eff_overlap)
            if orient is not None:
                seen_combos[ct].append(g['gkey'])

    return seen_combos, "Scalable Sign-Consistent Intersection Mining"


def generate_es_bar_chart_svg(df_es, lang='en'):
    """
    Generates a clean, responsive SVG Bar + Cumulative Line chart for Optimal_ES_SetCover.
    """
    if df_es is None or df_es.empty:
        return ""

    rows = df_es.to_dict('records')
    n = len(rows)
    chart_w = max(680, min(1150, 120 + n * 44))
    chart_h = 340
    pad_l, pad_r, pad_t, pad_b = 58, 58, 42, 76
    plot_w = chart_w - pad_l - pad_r
    plot_h = chart_h - pad_t - pad_b

    max_g = max((int(r['New_Gs_Covered']) for r in rows), default=1)
    max_g = max(max_g, 1)

    title = "Selected ES Signatures: G-Vectors Covered, Length & Cumulative Coverage (%)" if lang == 'en' else "חתימות ה-ES שנבחרו: וקטורי G שכוסו, אורך החתימה וכיסוי מצטבר (%)"
    leg_dir = "Direct Match (+)" if lang == 'en' else "התאמה ישרה (+)"
    leg_inv = "Inverse Match (-)" if lang == 'en' else "התאמה הפוכה (-)"
    leg_cum = "Cumulative Coverage (%)" if lang == 'en' else "כיסוי מצטבר (%)"
    leg_len = "L = Signature Length" if lang == 'en' else "L = אורך החתימה"

    svg = [
        f'<div style="overflow-x:auto; background:#FFFFFF; border:1px solid #CBD5E1; border-radius:8px; padding:12px; margin:10px 0;">',
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {chart_w} {chart_h}" width="100%" style="max-height:360px; font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;">',
        f'<text x="{chart_w/2}" y="22" text-anchor="middle" font-size="13.5" font-weight="bold" fill="#0F172A">{html.escape(title)}</text>'
    ]

    # Grid lines
    for i in range(5):
        frac = i / 4.0
        y = pad_t + plot_h - frac * plot_h
        g_val = int(round(frac * max_g))
        pct_val = int(round(frac * 100))
        svg.append(f'<line x1="{pad_l}" y1="{y:.1f}" x2="{pad_l + plot_w}" y2="{y:.1f}" stroke="#E2E8F0" stroke-dasharray="3,3" stroke-width="1"/>')
        svg.append(f'<text x="{pad_l - 8}" y="{y + 4:.1f}" text-anchor="end" font-size="10" fill="#334155">{g_val}</text>')
        svg.append(f'<text x="{pad_l + plot_w + 8}" y="{y + 4:.1f}" text-anchor="start" font-size="10" fill="#0284C7">{pct_val}%</text>')

    # Axes labels
    svg.append(f'<line x1="{pad_l}" y1="{pad_t}" x2="{pad_l}" y2="{pad_t + plot_h}" stroke="#475569" stroke-width="1.5"/>')
    svg.append(f'<line x1="{pad_l}" y1="{pad_t + plot_h}" x2="{pad_l + plot_w}" y2="{pad_t + plot_h}" stroke="#475569" stroke-width="1.5"/>')
    svg.append(f'<line x1="{pad_l + plot_w}" y1="{pad_t}" x2="{pad_l + plot_w}" y2="{pad_t + plot_h}" stroke="#0284C7" stroke-width="1.5"/>')

    step_w = plot_w / max(n, 1)
    bar_w = max(10, min(28, step_w * 0.62))

    pts = []
    for idx, r in enumerate(rows):
        cx = pad_l + (idx + 0.5) * step_w
        d_cnt = int(r['Direct (+)'])
        i_cnt = int(r['Inverse (-)'])
        tot_cnt = d_cnt + i_cnt
        seq_len = int(r['Length'])
        cov_str = str(r['Coverage_%']).replace('%', '').strip()
        cov_pct = float(cov_str) if cov_str else 0.0

        h_tot = (tot_cnt / max_g) * plot_h
        h_d = (d_cnt / max_g) * plot_h
        h_i = (i_cnt / max_g) * plot_h

        y_d = pad_t + plot_h - h_d
        y_i = y_d - h_i

        es_lbl = html.escape(str(r['ES_ID']))
        seq_tip = html.escape(str(r['Exact_Sign_Sequence'])[:120])

        # Stacked bars: Direct (+) in deep blue, Inverse (-) in coral red
        if d_cnt > 0:
            svg.append(f'<rect x="{cx - bar_w/2:.1f}" y="{y_d:.1f}" width="{bar_w:.1f}" height="{max(1.0, h_d):.1f}" fill="#2563EB" rx="2"><title>{es_lbl}: {d_cnt} Direct (+) | L={seq_len} | {seq_tip}</title></rect>')
        if i_cnt > 0:
            svg.append(f'<rect x="{cx - bar_w/2:.1f}" y="{y_i:.1f}" width="{bar_w:.1f}" height="{max(1.0, h_i):.1f}" fill="#DC2626" rx="2"><title>{es_lbl}: {i_cnt} Inverse (-) | L={seq_len} | {seq_tip}</title></rect>')

        # Top value label & length badge
        top_y = pad_t + plot_h - h_tot - 5
        svg.append(f'<text x="{cx:.1f}" y="{max(pad_t + 8, top_y):.1f}" text-anchor="middle" font-size="9.5" font-weight="bold" fill="#0F172A">{tot_cnt}</text>')

        # X-axis ES_ID and Length badge
        svg.append(f'<text x="{cx:.1f}" y="{pad_t + plot_h + 15}" text-anchor="middle" font-size="9.5" font-weight="bold" fill="#0F172A">{es_lbl}</text>')
        svg.append(f'<text x="{cx:.1f}" y="{pad_t + plot_h + 28}" text-anchor="middle" font-size="8.5" fill="#475569">L={seq_len}</text>')

        # Cumulative coverage point
        cy_pct = pad_t + plot_h - (cov_pct / 100.0) * plot_h
        pts.append((cx, cy_pct, cov_pct, es_lbl))

    # Draw cumulative coverage polyline
    if len(pts) > 1:
        poly_str = " ".join(f"{x:.1f},{y:.1f}" for x, y, _, _ in pts)
        svg.append(f'<polyline fill="none" stroke="#0284C7" stroke-width="2.2" points="{poly_str}"/>')
    for x, y, pct, lbl in pts:
        svg.append(f'<circle cx="{x:.1f}" cy="{y:.1f}" r="3.5" fill="#FFFFFF" stroke="#0284C7" stroke-width="2"><title>{lbl}: {pct:.1f}% Cumulative Coverage</title></circle>')

    # Legend at bottom
    leg_y = chart_h - 14
    svg.append(f'<rect x="{pad_l}" y="{leg_y - 9}" width="11" height="11" fill="#2563EB" rx="2"/>')
    svg.append(f'<text x="{pad_l + 15}" y="{leg_y}" font-size="10" fill="#0F172A">{html.escape(leg_dir)}</text>')

    svg.append(f'<rect x="{pad_l + 145}" y="{leg_y - 9}" width="11" height="11" fill="#DC2626" rx="2"/>')
    svg.append(f'<text x="{pad_l + 160}" y="{leg_y}" font-size="10" fill="#0F172A">{html.escape(leg_inv)}</text>')

    svg.append(f'<circle cx="{pad_l + 300}" cy="{leg_y - 4}" r="4" fill="#FFFFFF" stroke="#0284C7" stroke-width="2"/>')
    svg.append(f'<text x="{pad_l + 310}" y="{leg_y}" font-size="10" fill="#0284C7" font-weight="bold">{html.escape(leg_cum)}</text>')

    svg.append(f'<text x="{pad_l + 495}" y="{leg_y}" font-size="10" fill="#475569">{html.escape(leg_len)}</text>')

    svg.append('</svg></div>')
    return "\n".join(svg)


def generate_conflict_free_heatmap_html(df_mat, df_es, lang='en'):
    """
    Generates an interactive HTML Heatmap for ES_Conflict_Free_Matrix.
    Color-codes Direct (+) matches in Blue, Inverse (-) matches in Red/Coral, and empty cells in light gray.
    """
    if df_mat is None or df_mat.empty:
        return ""

    es_seq_map = {}
    if df_es is not None and not df_es.empty:
        for _, r in df_es.iterrows():
            es_seq_map[str(r['ES_ID'])] = f"L={r['Length']} | {r['Exact_Sign_Sequence']}"

    es_cols = [c for c in df_mat.columns if c.startswith('ES_')]
    title = "ES_Conflict_Free_Matrix Heatmap (Datasets × Selected ES Signatures)" if lang == 'en' else "מפת חום (Heatmap) – מטריצת הקצאה נקייה מסתירות (דאטה-סטים × חתימות ES)"
    sub = "Blue = Direct Match (+)  |  Red = Inverse Match (-)  |  Hover over any cell to inspect the assigned G-vector & ES signature" if lang == 'en' else "כחול = התאמה ישרה (+)  |  אדום = התאמה הפוכה (-)  |  עבור עם העכבר על כל תא לצפייה בווקטור ה-G ובחתימת ה-ES"

    html_parts = [
        '<div style="background:#FFFFFF; border:1px solid #CBD5E1; border-radius:8px; padding:12px; margin:10px 0;">',
        f'<div style="font-weight:bold; font-size:0.98rem; color:#0F172A; margin-bottom:4px;">{html.escape(title)}</div>',
        f'<div style="font-size:0.82rem; color:#475569; margin-bottom:10px;">{html.escape(sub)}</div>',
        '<div style="max-height:420px; overflow:auto; border:1px solid #E2E8F0; border-radius:6px;">',
        '<table style="border-collapse:collapse; width:100%; font-size:0.78rem; text-align:center;">',
        '<thead style="position:sticky; top:0; background:#F1F5F9; z-index:2;"><tr>',
        '<th style="padding:6px 10px; border:1px solid #CBD5E1; text-align:left; min-width:120px; background:#F1F5F9; color:#0F172A;">Dataset</th>'
    ]

    for col in es_cols:
        tip = html.escape(es_seq_map.get(col, col))
        html_parts.append(f'<th title="{tip}" style="padding:6px 6px; border:1px solid #CBD5E1; min-width:52px; cursor:help; background:#F1F5F9; color:#0F172A;">{html.escape(col)}</th>')
    html_parts.append('</tr></thead><tbody>')

    sample_col = 'Dataset' if 'Dataset' in df_mat.columns else df_mat.columns[0]
    for _, row in df_mat.iterrows():
        ds_name = str(row[sample_col])
        html_parts.append(f'<tr><td style="padding:5px 10px; border:1px solid #E2E8F0; text-align:left; font-weight:600; background:#F8FAFC; color:#0F172A; white-space:nowrap;">{html.escape(ds_name)}</td>')
        for col in es_cols:
            val = str(row.get(col, '')).strip()
            seq_info = es_seq_map.get(col, '')
            if '(+)' in val:
                bg = '#DBEAFE'
                fg = '#1E40AF'
                fw = 'bold'
                tip = html.escape(f"Dataset: {ds_name} | {col} -> {val} [Direct +] | {seq_info}")
            elif '(-)' in val:
                bg = '#FEE2E2'
                fg = '#991B1B'
                fw = 'bold'
                tip = html.escape(f"Dataset: {ds_name} | {col} -> {val} [Inverse -] | {seq_info}")
            else:
                bg = '#FFFFFF'
                fg = '#CBD5E1'
                fw = 'normal'
                val = '·'
                tip = html.escape(f"Dataset: {ds_name} | {col}: Unassigned")
            html_parts.append(
                f'<td title="{tip}" style="padding:4px 5px; border:1px solid #E2E8F0; background:{bg}; color:{fg}; font-weight:{fw}; white-space:nowrap;">{html.escape(val)}</td>'
            )
        html_parts.append('</tr>')

    html_parts.append('</tbody></table></div></div>')
    return "\n".join(html_parts)


def generate_html_summary_report(output_path, summary, params, df_es, df_mat, df_dataset, df_anc):
    """
    Generates a standalone, comprehensive HTML report summarizing the run,
    execution parameters, metrics, visual charts, heatmap, and core tables.
    """
    bar_svg = generate_es_bar_chart_svg(df_es, lang='en')
    heatmap_html = generate_conflict_free_heatmap_html(df_mat, df_es, lang='en')

    def df_to_styled_html(df, max_rows=200):
        if df is None or df.empty:
            return "<p>No data available.</p>"
        sub_df = df.head(max_rows)
        return sub_df.to_html(index=False, border=0, classes="report-table")

    html_doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Surprisal Analysis (SA) – Minimal ES Signatures Summary Report</title>
<style>
  body {{
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    background: #F8FAFC;
    color: #0F172A;
    margin: 0;
    padding: 28px;
    line-height: 1.55;
  }}
  .container {{
    max-width: 1240px;
    margin: 0 auto;
    background: #FFFFFF;
    border: 1px solid #CBD5E1;
    border-radius: 10px;
    padding: 28px;
    box-shadow: 0 4px 12px rgba(15, 23, 42, 0.05);
  }}
  h1 {{ margin: 0 0 6px 0; font-size: 1.65rem; color: #0F172A; }}
  h2 {{ margin-top: 28px; font-size: 1.2rem; border-bottom: 2px solid #0F172A; padding-bottom: 6px; color: #0F172A; }}
  .subtitle {{ color: #475569; font-size: 0.95rem; margin-bottom: 20px; }}
  .kpi-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(155px, 1fr));
    gap: 12px;
    margin: 16px 0 24px 0;
  }}
  .kpi-card {{
    border: 1px solid #CBD5E1;
    border-radius: 8px;
    padding: 12px;
    background: #F8FAFC;
  }}
  .kpi-label {{ font-size: 0.75rem; text-transform: uppercase; color: #475569; font-weight: 700; }}
  .kpi-val {{ font-size: 1.45rem; font-weight: 800; color: #0F172A; margin-top: 4px; }}
  .params-box {{
    background: #F1F5F9;
    border: 1px solid #CBD5E1;
    border-radius: 8px;
    padding: 12px 16px;
    font-size: 0.88rem;
    margin-bottom: 18px;
  }}
  .table-wrap {{ overflow-x: auto; margin: 12px 0 24px 0; border: 1px solid #E2E8F0; border-radius: 6px; }}
  table.report-table {{ width: 100%; border-collapse: collapse; font-size: 0.83rem; text-align: left; }}
  table.report-table th {{ background: #F1F5F9; color: #0F172A; padding: 8px 10px; border-bottom: 2px solid #CBD5E1; white-space: nowrap; }}
  table.report-table td {{ padding: 7px 10px; border-bottom: 1px solid #E2E8F0; color: #1E293B; }}
  table.report-table tr:nth-child(even) {{ background: #F8FAFC; }}
</style>
</head>
<body>
<div class="container">
  <h1>Surprisal Analysis (SA) – Minimal Exact Sign Signatures (ES) Report</h1>
  <div class="subtitle">Multi-Dataset G-Vector Compression &bull; Greedy Set Cover Optimization &bull; Reconstruction Fidelity (R²)</div>

  <div class="params-box">
    <strong>Execution Parameters:</strong>
    Min Sequence Length = <code>{params['min_es_length']}</code> &bull;
    Max Sequence Length = <code>{params['max_es_length']}</code> &bull;
    Cutoff Loading (|G|) &ge; <code>{params['threshold']}</code> &bull;
    Min G Recurrence &ge; <code>{params['min_g_recurrence']}</code> &bull;
    Mining Mode = <code>{summary.get('mining_mode', 'Standard')}</code> &bull;
    Runtime = <code>{summary['duration_sec']}s</code>
  </div>

  <h2>1. Executive Summary Metrics</h2>
  <div class="kpi-grid">
    <div class="kpi-card"><div class="kpi-label">Datasets Analyzed</div><div class="kpi-val">{summary['total_samples']}</div></div>
    <div class="kpi-card"><div class="kpi-label">Total G-Vectors</div><div class="kpi-val">{summary['total_g']:,}</div></div>
    <div class="kpi-card"><div class="kpi-label">Unique Genes/Proteins</div><div class="kpi-val">{summary.get('total_unique_features', 0):,}</div></div>
    <div class="kpi-card"><div class="kpi-label">Shared Across Datasets</div><div class="kpi-val">{summary.get('shared_all_datasets', 0):,}</div></div>
    <div class="kpi-card"><div class="kpi-label">Minimal ES Signatures</div><div class="kpi-val">{summary['es_count']}</div></div>
    <div class="kpi-card"><div class="kpi-label">G-Vector Coverage</div><div class="kpi-val">{summary['coverage_pct']}%</div></div>
    <div class="kpi-card"><div class="kpi-label">Explained Variance</div><div class="kpi-val">{summary['variance_explained_pct']}%</div></div>
    <div class="kpi-card"><div class="kpi-label">Mean Optimal R²</div><div class="kpi-val">{summary['avg_r2_optimal']:.4f}</div></div>
    <div class="kpi-card"><div class="kpi-label">Compression Loss (ΔR²)</div><div class="kpi-val">{summary['avg_delta_r2']:.4f}</div></div>
    <div class="kpi-card"><div class="kpi-label">High Fidelity Rate</div><div class="kpi-val">{summary['high_fidelity_pct']}%</div></div>
  </div>

  <h2>2. Visual Overview: Selected ES Signatures & Cumulative Coverage</h2>
  {bar_svg}

  <h2>3. Conflict-Free Allocation Heatmap (Datasets &times; ES Signatures)</h2>
  {heatmap_html}

  <h2>4. Optimal ES Set Cover Table (Optimal_ES_SetCover)</h2>
  <div class="table-wrap">{df_to_styled_html(df_es)}</div>

  <h2>5. Unified Per-Dataset Coverage &amp; R² Fidelity (Per_Dataset_Coverage_and_R2)</h2>
  <div class="table-wrap">{df_to_styled_html(df_dataset)}</div>

  <h2>6. Anchor Strategies Benchmark Overview</h2>
  <div class="table-wrap">{df_to_styled_html(df_anc)}</div>
</div>
</body>
</html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html_doc)
    return output_path


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
    Executes the complete pipeline on user-supplied SA datasets and produces the streamlined deliverables:
    1. Optimal_ES_Result.xlsx (5 focused sheets: Summary, Optimal_ES_SetCover, ES_Conflict_Free_Matrix, Per_Dataset_Coverage_and_R2, Per_G_Coverage_Detail)
    2. Analysis_Summary_Report.html (Visual HTML summary report)
    3. Anchor_Strategies_Comparison.xlsx (Overview benchmark sheet)
    4. ES_Combinations.csv (Recurring candidate signatures in 2+ Gs)
    Plus Deliverables_Bundle.zip.
    """
    def log(msg, pct=None):
        if progress_callback:
            progress_callback(pct, msg)
        print(f"[{time.strftime('%H:%M:%S')}] {msg}")

    t0 = time.time()
    os.makedirs(output_dir, exist_ok=True)
    min_es_length = max(2, int(min_es_length))
    max_es_length = max(min_es_length, int(max_es_length))
    min_g_recurrence = max(1, int(min_g_recurrence))
    min_overlap = max(2, int(min_overlap))

    log("Ingesting user-supplied Surprisal Analysis datasets...", 0.05)

    roi_info, g_data = ingest_user_data(
        file_paths=file_paths,
        folder_path=folder_path,
        threshold=threshold,
        min_overlap=min_overlap
    )

    total_g = len(g_data)
    total_rois = len(roi_info)

    # Compute feature sharing statistics across datasets
    ds_sets = [set(info.get('proteins', [])) or set().union(*(g['prots'] for g in g_data if g['roi'] == r_id)) for r_id, info in roi_info.items()]
    all_features = set().union(*ds_sets) if ds_sets else set()
    shared_all_features = set.intersection(*ds_sets) if ds_sets else set()

    log(f"Loaded {total_g:,} G-vectors across {total_rois} dataset(s) ({len(all_features):,} unique features).", 0.15)

    # Dual-Mode Candidate ES Mining
    log(f"Mining canonical Exact Sign Signatures (lengths {min_es_length} to {max_es_length})...", 0.25)
    seen_combos, mining_mode = mine_candidate_signatures(
        g_data=g_data,
        min_es_length=min_es_length,
        max_es_length=max_es_length,
        min_overlap=min_overlap,
        min_g_recurrence=min_g_recurrence
    )

    total_mined = len(seen_combos)
    log(f"Mined {total_mined:,} candidate signatures via {mining_mode}.", 0.40)

    # 1. Export ES_Combinations.csv (Filtered to recurring combinations >= min_g_recurrence so singletons don't bloat the file)
    combo_records = []
    for ct, gkeys in seen_combos.items():
        unique_gkeys = sorted(set(gkeys))
        if len(unique_gkeys) >= min_g_recurrence:
            cd = dict(ct)
            combo_records.append({
                'Length': len(cd),
                'Frequency_Count': len(unique_gkeys),
                'Exact_Sign_Sequence': tuple_to_str(cd),
                'Covered_G_Vectors': "; ".join(unique_gkeys)
            })

    # Fallback if no recurring combinations met min_g_recurrence
    if not combo_records:
        for ct, gkeys in seen_combos.items():
            unique_gkeys = sorted(set(gkeys))
            cd = dict(ct)
            combo_records.append({
                'Length': len(cd),
                'Frequency_Count': len(unique_gkeys),
                'Exact_Sign_Sequence': tuple_to_str(cd),
                'Covered_G_Vectors': "; ".join(unique_gkeys)
            })

    df_combos = pd.DataFrame(combo_records)
    if not df_combos.empty:
        df_combos.sort_values(by=['Frequency_Count', 'Length'], ascending=[False, False], inplace=True)
    es_combinations_csv_path = os.path.join(output_dir, "ES_Combinations.csv")
    df_combos.to_csv(es_combinations_csv_path, index=False)
    log("Generated deliverable: ES_Combinations.csv (filtered recurring signatures)", 0.45)

    # Select candidate modules for Set Cover optimization
    log("Selecting candidate modules for Set Cover optimization...", 0.50)
    cands_sorted = sorted(seen_combos.items(), key=lambda x: len(set(x[1])), reverse=True)
    if len(cands_sorted) > 4000:
        top_cands = [ct for ct, glist in cands_sorted if len(set(glist)) >= min_g_recurrence][:3500]
        seen_ct = set(top_cands)
        for g in g_data:
            if len(g['signs']) >= min_es_length:
                sub_d = _truncate_signature(g['signs'], g.get('magnitudes', {}), min(len(g['signs']), max_es_length))
                ct = tuple(sorted(canon_dict(sub_d).items()))
                if ct not in seen_ct:
                    seen_ct.add(ct)
                    top_cands.append(ct)
    else:
        top_cands = [ct for ct, glist in cands_sorted if len(set(glist)) >= min_g_recurrence]
        if not top_cands:
            top_cands = [ct for ct, _ in cands_sorted[:1000]]

    eff_overlap = max(min_overlap, min_es_length) if min_es_length > 8 else min_overlap
    eligible_modules = []
    for ct in top_cands:
        mod_dict = dict(ct)
        hits = {}
        tot_var = 0.0
        for g in g_data:
            orient, n_common, matched = match_es_g(mod_dict, g['signs'], g['prots'], min_overlap=eff_overlap)
            if orient is not None:
                hits[g['g_idx']] = (orient, n_common, matched)
                tot_var += g['weight']
        if len(hits) >= min_g_recurrence or (not eligible_modules and len(hits) >= 1):
            eligible_modules.append({
                'ct': ct,
                'mod': mod_dict,
                'len': len(mod_dict),
                'hits': hits,
                'initial_count': len(hits),
                'initial_var': tot_var
            })

    # 2. Evaluate Anchor Strategies
    log("Evaluating Greedy Set Cover & Anchor Strategies...", 0.55)
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

        # Compute R2 per dataset
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
            'Mean_R2_full': round(float(np.mean(r2_fulls)), 4) if r2_fulls else 1.0,
            'Mean_R2_opt': round(float(np.mean(r2_opts)), 4) if r2_opts else 1.0,
            'Mean_Delta_R2': round(float(np.mean(r2_diffs)), 4) if r2_diffs else 0.0,
            'Excellent_%': round(float(np.mean(np.array(r2_diffs) < 0.05) * 100), 1) if r2_diffs else 100.0,
            'Acceptable_%': round(float(np.mean((np.array(r2_diffs) >= 0.05) & (np.array(r2_diffs) <= 0.15)) * 100), 1) if r2_diffs else 0.0,
            'Alert_%': round(float(np.mean(np.array(r2_diffs) > 0.15) * 100), 1) if r2_diffs else 0.0,
            'Exec_Time_s': round(time.time() - t_strat, 2),
            'selected': selected,
            'g_assignment': g_assignment,
            'r2_per_sample': {r_id: (r2_fulls[i], r2_opts[i], r2_diffs[i], sorted(roi_g_assigned.get(r_id, set()))) for i, r_id in enumerate(roi_info.keys())}
        }

    strategies_eval = []
    # 1. SVD Variance Driver
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
    mid_len = max(5, (min_es_length + max_es_length) // 2)
    strategies_eval.append(evaluate_strategy(
        "3. Deep Specificity", None,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g) * (1.5 if m['len'] >= mid_len else 1.0)
    ))
    # 4. Dominant Co-expression Seed
    pair_counts = defaultdict(int)
    for g in g_data:
        prots_sample = sorted(g['prots'])[:40]  # Bound pairwise scan for 20k gene speed
        for p1, p2 in combinations(prots_sample, 2):
            if g['signs'][p1] == g['signs'][p2]:
                pair_counts[(p1, p2)] += 1
    top_pair = max(pair_counts.items(), key=lambda x: x[1])[0] if pair_counts else None
    seed_func = (lambda m: top_pair[0] in m['mod'] and top_pair[1] in m['mod']) if top_pair else None
    strategies_eval.append(evaluate_strategy(
        f"4. Dominant Seed ({top_pair[0]} & {top_pair[1]})" if top_pair else "4. Dominant Co-expression Seed",
        seed_func,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g) * np.sqrt(len(new_g))
    ))
    # 5. Hybrid Multi-Objective (Variance x Count)
    hybrid_res = evaluate_strategy(
        "5. Hybrid Multi-Objective (Variance x Count)", None,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g) * np.sqrt(len(new_g))
    )
    strategies_eval.append(hybrid_res)

    target_strat = hybrid_res

    # Export streamlined Anchor_Strategies_Comparison.xlsx (Only Overview sheet, no duplicate Strat_5)
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
    wb_anc.save(anchor_xlsx_path)
    log("Generated deliverable: Anchor_Strategies_Comparison.xlsx", 0.70)

    # 3. Build Streamlined Master Excel Deliverable: Optimal_ES_Result.xlsx (5 Non-Redundant Sheets)
    log("Constructing Streamlined Master Workbook: Optimal_ES_Result.xlsx...", 0.80)
    headers_es = [
        'Rank', 'ES_Label', 'Sequence_Length', 'New_G_Covered',
        'Direct_Matches_(+)', 'Inverse_Matches_(-)', 'Cumulative_G_Covered',
        'Coverage_%', 'New_Variance_Weight', 'Cumulative_Variance_%',
        'Datasets_Covered_Count', 'Exact_Sign_Sequence'
    ]

    wb = openpyxl.Workbook()
    # Sheet 1: Summary
    ws_sum = wb.active
    ws_sum.title = 'Summary'
    ws_sum.append(['SURPRISAL ANALYSIS (SA) – MINIMAL EXACT SIGN SIGNATURES (ES) REPORT'])
    ws_sum.append(['Multi-Dataset G-Vector Set Cover, Conflict-Free Allocation & R2 Reconstruction Fidelity'])
    ws_sum.append([])
    ws_sum.append(['Metric', 'Value'])
    ws_sum.append(['Optimization Strategy Applied', target_strat['Anchor_Strategy']])
    ws_sum.append(['Mining Mode', mining_mode])
    ws_sum.append(['Min / Max Sequence Length', f"{min_es_length} – {max_es_length}"])
    ws_sum.append(['Cutoff Loading (|G|)', threshold])
    ws_sum.append(['Min G Recurrence', min_g_recurrence])
    ws_sum.append(['Total Datasets Analyzed', total_rois])
    ws_sum.append(['Total G-Vectors Across Datasets', total_g])
    ws_sum.append(['Total Unique Genes / Proteins', len(all_features)])
    ws_sum.append(['Genes / Proteins Shared Across All Datasets', len(shared_all_features)])
    ws_sum.append(['Minimal ES Signatures Selected', target_strat['ES_Count']])
    ws_sum.append(['Total G-Vectors Covered', target_strat['Covered_Gs']])
    ws_sum.append(['G-Vector Coverage %', f"{target_strat['Coverage_%']}%"])
    ws_sum.append(['Total Variance Explained %', f"{target_strat['Variance_%']}%"])
    ws_sum.append(['Mean R2 Full Reconstruction', target_strat['Mean_R2_full']])
    ws_sum.append(['Mean R2 Optimal Reconstruction', target_strat['Mean_R2_opt']])
    ws_sum.append(['Mean Delta R2 Compression Loss', target_strat['Mean_Delta_R2']])
    ws_sum.append(['High Fidelity Datasets (<0.05 Delta R2)', f"{target_strat['Excellent_%']}%"])

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
        ds_cnt = len(set(g_data[gi]['roi'] for gi in es['new_g']))
        ws_opt.append([
            rk, es['lbl'], es['len'], len(es['new_g']), d_cnt, i_cnt,
            cum_g_local, round(cum_g_local / total_g * 100, 2) if total_g > 0 else 0,
            round(es['var_gain'], 4), round(cum_var_local / total_rois * 100, 2) if total_rois > 0 else 0,
            ds_cnt, tuple_to_str(es['mod'])
        ])

    # Sheet 3: ES_Conflict_Free_Matrix
    ws_mat = wb.create_sheet(title='ES_Conflict_Free_Matrix')
    mat_cols = ['Dataset', 'Cells_or_Observations', 'Total_G'] + [es['lbl'] for es in target_strat['selected']]
    ws_mat.append(mat_cols)
    for r_id, s in roi_info.items():
        row = [r_id, s['n_cells'], len([g for g in g_data if g['roi'] == r_id])]
        used_g = set()
        for es in target_strat['selected']:
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

    # Sheet 4: Per_Dataset_Coverage_and_R2 (Merged from Best_Combination_Details + Per_ROI_Coverage + R2_Comparison_Per_ROI)
    ws_ds = wb.create_sheet(title='Per_Dataset_Coverage_and_R2')
    ws_ds.append([
        'Dataset', 'Cells_or_Observations', 'Total_G', 'Covered_G', 'Coverage_%',
        'Variance_Captured_%', 'R2_Full', 'R2_Optimal', 'Delta_R2_Loss',
        'Fidelity_Tier', 'Assigned_ES_Signatures', 'Detailed_G_Assignments'
    ])
    dataset_preview_rows = []
    for r_id, s in roi_info.items():
        sample_gs = [g for g in g_data if g['roi'] == r_id]
        cov_gs = [g for g in sample_gs if g['g_idx'] in target_strat['g_assignment']]
        var_cap = sum(g['weight'] for g in cov_gs)
        assigned_es = sorted(set(target_strat['g_assignment'][g['g_idx']][0] for g in cov_gs), key=lambda x: int(x.split('_')[1]) if '_' in x and x.split('_')[1].isdigit() else x)
        assignments_str = []
        for g in sorted(cov_gs, key=lambda x: x['g_num']):
            lbl, hit = target_strat['g_assignment'][g['g_idx']]
            orient = hit[0]
            pct = (hit[1] / g['count'] * 100) if g['count'] > 0 else 0
            assignments_str.append(f"G{g['g_num']}->{lbl}({orient})[{pct:.0f}%]")

        r2_f, r2_o, delta, _ = target_strat['r2_per_sample'].get(r_id, (1.0, 1.0, 0.0, []))
        interp = 'High Fidelity (<0.05 loss)' if delta < 0.05 else ('Acceptable (0.05-0.15)' if delta <= 0.15 else 'Noticeable Loss (>0.15)')
        cov_pct_ds = round(len(cov_gs) / len(sample_gs) * 100, 2) if sample_gs else 0.0
        var_pct_ds = round(var_cap * 100, 2)

        ws_ds.append([
            r_id, s['n_cells'], len(sample_gs), len(cov_gs),
            cov_pct_ds, var_pct_ds,
            round(r2_f, 4), round(r2_o, 4), round(delta, 4),
            interp, "; ".join(assigned_es), ", ".join(assignments_str)
        ])
        dataset_preview_rows.append({
            'Dataset': r_id,
            'Total_G': len(sample_gs),
            'Covered_G': len(cov_gs),
            'Coverage_%': f"{cov_pct_ds:.1f}%",
            'Variance_%': f"{var_pct_ds:.1f}%",
            'R2_Full': round(r2_f, 4),
            'R2_Optimal': round(r2_o, 4),
            'Delta_R2': round(delta, 4),
            'Fidelity_Tier': interp,
            'Assigned_ES': "; ".join(assigned_es)
        })

    # Sheet 5: Per_G_Coverage_Detail
    ws_perg = wb.create_sheet(title='Per_G_Coverage_Detail')
    ws_perg.append([
        'Dataset', 'G_Key', 'G_Num', 'Total_Features_In_G', 'Assigned_ES',
        'Orientation', 'Matched_Features_Count', 'Overlap_%',
        'Variance_Weight', 'Status', 'Matched_Features_List'
    ])
    for g in g_data:
        if g['g_idx'] in target_strat['g_assignment']:
            lbl, hit = target_strat['g_assignment'][g['g_idx']]
            orient, n_m, matched_list = hit
            pct = round(n_m / g['count'] * 100, 1) if g['count'] > 0 else 0
            ws_perg.append([
                g['roi'], g['gkey'], g['g_num'], g['count'], lbl, orient,
                n_m, pct, round(g['weight'], 4), 'Covered', ", ".join(matched_list[:100])
            ])
        else:
            ws_perg.append([
                g['roi'], g['gkey'], g['g_num'], g['count'], 'UNASSIGNED', 'None',
                0, 0.0, round(g['weight'], 4), 'Uncovered', ''
            ])

    opt_es_result_path = os.path.join(output_dir, "Optimal_ES_Result.xlsx")
    wb.save(opt_es_result_path)
    log("Generated streamlined master deliverable: Optimal_ES_Result.xlsx (5 sheets)", 0.88)

    # Prepare DataFrames for GUI preview & HTML report
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
            'Variance_%': f"{(cum_var / total_rois * 100.0):.1f}%",
            'Exact_Sign_Sequence': tuple_to_str(es['mod'], max_items=25)
        })
    df_es = pd.DataFrame(es_preview_rows)

    mat_rows = []
    for r_id, s in roi_info.items():
        row_dict = {'Dataset': r_id, 'Total_G': len([g for g in g_data if g['roi'] == r_id])}
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

    df_dataset = pd.DataFrame(dataset_preview_rows)

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

    duration_sec = round(time.time() - t0, 2)
    summary = {
        'total_samples': total_rois,
        'total_g': total_g,
        'total_unique_features': len(all_features),
        'shared_all_datasets': len(shared_all_features),
        'mining_mode': mining_mode,
        'es_count': target_strat['ES_Count'],
        'g_covered': target_strat['Covered_Gs'],
        'coverage_pct': target_strat['Coverage_%'],
        'variance_explained_pct': target_strat['Variance_%'],
        'avg_r2_optimal': target_strat['Mean_R2_opt'],
        'avg_delta_r2': target_strat['Mean_Delta_R2'],
        'high_fidelity_pct': target_strat['Excellent_%'],
        'duration_sec': duration_sec,
        'opt_es_result_path': opt_es_result_path,
        'es_combinations_csv_path': es_combinations_csv_path,
        'anchor_xlsx_path': anchor_xlsx_path,
    }

    # 4. Generate HTML Summary Report
    html_report_path = os.path.join(output_dir, "Analysis_Summary_Report.html")
    generate_html_summary_report(
        output_path=html_report_path,
        summary=summary,
        params={
            'min_es_length': min_es_length,
            'max_es_length': max_es_length,
            'threshold': threshold,
            'min_g_recurrence': min_g_recurrence
        },
        df_es=df_es,
        df_mat=df_mat,
        df_dataset=df_dataset,
        df_anc=df_anc
    )
    summary['html_report_path'] = html_report_path
    log("Generated visual summary report: Analysis_Summary_Report.html", 0.94)

    # 5. Create ZIP Bundle containing all deliverables
    zip_bundle_path = os.path.join(output_dir, "Deliverables_Bundle.zip")
    with zipfile.ZipFile(zip_bundle_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(opt_es_result_path, arcname="Optimal_ES_Result.xlsx")
        zf.write(html_report_path, arcname="Analysis_Summary_Report.html")
        zf.write(anchor_xlsx_path, arcname="Anchor_Strategies_Comparison.xlsx")
        zf.write(es_combinations_csv_path, arcname="ES_Combinations.csv")
    summary['zip_bundle_path'] = zip_bundle_path
    log(f"All deliverables generated in {summary['duration_sec']}s!", 1.00)

    return summary, df_es, df_mat, df_dataset, df_anc
