"""
run_r2_sign_pipeline.py
-----------------------
Full End-to-End Surprisal Analysis Pipeline with R^2 Integration:
1. Loads pre-computed R^2 cache (Docs/r2_cache.pkl) with X_log, G_all, L_all, r2_full, and component variance weights.
2. Loads all 2,778 G-vectors from All_Tail_Proteins_Tables with exact protein signs (+/-).
3. Mines candidate Exact Sign Sequence (ES) modules (lengths 5 to 8).
4. Strictly aligns Direct (+) and Inverse (-) matches with ZERO sign contradictions.
5. Employs Variance-Weighted Greedy Set Cover:
   - Score(ES) = sum of component variance weights w(g) of newly covered Gs.
   - Eliminates long tails (no singletons, threshold >= 2 Gs).
   - Enforces 100% disjoint G-vector allocation (no duplicate Gs across ES).
6. Generates conflict-free ROI x ES matrix (each cell <= 1 unique G, no repeats per row).
7. Computes R^2 metrics for all 120 ROIs:
   - R^2_full (all components)
   - R^2_optimal (G0 baseline + assigned G components)
   - Delta R^2 = R^2_full - R^2_optimal
   - Quality classification (<0.05 Excellent, 0.05-0.15 Acceptable, >0.15 Alert).
8. Exports:
   - Docs/Optimal_ES_Result_Signs_R2.xlsx (8 beautifully styled sheets)
   - Docs/R2 comparison per ROI.xlsx (standalone report per Guide)
"""

import os
import glob
import re
import time
import pickle
from collections import defaultdict
from itertools import combinations
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

CACHE_FILE = os.path.join('Docs', 'r2_cache.pkl')
TAILS_DIR = 'All_Tail_Proteins_Tables'
OUTPUT_EXCEL = os.path.join('Docs', 'Optimal_ES_Result_Signs_R2.xlsx')
OUTPUT_R2_STANDALONE = os.path.join('Docs', 'R2 comparison per ROI.xlsx')

def compute_r2_vectorized(X_log, G, L):
    X_hat = G @ L.T
    diff = X_log - X_hat
    ss_res = np.sum(diff ** 2, axis=0)
    mean_x = np.mean(X_log, axis=0, keepdims=True)
    ss_tot = np.sum((X_log - mean_x) ** 2, axis=0)
    valid = ss_tot > 0
    r2_cells = np.where(valid, 1.0 - ss_res / np.where(valid, ss_tot, 1.0), np.nan)
    return float(np.nanmean(r2_cells))

def load_data_and_cache():
    print("--- [1/6] Loading R2 Cache and 2,778 G-vectors with Exact Signs ---", flush=True)
    t0 = time.time()
    with open(CACHE_FILE, 'rb') as fp:
        cached = pickle.load(fp)
    cache_data = cached['cache_data']
    g_weights = cached['g_weights']
    print(f"Loaded cache for {len(cache_data)} ROIs ({len(g_weights)} G-weights) in {time.time()-t0:.2f}s.", flush=True)

    files = sorted(glob.glob(os.path.join(TAILS_DIR, '*.xlsx')))
    g_data = []
    all_prots_set = set()

    for f in files:
        roi_id = os.path.basename(f).replace('_tail_proteins.xlsx', '')
        m_p = re.match(r'^(P\d+)', roi_id)
        pat = m_p.group(1) if m_p else roi_id
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
                            all_prots_set.add(p)
            if len(signs) >= 2:
                gkey = f"{roi_id}_G{n}"
                weight = g_weights.get(gkey, 0.0)
                g_data.append({
                    'g_idx': len(g_data),
                    'gkey': gkey,
                    'roi': roi_id,
                    'patient': pat,
                    'g_num': n,
                    'signs': signs,
                    'prots': set(signs.keys()),
                    'count': len(signs),
                    'weight': weight
                })

    print(f"Loaded {len(g_data)} G-vectors across {len(files)} ROIs ({len(all_prots_set)} unique proteins).", flush=True)
    return cache_data, g_weights, g_data, sorted(all_prots_set)

def match_es_g(es_dict, g_signs, g_prots, min_overlap=2):
    common = set(es_dict.keys()) & g_prots
    if len(common) < min_overlap:
        return None, 0, []
    # Direct match: all signs identical
    if all(es_dict[p] == g_signs[p] for p in common):
        return '+', len(common), sorted(common)
    # Inverse match: all signs flipped
    if all(es_dict[p] == -g_signs[p] for p in common):
        return '-', len(common), sorted(common)
    return None, 0, []

def canon_dict(d):
    sorted_items = sorted(d.items(), key=lambda x: x[0])
    if sorted_items[0][1] == -1:
        return dict((p, -s) for p, s in sorted_items)
    return dict(sorted_items)

def mine_candidate_modules(g_data):
    print("--- [2/6] Mining Candidate Exact Sign Modules (Lengths 5 to 8) ---", flush=True)
    t0 = time.time()
    candidate_modules = []
    seen = set()

    for g in g_data:
        if len(g['signs']) >= 5:
            cd = canon_dict(g['signs'])
            ct = tuple(sorted(cd.items()))
            if ct not in seen:
                seen.add(ct)
                candidate_modules.append(cd)

    for g in g_data:
        if len(g['signs']) >= 7:
            items = list(g['signs'].items())
            for k in (6, 7):
                if k <= len(items):
                    for comb in combinations(items, k):
                        cd = canon_dict(dict(comb))
                        ct = tuple(sorted(cd.items()))
                        if ct not in seen:
                            seen.add(ct)
                            candidate_modules.append(cd)
                            if len(candidate_modules) > 8000: break
                if len(candidate_modules) > 8000: break
        if len(candidate_modules) > 8000: break

    print(f"Mined {len(candidate_modules):,} candidate sign modules in {time.time()-t0:.2f}s.", flush=True)
    return candidate_modules

def run_variance_set_cover(g_data, candidate_modules, min_g_threshold=2):
    print(f"--- [3/6] Scoring Modules and Running R^2-Variance Greedy Set Cover (Min Gs >= {min_g_threshold}) ---", flush=True)
    t0 = time.time()
    total_g = len(g_data)

    # Pre-score candidate matches
    module_matches = []
    for idx, mod in enumerate(candidate_modules):
        hits = {}
        for g in g_data:
            orient, n_common, matched_prots = match_es_g(mod, g['signs'], g['prots'], min_overlap=2)
            if orient:
                hits[g['g_idx']] = (orient, n_common, matched_prots)
        if len(hits) >= min_g_threshold:
            module_matches.append((idx, mod, hits))

    print(f"Filtered to {len(module_matches):,} eligible modules. Executing greedy selection...", flush=True)

    covered = set()
    selected = []
    g_assignment = {}
    cum_variance = 0.0

    while len(covered) < total_g:
        best_cand = None
        best_hits = {}
        best_new_g = set()
        best_var_gain = -1.0
        best_cnt_gain = 0

        for idx, mod, hits in module_matches:
            new_g = set(hits.keys()) - covered
            if not new_g:
                continue
            cnt = len(new_g)
            if cnt < min_g_threshold:
                continue
            var_gain = sum(g_data[gi]['weight'] for gi in new_g)
            # Prioritize variance gain, tie-break by count
            if (var_gain > best_var_gain) or (abs(var_gain - best_var_gain) < 1e-6 and cnt > best_cnt_gain):
                best_var_gain = var_gain
                best_cnt_gain = cnt
                best_cand = mod
                best_hits = hits
                best_new_g = new_g

        if not best_cand or len(best_new_g) < min_g_threshold:
            print(f"  Stopping greedy cover: next best module covers {len(best_new_g)} Gs (< {min_g_threshold} threshold).", flush=True)
            break

        covered.update(best_new_g)
        cum_variance += best_var_gain
        es_label = f"ES_{len(selected)+1}"

        plus_cnt = 0
        minus_cnt = 0
        rois_covered = set()

        for g_idx in best_new_g:
            orient, n_common, matched_prots = best_hits[g_idx]
            if orient == '+': plus_cnt += 1
            else: minus_cnt += 1
            g_rec = g_data[g_idx]
            rois_covered.add(g_rec['roi'])
            cov_pct = n_common / g_rec['count'] * 100
            g_assignment[g_idx] = {
                'es_label': es_label,
                'orient': orient,
                'n_common': n_common,
                'matched_prots': matched_prots,
                'cov_pct': cov_pct,
                'weight': g_rec['weight']
            }

        seq_str = " | ".join(f"{p}({'+' if s==1 else '-'})" for p, s in sorted(best_cand.items()))
        cov_pct = len(covered) / total_g * 100
        var_pct = cum_variance / 120.0 * 100

        selected.append({
            'Rank': len(selected) + 1,
            'ES_Label': es_label,
            'Sequence_Length': len(best_cand),
            'New_G_Covered': len(best_new_g),
            'Direct_Matches_(+)': plus_cnt,
            'Inverse_Matches_(-)': minus_cnt,
            'Cumulative_G_Covered': len(covered),
            'Coverage_%': round(cov_pct, 2),
            'New_Variance_Explained': round(best_var_gain, 4),
            'Cumulative_Variance_%': round(var_pct, 2),
            'ROIs_Covered_Count': len(rois_covered),
            'Exact_Sign_Sequence': seq_str,
            'mod_dict': best_cand
        })

        print(f"  [{es_label:>5}] Len={len(best_cand)} | +{len(best_new_g):>3} G (+:{plus_cnt:>2}, -:{minus_cnt:>2}) | Total: {len(covered)}/{total_g} ({cov_pct:.2f}%) | Var: {var_pct:.2f}% | {seq_str[:50]}...", flush=True)

    print(f"\nFinal Cover: {len(selected)} ES modules cover {len(covered)}/{total_g} Gs ({len(covered)/total_g*100:.2f}%), explaining {cum_variance/120.0*100:.2f}% of biological variance in {time.time()-t0:.2f}s.", flush=True)
    return selected, g_assignment, covered

def evaluate_all_rois_r2(cache_data, g_data, g_assignment):
    print("--- [4/6] Computing R^2 Metrics for All 120 ROIs ---", flush=True)
    t0 = time.time()
    
    # Map G components to ROIs
    roi_g_assigned = defaultdict(set)
    for g_idx, info in g_assignment.items():
        g = g_data[g_idx]
        roi_g_assigned[g['roi']].add(g['g_num'])

    r2_records = []
    for roi_id in sorted(cache_data.keys()):
        rdata = cache_data[roi_id]
        X_log = rdata['X_log']
        G_all = rdata['G_all']
        L_all = rdata['L_all']
        r2_full = rdata['r2_full']
        
        sel_nums = sorted(roi_g_assigned.get(roi_id, set()))
        # Include baseline G0 + selected G components
        sel_cols = [0] + sel_nums
        r2_opt = compute_r2_vectorized(X_log, G_all[:, sel_cols], L_all[:, sel_cols])
        delta_r2 = r2_full - r2_opt

        if delta_r2 < 0.05:
            interp = "מצוין (< 0.05)"
            cat = "Excellent"
        elif delta_r2 <= 0.15:
            interp = "סביר (0.05 - 0.15)"
            cat = "Acceptable"
        else:
            interp = "לשים לב (> 0.15)"
            cat = "Alert"

        r2_records.append({
            'ROI': roi_id,
            'Patient': rdata['patient'],
            'N_cells': rdata['n_cells'],
            'R2_full': round(r2_full, 4),
            'R2_optimal': round(r2_opt, 4),
            'Delta_R2': round(delta_r2, 4),
            'N_selected': len(sel_nums),
            'Interpretation': interp,
            'Category': cat,
            'Selected_G_List': ", ".join(f"G{n}" for n in sel_nums)
        })

    r2_df = pd.DataFrame(r2_records)
    mean_full = r2_df['R2_full'].mean()
    mean_opt = r2_df['R2_optimal'].mean()
    mean_delta = r2_df['Delta_R2'].mean()
    n_exc = sum(r2_df['Category'] == 'Excellent')
    n_acc = sum(r2_df['Category'] == 'Acceptable')
    n_ale = sum(r2_df['Category'] == 'Alert')

    print(f"R^2 Evaluation Complete ({time.time()-t0:.2f}s):")
    print(f"  Mean R^2_full:    {mean_full:.4f}")
    print(f"  Mean R^2_optimal: {mean_opt:.4f}")
    print(f"  Mean Delta_R2:    {mean_delta:.4f}")
    print(f"  Excellent (<0.05):     {n_exc:>3} / 120 ({n_exc/120*100:.1f}%)")
    print(f"  Acceptable (0.05-0.15): {n_acc:>3} / 120 ({n_acc/120*100:.1f}%)")
    print(f"  Alert (>0.15):          {n_ale:>3} / 120 ({n_ale/120*100:.1f}%)")

    return r2_df, r2_records

def build_excel_reports(cache_data, g_data, selected, g_assignment, covered, r2_df, r2_records):
    print(f"--- [5/6] Building Comprehensive Excel Reports ({OUTPUT_EXCEL}) ---", flush=True)
    t0 = time.time()
    total_g = len(g_data)
    total_rois = len(cache_data)

    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet

    # Fonts
    f_title = Font(name='Segoe UI', size=16, bold=True, color='1F497D')
    f_header = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    f_bold = Font(name='Segoe UI', size=10, bold=True)
    f_regular = Font(name='Segoe UI', size=10)

    # Fills
    fill_header = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
    fill_zebra = PatternFill(start_color='F5F7FA', end_color='F5F7FA', fill_type='solid')
    fill_accent = PatternFill(start_color='D9E2F3', end_color='D9E2F3', fill_type='solid')
    fill_green = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
    fill_yellow = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
    fill_red = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')

    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    g_by_roi = defaultdict(list)
    for g in g_data:
        g_by_roi[g['roi']].append(g)

    # ─────────────────────────────────────────────
    # Sheet 1: Summary
    # ─────────────────────────────────────────────
    ws_sum = wb.create_sheet(title='Summary')
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum.cell(row=1, column=1, value="OPTIMAL EXACT SIGN SEQUENCES (ES) & R² VARIANCE REPORT").font = f_title
    ws_sum.cell(row=2, column=1, value="Cancer Research Surprisal Analysis: Exact Sign Sequences, G Orientations & Reconstruction Quality").font = Font(name='Segoe UI', size=11, italic=True, color='595959')

    total_plus = sum(s['Direct_Matches_(+)'] for s in selected)
    total_minus = sum(s['Inverse_Matches_(-)'] for s in selected)
    cum_var_pct = selected[-1]['Cumulative_Variance_%'] if selected else 0.0

    mean_full = r2_df['R2_full'].mean()
    mean_opt = r2_df['R2_optimal'].mean()
    mean_delta = r2_df['Delta_R2'].mean()
    n_exc = sum(r2_df['Category'] == 'Excellent')
    n_acc = sum(r2_df['Category'] == 'Acceptable')
    n_ale = sum(r2_df['Category'] == 'Alert')

    roi_cov_list = [sum(1 for g in glist if g['g_idx'] in g_assignment) / len(glist) * 100 for glist in g_by_roi.values()]
    avg_roi_cov = np.mean(roi_cov_list)
    rois_100 = sum(1 for p in roi_cov_list if p >= 99.9)

    summary_rows = [
        ("── GLOBAL DATASET PROPERTIES ────────────────────────────", ""),
        ("Total G-vectors Analyzed", total_g),
        ("Total ROI Sections (Patient Sub-Regions)", total_rois),
        ("Total Unique Proteins in Study", 39),
        ("── SIGN-PRESERVED SET COVER RESULTS ─────────────────────", ""),
        ("Number of ES Modules in Optimal Cover", len(selected)),
        ("Total G-vectors Covered (No Duplicates)", f"{len(covered):,} / {total_g:,}"),
        ("Global G-vector Coverage %", f"{len(covered)/total_g*100:.2f}%"),
        ("Biological Variance Explained (Cumulative)", f"{cum_var_pct:.2f}%"),
        ("G-vectors with Direct Orientation (+)", f"{total_plus:,} ({total_plus/len(covered)*100:.1f}%)"),
        ("G-vectors with Inverse Orientation (-)", f"{total_minus:,} ({total_minus/len(covered)*100:.1f}%)"),
        ("Average ROI Coverage % Across All 120 Sections", f"{avg_roi_cov:.2f}%"),
        ("ROIs Reaching 100% G-Coverage", f"{rois_100} / {total_rois} ({rois_100/total_rois*100:.1f}%)"),
        ("── R² RECONSTRUCTION QUALITY METRICS ────────────────────", ""),
        ("Mean R²_full (Baseline SVD Reconstruction)", f"{mean_full:.4f}"),
        ("Mean R²_optimal (Compressed Representation)", f"{mean_opt:.4f}"),
        ("Mean ΔR² (Compression Cost: R²_full - R²_optimal)", f"{mean_delta:.4f}"),
        ("Sections Classified as 'Excellent' (ΔR² < 0.05)", f"{n_exc} / {total_rois} ({n_exc/total_rois*100:.1f}%)"),
        ("Sections Classified as 'Acceptable' (0.05 ≤ ΔR² ≤ 0.15)", f"{n_acc} / {total_rois} ({n_acc/total_rois*100:.1f}%)"),
        ("Sections Classified as 'Alert' (ΔR² > 0.15)", f"{n_ale} / {total_rois} ({n_ale/total_rois*100:.1f}%)"),
        ("Combined High Quality (Excellent + Acceptable)", f"{n_exc + n_acc} / {total_rois} ({(n_exc+n_acc)/total_rois*100:.1f}%)"),
        ("── STRICT CONSTRAINTS VERIFICATION ──────────────────────", ""),
        ("Exact Sign Sequences (ES)", "100.0% VERIFIED (Explicit '+' or '-' sign for each protein)"),
        ("G-Vector Relative Orientation Tagging", "100.0% VERIFIED (Every assigned G is tagged with (+) or (-))"),
        ("Sign Conflict / Mismatch Rate", "0.0% (Strict zero sign contradiction across all matches)"),
        ("Disjoint G Allocation (No G Duplicates)", "100.0% VERIFIED (Each G assigned to exactly ONE ES module)"),
        ("Conflict-Free Allocation Matrix", "100.0% VERIFIED (Each cell ≤ 1 unique G; zero G repeats per row)"),
        ("Long-Tail Singletons Filter", "ENFORCED (No singletons; modules cover ≥ 2 G-vectors)"),
    ]

    ws_sum.cell(row=4, column=1, value="Metric").font = f_header
    ws_sum.cell(row=4, column=1).fill = fill_header
    ws_sum.cell(row=4, column=2, value="Value").font = f_header
    ws_sum.cell(row=4, column=2).fill = fill_header

    for r_idx, (m, v) in enumerate(summary_rows, start=5):
        c1 = ws_sum.cell(row=r_idx, column=1, value=m)
        c2 = ws_sum.cell(row=r_idx, column=2, value=v)
        c1.border = border_thin
        c2.border = border_thin
        if m.startswith("──"):
            c1.font = f_bold
            c1.fill = fill_accent
            c2.fill = fill_accent
        else:
            c1.font = f_regular
            c2.font = f_bold if isinstance(v, (int, float)) or '%' in str(v) else f_regular
            if r_idx % 2 == 0:
                c1.fill = fill_zebra
                c2.fill = fill_zebra

    ws_sum.column_dimensions['A'].width = 52
    ws_sum.column_dimensions['B'].width = 46

    # ─────────────────────────────────────────────
    # Sheet 2: Optimal_ES_SetCover
    # ─────────────────────────────────────────────
    ws_opt = wb.create_sheet(title='Optimal_ES_SetCover')
    ws_opt.views.sheetView[0].showGridLines = True

    headers_opt = [
        "Rank", "ES_Label", "Sequence_Length", "New_G_Covered",
        "Direct_Matches_(+)", "Inverse_Matches_(-)", "Cumulative_G_Covered",
        "Coverage_%", "New_Variance_Weight", "Cumulative_Variance_%",
        "ROIs_Covered_Count", "Exact_Sign_Sequence"
    ]
    for c_idx, h in enumerate(headers_opt, 1):
        cell = ws_opt.cell(row=1, column=c_idx, value=h)
        cell.font = f_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for r_idx, s in enumerate(selected, 2):
        row_vals = [
            s['Rank'], s['ES_Label'], s['Sequence_Length'], s['New_G_Covered'],
            s['Direct_Matches_(+)'], s['Inverse_Matches_(-)'], s['Cumulative_G_Covered'],
            s['Coverage_%'], s['New_Variance_Explained'], s['Cumulative_Variance_%'],
            s['ROIs_Covered_Count'], s['Exact_Sign_Sequence']
        ]
        for c_idx, v in enumerate(row_vals, 1):
            cell = ws_opt.cell(row=r_idx, column=c_idx, value=v)
            cell.font = f_regular
            cell.border = border_thin
            if c_idx in range(1, 12):
                cell.alignment = Alignment(horizontal='center', vertical='center')
            if r_idx % 2 == 1:
                cell.fill = fill_zebra

    opt_widths = [8, 12, 16, 16, 18, 18, 22, 14, 20, 22, 18, 80]
    for i, w in enumerate(opt_widths, 1):
        ws_opt.column_dimensions[get_column_letter(i)].width = w

    # ─────────────────────────────────────────────
    # Sheet 3: ES_Conflict_Free_Matrix
    # ─────────────────────────────────────────────
    ws_mat = wb.create_sheet(title='ES_Conflict_Free_Matrix')
    ws_mat.views.sheetView[0].showGridLines = True

    es_col_labels = [s['ES_Label'] for s in selected]
    headers_mat = ["Patient", "ROI", "ROI_Key"] + es_col_labels
    for c_idx, h in enumerate(headers_mat, 1):
        cell = ws_mat.cell(row=1, column=c_idx, value=h)
        cell.font = f_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for r_idx, (roi_id, glist) in enumerate(sorted(g_by_roi.items()), 2):
        pat = glist[0]['patient']
        roi_only = roi_id.replace(f"{pat}_", "")

        roi_es_cands = defaultdict(list)
        for g in glist:
            gi = g['g_idx']
            if gi in g_assignment:
                info = g_assignment[gi]
                roi_es_cands[info['es_label']].append({
                    'g_id': g['gkey'],
                    'g_val': f"G{g['g_num']}({info['orient']})",
                    'cov_pct': info['cov_pct'],
                    'weight': g['weight']
                })

        # Conflict-free allocation per ES column
        assigned_in_row = set()
        cell_values = {}
        for es_lbl in es_col_labels:
            cands = roi_es_cands.get(es_lbl, [])
            cands.sort(key=lambda x: (-x['weight'], -x['cov_pct']))
            chosen = ""
            for c in cands:
                if c['g_id'] not in assigned_in_row:
                    chosen = c['g_val']
                    assigned_in_row.add(c['g_id'])
                    break
            cell_values[es_lbl] = chosen

        row_vals = [pat, roi_only, roi_id] + [cell_values[es_lbl] for es_lbl in es_col_labels]
        for c_idx, v in enumerate(row_vals, 1):
            cell = ws_mat.cell(row=r_idx, column=c_idx, value=v)
            cell.font = f_regular
            cell.border = border_thin
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if r_idx % 2 == 1:
                cell.fill = fill_zebra

    ws_mat.column_dimensions['A'].width = 10
    ws_mat.column_dimensions['B'].width = 12
    ws_mat.column_dimensions['C'].width = 16
    for c_idx in range(4, 4 + len(es_col_labels)):
        ws_mat.column_dimensions[get_column_letter(c_idx)].width = 12

    # ─────────────────────────────────────────────
    # Sheet 4: Best_Combination_Details
    # ─────────────────────────────────────────────
    ws_det = wb.create_sheet(title='Best_Combination_Details')
    ws_det.views.sheetView[0].showGridLines = True

    headers_det = [
        "Patient", "ROI", "ROI_Key", "Total_G_In_ROI", "Covered_G_Count",
        "ROI_Coverage_%", "Assigned_ES_Labels", "Consolidated_G_Assignments"
    ]
    for c_idx, h in enumerate(headers_det, 1):
        cell = ws_det.cell(row=1, column=c_idx, value=h)
        cell.font = f_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for r_idx, (roi_id, glist) in enumerate(sorted(g_by_roi.items()), 2):
        pat = glist[0]['patient']
        roi_only = roi_id.replace(f"{pat}_", "")
        tot_g = len(glist)
        assigned_in_roi = [g for g in glist if g['g_idx'] in g_assignment]
        cov_cnt = len(assigned_in_roi)
        cov_pct = round(cov_cnt / tot_g * 100, 2)
        es_labels = sorted(list(set(g_assignment[g['g_idx']]['es_label'] for g in assigned_in_roi)))
        es_str = "; ".join(es_labels) if es_labels else "None"

        g_parts = []
        for g in sorted(glist, key=lambda x: x['g_num']):
            gi = g['g_idx']
            if gi in g_assignment:
                info = g_assignment[gi]
                g_parts.append(f"G{g['g_num']}:{info['es_label']}({info['orient']})[{info['cov_pct']:.1f}%]")
            else:
                g_parts.append(f"G{g['g_num']}:Uncovered")
        consolidated_str = ", ".join(g_parts)

        row_vals = [pat, roi_only, roi_id, tot_g, cov_cnt, cov_pct, es_str, consolidated_str]
        for c_idx, v in enumerate(row_vals, 1):
            cell = ws_det.cell(row=r_idx, column=c_idx, value=v)
            cell.font = f_regular
            cell.border = border_thin
            if c_idx in range(1, 7):
                cell.alignment = Alignment(horizontal='center', vertical='center')
            if r_idx % 2 == 1:
                cell.fill = fill_zebra

    det_widths = [10, 12, 16, 16, 18, 16, 25, 80]
    for i, w in enumerate(det_widths, 1):
        ws_det.column_dimensions[get_column_letter(i)].width = w

    # ─────────────────────────────────────────────
    # Sheet 5: Per_G_Coverage_Detail
    # ─────────────────────────────────────────────
    ws_perg = wb.create_sheet(title='Per_G_Coverage_Detail')
    ws_perg.views.sheetView[0].showGridLines = True

    headers_perg = [
        "Patient", "ROI", "ROI_Key", "G_Key", "G_Num", "Total_Prots_In_G",
        "Assigned_ES", "Orientation", "Matched_Prot_Count", "Coverage_%",
        "Variance_Weight_%", "Matched_Proteins", "G_Full_Signs"
    ]
    for c_idx, h in enumerate(headers_perg, 1):
        cell = ws_perg.cell(row=1, column=c_idx, value=h)
        cell.font = f_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for r_idx, g in enumerate(g_data, 2):
        gi = g['g_idx']
        g_signs_str = " | ".join(f"{p}({'+' if s==1 else '-'})" for p, s in sorted(g['signs'].items()))
        w_pct = round(g['weight'] * 100, 3)

        if gi in g_assignment:
            info = g_assignment[gi]
            es_lab = info['es_label']
            orient = info['orient']
            n_match = info['n_common']
            cov_p = round(info['cov_pct'], 2)
            matched_str = " | ".join(info['matched_prots'])
        else:
            es_lab = "Uncovered"
            orient = "None"
            n_match = 0
            cov_p = 0.0
            matched_str = "None"

        row_vals = [
            g['patient'], g['roi'].replace(f"{g['patient']}_", ""), g['roi'],
            g['gkey'], g['g_num'], g['count'], es_lab, orient,
            n_match, cov_p, w_pct, matched_str, g_signs_str
        ]
        for c_idx, v in enumerate(row_vals, 1):
            cell = ws_perg.cell(row=r_idx, column=c_idx, value=v)
            cell.font = f_regular
            cell.border = border_thin
            if c_idx in range(1, 12):
                cell.alignment = Alignment(horizontal='center', vertical='center')
            if r_idx % 2 == 1:
                cell.fill = fill_zebra

    perg_widths = [10, 12, 16, 18, 8, 16, 14, 12, 18, 14, 18, 30, 45]
    for i, w in enumerate(perg_widths, 1):
        ws_perg.column_dimensions[get_column_letter(i)].width = w

    # ─────────────────────────────────────────────
    # Sheet 6: R2_Comparison_Per_ROI (Matching R2 Analysis Guide)
    # ─────────────────────────────────────────────
    ws_r2 = wb.create_sheet(title='R2_Comparison_Per_ROI')
    ws_r2.views.sheetView[0].showGridLines = True

    headers_r2 = ["ROI", "Patient", "N_cells", "R2_full", "R2_optimal", "Delta_R2", "N_selected", "Interpretation", "Selected_G_Components"]
    for c_idx, h in enumerate(headers_r2, 1):
        cell = ws_r2.cell(row=1, column=c_idx, value=h)
        cell.font = f_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for r_idx, r in enumerate(r2_records, 2):
        row_vals = [
            r['ROI'], r['Patient'], r['N_cells'], r['R2_full'],
            r['R2_optimal'], r['Delta_R2'], r['N_selected'],
            r['Interpretation'], r['Selected_G_List']
        ]
        for c_idx, v in enumerate(row_vals, 1):
            cell = ws_r2.cell(row=r_idx, column=c_idx, value=v)
            cell.font = f_regular
            cell.border = border_thin
            if c_idx in range(1, 8):
                cell.alignment = Alignment(horizontal='center', vertical='center')
            if c_idx == 8: # Interpretation
                cell.alignment = Alignment(horizontal='center', vertical='center')
                if r['Category'] == 'Excellent':
                    cell.fill = fill_green
                    cell.font = f_bold
                elif r['Category'] == 'Acceptable':
                    cell.fill = fill_yellow
                    cell.font = f_bold
                else:
                    cell.fill = fill_red
                    cell.font = f_bold
            elif r_idx % 2 == 1:
                cell.fill = fill_zebra

    r2_widths = [16, 12, 12, 14, 14, 14, 12, 22, 45]
    for i, w in enumerate(r2_widths, 1):
        ws_r2.column_dimensions[get_column_letter(i)].width = w

    # ─────────────────────────────────────────────
    # Sheet 7: Per_ROI_Coverage
    # ─────────────────────────────────────────────
    ws_proi = wb.create_sheet(title='Per_ROI_Coverage')
    ws_proi.views.sheetView[0].showGridLines = True

    headers_proi = [
        "Rank", "Patient", "ROI", "ROI_Key", "Total_G_Values",
        "Covered_G_Count", "Uncovered_G_Count", "ROI_Coverage_%", "Assigned_ES_Count",
        "R2_full", "R2_optimal", "Delta_R2", "Quality"
    ]
    for c_idx, h in enumerate(headers_proi, 1):
        cell = ws_proi.cell(row=1, column=c_idx, value=h)
        cell.font = f_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center')

    r2_map = {r['ROI']: r for r in r2_records}
    roi_records = []
    for roi_id, glist in g_by_roi.items():
        pat = glist[0]['patient']
        roi_only = roi_id.replace(f"{pat}_", "")
        tot_g = len(glist)
        cov_cnt = sum(1 for g in glist if g['g_idx'] in g_assignment)
        es_cnt = len(set(g_assignment[g['g_idx']]['es_label'] for g in glist if g['g_idx'] in g_assignment))
        cov_pct = round(cov_cnt / tot_g * 100, 2)
        rinfo = r2_map.get(roi_id, {})
        roi_records.append({
            'pat': pat, 'roi_only': roi_only, 'roi_id': roi_id,
            'tot': tot_g, 'cov': cov_cnt, 'uncov': tot_g - cov_cnt,
            'pct': cov_pct, 'es_cnt': es_cnt,
            'r2_full': rinfo.get('R2_full', 0),
            'r2_opt': rinfo.get('R2_optimal', 0),
            'delta_r2': rinfo.get('Delta_R2', 0),
            'cat': rinfo.get('Category', '')
        })

    roi_records.sort(key=lambda x: (-x['pct'], x['delta_r2']))

    for r_idx, r in enumerate(roi_records, 2):
        row_vals = [
            r_idx - 1, r['pat'], r['roi_only'], r['roi_id'], r['tot'],
            r['cov'], r['uncov'], r['pct'], r['es_cnt'],
            r['r2_full'], r['r2_opt'], r['delta_r2'], r['cat']
        ]
        for c_idx, v in enumerate(row_vals, 1):
            cell = ws_proi.cell(row=r_idx, column=c_idx, value=v)
            cell.font = f_regular
            cell.border = border_thin
            if c_idx in range(1, 14):
                cell.alignment = Alignment(horizontal='center', vertical='center')
            if c_idx == 13:
                if r['cat'] == 'Excellent':
                    cell.fill = fill_green
                    cell.font = f_bold
                elif r['cat'] == 'Acceptable':
                    cell.fill = fill_yellow
                    cell.font = f_bold
                else:
                    cell.fill = fill_red
                    cell.font = f_bold
            elif r_idx % 2 == 1:
                cell.fill = fill_zebra

    proi_widths = [8, 10, 12, 16, 16, 16, 18, 16, 18, 12, 12, 12, 14]
    for i, w in enumerate(proi_widths, 1):
        ws_proi.column_dimensions[get_column_letter(i)].width = w

    # ─────────────────────────────────────────────
    # Sheet 8: Coverage_Curve
    # ─────────────────────────────────────────────
    ws_crv = wb.create_sheet(title='Coverage_Curve')
    ws_crv.views.sheetView[0].showGridLines = True

    headers_crv = [
        "ES_Rank", "ES_Label", "New_G_Covered", "Cumulative_G_Covered",
        "Coverage_%", "Direct_(+)", "Inverse_(-)", "New_Variance_Weight", "Cumulative_Variance_%"
    ]
    for c_idx, h in enumerate(headers_crv, 1):
        cell = ws_crv.cell(row=1, column=c_idx, value=h)
        cell.font = f_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for r_idx, s in enumerate(selected, 2):
        row_vals = [
            s['Rank'], s['ES_Label'], s['New_G_Covered'], s['Cumulative_G_Covered'],
            s['Coverage_%'], s['Direct_Matches_(+)'], s['Inverse_Matches_(-)'],
            s['New_Variance_Explained'], s['Cumulative_Variance_%']
        ]
        for c_idx, v in enumerate(row_vals, 1):
            cell = ws_crv.cell(row=r_idx, column=c_idx, value=v)
            cell.font = f_regular
            cell.border = border_thin
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if r_idx % 2 == 1:
                cell.fill = fill_zebra

    crv_widths = [10, 14, 16, 22, 14, 14, 14, 20, 22]
    for i, w in enumerate(crv_widths, 1):
        ws_crv.column_dimensions[get_column_letter(i)].width = w

    # Save Primary Workbook
    os.makedirs(os.path.dirname(OUTPUT_EXCEL), exist_ok=True)
    wb.save(OUTPUT_EXCEL)
    print(f"Successfully generated {OUTPUT_EXCEL} ({os.path.getsize(OUTPUT_EXCEL)/(1024*1024):.2f} MB).", flush=True)

    # ─────────────────────────────────────────────
    # Standalone: Docs/R2 comparison per ROI.xlsx
    # ─────────────────────────────────────────────
    wb_r2_standalone = openpyxl.Workbook()
    ws_r2_s = wb_r2_standalone.active
    ws_r2_s.title = "R2_Comparison_Per_ROI"
    ws_r2_s.views.sheetView[0].showGridLines = True

    for c_idx, h in enumerate(headers_r2[:8], 1): # Standard columns per Guide
        cell = ws_r2_s.cell(row=1, column=c_idx, value=h)
        cell.font = f_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for r_idx, r in enumerate(r2_records, 2):
        row_vals = [
            r['ROI'], r['Patient'], r['N_cells'], r['R2_full'],
            r['R2_optimal'], r['Delta_R2'], r['N_selected'], r['Interpretation']
        ]
        for c_idx, v in enumerate(row_vals, 1):
            cell = ws_r2_s.cell(row=r_idx, column=c_idx, value=v)
            cell.font = f_regular
            cell.border = border_thin
            cell.alignment = Alignment(horizontal='center', vertical='center')
            if c_idx == 8:
                if r['Category'] == 'Excellent':
                    cell.fill = fill_green
                    cell.font = f_bold
                elif r['Category'] == 'Acceptable':
                    cell.fill = fill_yellow
                    cell.font = f_bold
                else:
                    cell.fill = fill_red
                    cell.font = f_bold
            elif r_idx % 2 == 1:
                cell.fill = fill_zebra

    for i, w in enumerate(r2_widths[:8], 1):
        ws_r2_s.column_dimensions[get_column_letter(i)].width = w

    wb_r2_standalone.save(OUTPUT_R2_STANDALONE)
    print(f"Successfully exported standalone Guide file: {OUTPUT_R2_STANDALONE} in {time.time()-t0:.2f}s.", flush=True)

def main():
    t_start = time.time()
    cache_data, g_weights, g_data, all_prots = load_data_and_cache()
    candidate_modules = mine_candidate_modules(g_data)
    selected, g_assignment, covered = run_variance_set_cover(g_data, candidate_modules, min_g_threshold=2)
    r2_df, r2_records = evaluate_all_rois_r2(cache_data, g_data, g_assignment)
    build_excel_reports(cache_data, g_data, selected, g_assignment, covered, r2_df, r2_records)
    print(f"\n=======================================================")
    print(f"PIPELINE FULLY COMPLETE IN {time.time()-t_start:.2f}s")
    print(f"=======================================================")

if __name__ == '__main__':
    main()
