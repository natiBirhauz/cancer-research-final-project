"""
04_compare_anchor_strategies.py
-------------------------------
Step 4 of the Cancer Research TME Pipeline:
Evaluates 6 different greedy anchor strategies on the universe of all candidate ES combinations:
1. SVD Variance Driver
2. Max Frequency / Universality
3. Deep Specificity (Lengths 7-8 Core)
4. Biological Hallmark (Immune CD3+/CD8+ Seed)
5. Biological Hallmark (Stroma FAP+/SMA+ Seed)
6. Hybrid Multi-Objective (Variance x Count) — WINNER

Outputs:
- Docs/anchor_comparison_results.pkl
- Docs/Anchor_Strategies_Comparison.xlsx
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

CACHE_FILE = os.path.join('Docs', 'r2_cache.pkl')
TAILS_DIR = 'All_Tail_Proteins_Tables'
OUTPUT_PKL = os.path.join('Docs', 'anchor_comparison_results.pkl')
OUTPUT_XLSX = os.path.join('Docs', 'Anchor_Strategies_Comparison.xlsx')

def canon_dict(d):
    s = sorted(d.items(), key=lambda x: x[0])
    if s[0][1] == -1:
        return dict((p, -v) for p, v in s)
    return dict(s)

def match_es_g(es_dict, g_signs, g_prots, min_overlap=2):
    common = set(es_dict.keys()) & g_prots
    if len(common) < min_overlap:
        return None, 0, []
    if all(es_dict[p] == g_signs[p] for p in common):
        return '+', len(common), sorted(common)
    if all(es_dict[p] == -g_signs[p] for p in common):
        return '-', len(common), sorted(common)
    return None, 0, []

def compute_r2_vectorized(X_log, G, L):
    X_hat = G @ L.T
    diff = X_log - X_hat
    ss_res = np.sum(diff ** 2, axis=0)
    mean_x = np.mean(X_log, axis=0, keepdims=True)
    ss_tot = np.sum((X_log - mean_x) ** 2, axis=0)
    valid = ss_tot > 0
    r2_cells = np.where(valid, 1.0 - ss_res / np.where(valid, ss_tot, 1.0), np.nan)
    return float(np.nanmean(r2_cells))

def main():
    t0 = time.time()
    print("====================================================================")
    print(" [Step 4/5] Evaluating 6 Greedy Anchor Strategies & R² Optimization")
    print("====================================================================")

    with open(CACHE_FILE, 'rb') as fp:
        c = pickle.load(fp)
    cache_data = c['cache_data']
    g_weights = c['g_weights']

    files = sorted(glob.glob(os.path.join(TAILS_DIR, '*.xlsx')))
    g_data = []

    for f in files:
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
            if len(signs) >= 2:
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

    total_g = len(g_data)
    print(f"Loaded {total_g} G-vectors across 120 ROIs.")

    # Mining candidate modules
    print("Mining candidate ES modules across lengths 3 to 8...")
    candidate_modules = []
    seen = set()

    for g in g_data:
        if len(g['signs']) >= 3:
            cd = canon_dict(g['signs'])
            ct = tuple(sorted(cd.items()))
            if ct not in seen:
                seen.add(ct)
                candidate_modules.append(cd)

    for g in g_data:
        n = len(g['signs'])
        if n < 4: continue
        items = list(g['signs'].items())
        for k in (6, 7):
            if k < n:
                for combo in combinations(items, k):
                    cd = canon_dict(dict(combo))
                    ct = tuple(sorted(cd.items()))
                    if ct not in seen:
                        seen.add(ct)
                        candidate_modules.append(cd)

    print(f"Total candidate modules: {len(candidate_modules):,}")
    print("Mapping candidate modules to G-vectors...")

    module_matches = []
    for idx, mod in enumerate(candidate_modules):
        hits = {}
        tot_var = 0.0
        for g in g_data:
            orient, n_common, matched = match_es_g(mod, g['signs'], g['prots'], min_overlap=2)
            if orient is not None:
                hits[g['g_idx']] = (orient, n_common, matched)
                tot_var += g['weight']
        if len(hits) >= 2:
            module_matches.append({
                'idx': idx,
                'mod': mod,
                'len': len(mod),
                'hits': hits,
                'initial_count': len(hits),
                'initial_var': tot_var
            })

    print(f"Eligible modules with >= 2 G matches: {len(module_matches):,}")

    def evaluate_strategy(name, seed_filter, scoring_func, min_new_g=2):
        t_strat = time.time()
        covered = set()
        selected = []
        g_assignment = {}

        if seed_filter:
            cands = [m for m in module_matches if seed_filter(m)]
            if cands:
                best_seed = max(cands, key=lambda m: scoring_func(m, set()))
                best_new = set(best_seed['hits'].keys())
                covered.update(best_new)
                lbl = f"ES_{len(selected)+1}"
                for gi in best_new:
                    g_assignment[gi] = (lbl, best_seed['hits'][gi])
                selected.append((lbl, best_seed['mod'], best_new, best_seed['len']))

        while len(covered) < total_g:
            best_mod = None
            best_new = set()
            best_score = -1e9

            for m in module_matches:
                new_g = set(m['hits'].keys()) - covered
                if len(new_g) < min_new_g: continue
                score = scoring_func(m, new_g)
                if score > best_score:
                    best_score = score
                    best_new = new_g
                    best_mod = m

            if not best_mod or len(best_new) < min_new_g:
                break

            covered.update(best_new)
            lbl = f"ES_{len(selected)+1}"
            for gi in best_new:
                g_assignment[gi] = (lbl, best_mod['hits'][gi])
            selected.append((lbl, best_mod['mod'], best_new, best_mod['len']))

        # Compute R2
        roi_g_assigned = defaultdict(set)
        for gi in covered:
            g = g_data[gi]
            roi_g_assigned[g['roi']].add(g['g_num'])

        r2_diffs, r2_opts, r2_fulls = [], [], []
        for roi_id, rdata in cache_data.items():
            X_log = rdata['X_log']
            G_all = rdata['G_all']
            L_all = rdata['L_all']
            r2_f = rdata['r2_full']
            sel_nums = sorted(roi_g_assigned.get(roi_id, set()))
            sel_cols = [0] + sel_nums
            r2_o = compute_r2_vectorized(X_log, G_all[:, sel_cols], L_all[:, sel_cols])
            r2_fulls.append(r2_f)
            r2_opts.append(r2_o)
            r2_diffs.append(r2_f - r2_o)

        tot_var = sum(g_data[gi]['weight'] for gi in covered)

        return {
            'Anchor_Strategy': name,
            'ES_Count': len(selected),
            'Covered_Gs': len(covered),
            'Coverage_%': round(len(covered) / total_g * 100, 2),
            'Variance_%': round(tot_var / 120.0 * 100, 2),
            'Mean_R2_full': round(float(np.mean(r2_fulls)), 4),
            'Mean_R2_opt': round(float(np.mean(r2_opts)), 4),
            'Mean_Delta_R2': round(float(np.mean(r2_diffs)), 4),
            'Excellent_%': round(float(np.mean(np.array(r2_diffs) < 0.05) * 100), 1),
            'Acceptable_%': round(float(np.mean((np.array(r2_diffs) >= 0.05) & (np.array(r2_diffs) <= 0.15)) * 100), 1),
            'Alert_%': round(float(np.mean(np.array(r2_diffs) > 0.15) * 100), 1),
            'Exec_Time_s': round(time.time() - t_strat, 2),
            'selected': selected,
            'g_assignment': g_assignment
        }

    strategies = []
    print("\nRunning Strategy 1: Pure SVD Variance Driver...")
    strategies.append(evaluate_strategy(
        "1. SVD Variance Driver", None,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g)
    ))

    print("Running Strategy 2: Max Frequency (Universality)...")
    strategies.append(evaluate_strategy(
        "2. Max Frequency (Universality)", None,
        lambda m, new_g: len(new_g)
    ))

    print("Running Strategy 3: Deep Specificity (Lengths 7-8 Core)...")
    strategies.append(evaluate_strategy(
        "3. Deep Specificity (Lengths 7-8 Core)", None,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g) * (1.2 if m['len'] >= 7 else 1.0)
    ))

    print("Running Strategy 4: Hallmark: Immune (CD3+/CD8+ Seed)...")
    def immune_seed(m):
        d = m['mod']
        return ('CD3' in d and 'CD8' in d and d['CD3'] == 1 and d['CD8'] == 1)
    strategies.append(evaluate_strategy(
        "4. Hallmark: Immune (CD3+/CD8+ Seed)", immune_seed,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g)
    ))

    print("Running Strategy 5: Hallmark: Stroma (FAP+/SMA+ Seed)...")
    def stroma_seed(m):
        d = m['mod']
        has_fap = ('FAP' in d and d['FAP'] == 1)
        has_sma = (('SMA' in d and d['SMA'] == 1) or ('aSMA' in d and d['aSMA'] == 1))
        return has_fap and has_sma
    strategies.append(evaluate_strategy(
        "5. Hallmark: Stroma (FAP+/SMA+ Seed)", stroma_seed,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g)
    ))

    print("Running Strategy 6: Hybrid (Variance x Count) — WINNER...")
    strategies.append(evaluate_strategy(
        "6. Hybrid (Variance x Count)", None,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g) * np.sqrt(len(new_g))
    ))

    with open(OUTPUT_PKL, 'wb') as fp:
        pickle.dump(strategies, fp)

    # Build Excel
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws_ov = wb.create_sheet('Anchor_Strategies_Overview')
    ws_ov.views.sheetView[0].showGridLines = True

    f_title = Font(name='Segoe UI', size=14, bold=True, color='1F497D')
    f_header = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    f_bold = Font(name='Segoe UI', size=10, bold=True)
    f_regular = Font(name='Segoe UI', size=10)
    fill_header = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
    fill_zebra = PatternFill(start_color='F5F7FA', end_color='F5F7FA', fill_type='solid')
    fill_winner = PatternFill(start_color='D9E2F3', end_color='D9E2F3', fill_type='solid')
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    ws_ov.cell(row=1, column=1, value="COMPARISON OF 6 GREEDY ANCHOR STRATEGIES & R² OPTIMIZATION").font = f_title
    headers_ov = [
        "Anchor_Strategy", "ES_Count", "Covered_Gs", "Coverage_%", "Variance_%",
        "Mean_R2_full", "Mean_R2_opt", "Mean_Delta_R2", "Excellent_% (<0.05)",
        "Acceptable_% (0.05-0.15)", "Alert_% (>0.15)", "Exec_Time_s"
    ]
    for c, h in enumerate(headers_ov, 1):
        cell = ws_ov.cell(row=3, column=c, value=h)
        cell.font = f_header; cell.fill = fill_header; cell.alignment = Alignment(horizontal='center')

    for r_idx, s in enumerate(strategies, 4):
        vals = [
            s['Anchor_Strategy'], s['ES_Count'], s['Covered_Gs'], f"{s['Coverage_%']:.2f}%", f"{s['Variance_%']:.2f}%",
            f"{s['Mean_R2_full']:.4f}", f"{s['Mean_R2_opt']:.4f}", f"{s['Mean_Delta_R2']:.4f}",
            f"{s['Excellent_%']:.1f}%", f"{s['Acceptable_%']:.1f}%", f"{s['Alert_%']:.1f}%", f"{s['Exec_Time_s']:.2f}s"
        ]
        is_winner = (r_idx == 4 + 5)
        for c, v in enumerate(vals, 1):
            cell = ws_ov.cell(row=r_idx, column=c, value=v)
            cell.font = f_bold if is_winner or c <= 2 else f_regular
            cell.border = border_thin
            if c >= 2: cell.alignment = Alignment(horizontal='center')
            if is_winner: cell.fill = fill_winner
            elif r_idx % 2 == 1: cell.fill = fill_zebra

    ov_widths = [35, 12, 14, 14, 14, 14, 14, 16, 20, 22, 16, 14]
    for i, w in enumerate(ov_widths, 1):
        ws_ov.column_dimensions[get_column_letter(i)].width = w

    # Sheet 2: Winning_Hybrid_32_ES
    s_win = strategies[5]
    ws_win = wb.create_sheet('Winning_Hybrid_32_ES')
    ws_win.views.sheetView[0].showGridLines = True
    headers_opt = [
        "Rank", "ES_Label", "Sequence_Length", "New_G_Covered",
        "Direct_Matches_(+)", "Inverse_Matches_(-)", "Cumulative_G_Covered",
        "Coverage_%", "New_Variance_Weight", "Cumulative_Variance_%",
        "ROIs_Covered_Count", "Exact_Sign_Sequence"
    ]
    for c, h in enumerate(headers_opt, 1):
        cell = ws_win.cell(row=1, column=c, value=h)
        cell.font = f_header; cell.fill = fill_header; cell.alignment = Alignment(horizontal='center')

    cum_g = 0
    cum_v = 0.0
    ga_win = s_win['g_assignment']
    for rank, (lbl, es_dict, g_set, l) in enumerate(s_win['selected'], 1):
        cum_g += len(g_set)
        var_gain = sum(g_data[gi]['weight'] for gi in g_set)
        cum_v += var_gain
        plus = sum(1 for gi in g_set if ga_win[gi][1][0] == '+')
        minus = sum(1 for gi in g_set if ga_win[gi][1][0] == '-')
        rois = len(set(g_data[gi]['roi'] for gi in g_set))
        seq = " | ".join(f"{p}({'+' if s==1 else '-'})" for p, s in sorted(es_dict.items()))
        row_vals = [
            rank, lbl, l, len(g_set), plus, minus, cum_g,
            round(cum_g / total_g * 100, 2), round(var_gain, 4), round(cum_v / 120.0 * 100, 2),
            rois, seq
        ]
        for c, v in enumerate(row_vals, 1):
            cell = ws_win.cell(row=rank + 1, column=c, value=v)
            cell.font = f_regular; cell.border = border_thin
            if c <= 11: cell.alignment = Alignment(horizontal='center')
            if rank % 2 == 1: cell.fill = fill_zebra

    win_widths = [8, 12, 16, 16, 18, 18, 20, 14, 18, 20, 18, 85]
    for i, w in enumerate(win_widths, 1):
        ws_win.column_dimensions[get_column_letter(i)].width = w

    wb.save(OUTPUT_XLSX)
    print(f"\nStep 4 Complete: Evaluated all 6 strategies in {time.time()-t0:.2f}s.\n")
    return 0

if __name__ == '__main__':
    main()
