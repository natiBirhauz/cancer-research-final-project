"""
run_full_evaluation_pipeline.py
--------------------------------
Generates the 4 master deliverable files for 'Data for algorithm evaluation':
1. Docs/Evaluation_Optimal_ES_Result.xlsx
2. Docs/Evaluation_Optimal_ES_Result_Signs_R2.xlsx
3. Docs/Evaluation_ES_Combinations.csv
4. Docs/Evaluation_Anchor_Strategies_Comparison.xlsx
"""

import os
import re
import glob
import time
from collections import defaultdict
from itertools import combinations
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from pipeline_engine import ingest_user_data, match_es_g, canon_dict, tuple_to_str, compute_r2_vectorized

def build_excel_deliverables(
    eval_dir='Data for algorithm evaluation',
    output_prefix='Evaluation_',
    output_dir='Docs'
):
    t0 = time.time()
    os.makedirs(output_dir, exist_ok=True)
    print("=" * 75)
    print("Running Full Evaluation Pipeline on 'Data for algorithm evaluation'")
    print("=" * 75)

    # 1. Ingest Data
    roi_info, g_data = ingest_user_data(file_paths=None, folder_path=eval_dir, threshold=0.05, min_overlap=2)
    total_g = len(g_data)
    total_rois = len(roi_info)
    print(f"Loaded {total_g} G-vectors across {total_rois} PDE samples.")
    for r_id, s in roi_info.items():
        print(f"  - {r_id}: {s['n_cells']} cells, {len(s['proteins'])} proteins, R2_full: {s['r2_full']:.4f}")

    # 2. Mine All Candidate ES Combinations (lengths 3 to 8)
    print("\nMining all candidate ES combinations (lengths 3 to 8)...")
    seen_combos = defaultdict(list)
    all_prots = set()

    for g in g_data:
        all_prots.update(g['prots'])
        items = list(g['signs'].items())
        n = len(items)
        for k in range(3, min(n + 1, 9)):
            for c in combinations(items, k):
                cd = canon_dict(dict(c))
                ct = tuple(sorted(cd.items()))
                # Check match orientation
                orient, _, _ = match_es_g(cd, g['signs'], g['prots'], min_overlap=2)
                if orient is not None:
                    seen_combos[ct].append(g['gkey'])

    print(f"Total unique ES candidate signatures mined: {len(seen_combos):,}")

    # Export ES_Combinations.csv
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
    
    es_combos_csv = os.path.join(output_dir, f"{output_prefix}ES_Combinations.csv")
    df_combos.to_csv(es_combos_csv, index=False)
    print(f"Exported: {es_combos_csv} ({len(df_combos):,} candidate sequences)")

    # 3. Filter eligible modules (>= 2 G matches)
    eligible_modules = []
    for ct, gkeys in seen_combos.items():
        unique_gkeys = set(gkeys)
        if len(unique_gkeys) >= 2:
            mod_dict = dict(ct)
            # Precompute matches with orientations
            hits = {}
            tot_var = 0.0
            for g in g_data:
                orient, n_common, matched = match_es_g(mod_dict, g['signs'], g['prots'], min_overlap=2)
                if orient is not None:
                    hits[g['g_idx']] = (orient, n_common, matched)
                    tot_var += g['weight']
            
            eligible_modules.append({
                'ct': ct,
                'mod': mod_dict,
                'len': len(mod_dict),
                'hits': hits,
                'initial_count': len(hits),
                'initial_var': tot_var
            })

    print(f"Eligible candidate modules with >= 2 G matches: {len(eligible_modules):,}")

    # 4. Evaluate Anchor Strategies
    def evaluate_strategy(name, seed_filter, scoring_func, min_new_g=2):
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
                if len(new_g) < min_new_g:
                    continue
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

        tot_var = sum(g_data[gi]['weight'] for gi in covered)

        return {
            'Anchor_Strategy': name,
            'ES_Count': len(selected),
            'Covered_Gs': len(covered),
            'Coverage_%': round(len(covered) / total_g * 100, 2),
            'Variance_%': round(tot_var / total_rois * 100, 2),
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

    # Evaluate all 6 strategies
    print("\nBenchmarking 6 Greedy Anchor Strategies:")
    strats = []
    # 1. Pure SVD Variance
    strats.append(evaluate_strategy(
        "1. SVD Variance Driver", None,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g)
    ))
    # 2. Max Frequency / Universality
    strats.append(evaluate_strategy(
        "2. Max Frequency / Universality", None,
        lambda m, new_g: float(len(new_g))
    ))
    # 3. Deep Specificity (Lengths 5-8 priority)
    strats.append(evaluate_strategy(
        "3. Deep Specificity (Lengths 5-8 Core)", None,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g) * (1.5 if m['len'] >= 5 else 1.0)
    ))
    # 4. Biological Seed 1: Epithelial / Carcinoma (EPCAM + EGFR co-expression)
    epcam_egfr = lambda m: ('EPCAM' in m['mod'] and 'EGFR' in m['mod'] and m['mod']['EPCAM'] == m['mod']['EGFR'])
    strats.append(evaluate_strategy(
        "4. Biological Seed (EPCAM+/EGFR+)", epcam_egfr,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g) * np.sqrt(len(new_g))
    ))
    # 5. Biological Seed 2: Stemness / Adhesion (CD44 + CD24 co-expression)
    cd44_cd24 = lambda m: ('CD44' in m['mod'] and 'CD24' in m['mod'])
    strats.append(evaluate_strategy(
        "5. Biological Seed (CD44/CD24)", cd44_cd24,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g) * np.sqrt(len(new_g))
    ))
    # 6. Hybrid Multi-Objective (Variance x sqrt(Count))
    strats.append(evaluate_strategy(
        "6. Hybrid Multi-Objective (Variance x Count)", None,
        lambda m, new_g: sum(g_data[gi]['weight'] for gi in new_g) * np.sqrt(len(new_g))
    ))

    for s in strats:
        print(f"  - {s['Anchor_Strategy']}: {s['ES_Count']} ES, {s['Covered_Gs']}/{total_g} Gs ({s['Coverage_%']}%), Var: {s['Variance_%']}%, R2_opt: {s['Mean_R2_opt']:.4f}, Delta_R2: {s['Mean_Delta_R2']:.4f}")

    # Write Anchor_Strategies_Comparison.xlsx
    anchor_xlsx = os.path.join(output_dir, f"{output_prefix}Anchor_Strategies_Comparison.xlsx")
    wb_anc = openpyxl.Workbook()
    ws_ov = wb_anc.active
    ws_ov.title = "Anchor_Strategies_Overview"

    headers_ov = ['Anchor_Strategy', 'ES_Count', 'Covered_Gs', 'Coverage_%', 'Variance_%', 'Mean_R2_full', 'Mean_R2_opt', 'Mean_Delta_R2', 'Excellent_%', 'Acceptable_%', 'Alert_%', 'Exec_Time_s']
    ws_ov.append(headers_ov)
    for s in strats:
        ws_ov.append([
            s['Anchor_Strategy'], s['ES_Count'], s['Covered_Gs'], s['Coverage_%'],
            s['Variance_%'], s['Mean_R2_full'], s['Mean_R2_opt'], s['Mean_Delta_R2'],
            s['Excellent_%'], s['Acceptable_%'], s['Alert_%'], s['Exec_Time_s']
        ])

    # Choose winning / primary strategy: Hybrid (or highest R2 / Coverage)
    hybrid_strat = [s for s in strats if 'Hybrid' in s['Anchor_Strategy']][0]
    ws_hyb = wb_anc.create_sheet(title=f"Optimal_Strategy_{len(hybrid_strat['selected'])}_ES")
    headers_es = ['Rank', 'ES_Label', 'Sequence_Length', 'New_G_Covered', 'Direct_Matches_(+)', 'Inverse_Matches_(-)', 'Cumulative_G_Covered', 'Coverage_%', 'New_Variance_Weight', 'Cumulative_Variance_%', 'ROIs_Covered_Count', 'Exact_Sign_Sequence']
    ws_hyb.append(headers_es)

    cum_g = 0
    cum_var = 0.0
    for rk, es in enumerate(hybrid_strat['selected'], 1):
        cum_g += len(es['new_g'])
        cum_var += es['var_gain']
        d_cnt = sum(1 for gi in es['new_g'] if hybrid_strat['g_assignment'][gi][1][0] == '+')
        i_cnt = sum(1 for gi in es['new_g'] if hybrid_strat['g_assignment'][gi][1][0] == '-')
        rois_cnt = len(set(g_data[gi]['roi'] for gi in es['new_g']))
        ws_hyb.append([
            rk, es['lbl'], es['len'], len(es['new_g']), d_cnt, i_cnt,
            cum_g, round(cum_g / total_g * 100, 2),
            round(es['var_gain'], 4), round(cum_var / total_rois * 100, 2),
            rois_cnt, tuple_to_str(es['mod'])
        ])

    wb_anc.save(anchor_xlsx)
    print(f"Exported: {anchor_xlsx}")

    # 5. Build Master Deliverable Excel (Optimal_ES_Result_Signs_R2.xlsx and Optimal_ES_Result.xlsx)
    def export_master_workbook(wb_path, strategy_res):
        wb = openpyxl.Workbook()
        # Sheet 1: Summary
        ws_sum = wb.active
        ws_sum.title = 'Summary'
        ws_sum.append(['OPTIMAL EXACT SIGN SEQUENCES (ES) & R2 VARIANCE REPORT'])
        ws_sum.append(['Surprisal Analysis: Exact Sign Sequences, G Orientations & Reconstruction Quality'])
        ws_sum.append([])
        ws_sum.append(['Metric', 'Value'])
        ws_sum.append(['Strategy Applied', strategy_res['Anchor_Strategy']])
        ws_sum.append(['Total Cohort Samples / ROIs', total_rois])
        ws_sum.append(['Total Cellular Microenvironment Vectors (G)', total_g])
        ws_sum.append(['Number of Optimal ES Modules', strategy_res['ES_Count']])
        ws_sum.append(['Total G-Vectors Covered', strategy_res['Covered_Gs']])
        ws_sum.append(['Universe G Coverage %', f"{strategy_res['Coverage_%']}%"])
        ws_sum.append(['Total Kinetic Variance Explained', f"{strategy_res['Variance_%']}%"])
        ws_sum.append(['Mean R2 Full Reconstruction', strategy_res['Mean_R2_full']])
        ws_sum.append(['Mean R2 Optimal Reconstruction', strategy_res['Mean_R2_opt']])
        ws_sum.append(['Mean Delta R2 Compression Penalty', strategy_res['Mean_Delta_R2']])
        ws_sum.append(['High Fidelity Samples (<0.05 Delta R2)', f"{strategy_res['Excellent_%']}%"])
        ws_sum.append(['Acceptable Fidelity Samples (0.05-0.15)', f"{strategy_res['Acceptable_%']}%"])
        ws_sum.append(['Noticeable Loss Samples (>0.15)', f"{strategy_res['Alert_%']}%"])

        # Sheet 2: Optimal_ES_SetCover
        ws_opt = wb.create_sheet(title='Optimal_ES_SetCover')
        ws_opt.append(headers_es)
        cum_g = 0
        cum_var = 0.0
        for rk, es in enumerate(strategy_res['selected'], 1):
            cum_g += len(es['new_g'])
            cum_var += es['var_gain']
            d_cnt = sum(1 for gi in es['new_g'] if strategy_res['g_assignment'][gi][1][0] == '+')
            i_cnt = sum(1 for gi in es['new_g'] if strategy_res['g_assignment'][gi][1][0] == '-')
            rois_cnt = len(set(g_data[gi]['roi'] for gi in es['new_g']))
            ws_opt.append([
                rk, es['lbl'], es['len'], len(es['new_g']), d_cnt, i_cnt,
                cum_g, round(cum_g / total_g * 100, 2),
                round(es['var_gain'], 4), round(cum_var / total_rois * 100, 2),
                rois_cnt, tuple_to_str(es['mod'])
            ])

        # Sheet 3: ES_Conflict_Free_Matrix
        ws_mat = wb.create_sheet(title='ES_Conflict_Free_Matrix')
        mat_cols = ['Sample', 'Cells', 'Total_G'] + [es['lbl'] for es in strategy_res['selected']]
        ws_mat.append(mat_cols)
        
        # Disjoint allocation mapping
        roi_es_used = defaultdict(set)
        for r_id, s in roi_info.items():
            row = [r_id, s['n_cells'], len([g for g in g_data if g['roi'] == r_id])]
            used_g = set()
            for es in strategy_res['selected']:
                lbl = es['lbl']
                cands = [gi for gi in es['new_g'] if g_data[gi]['roi'] == r_id and g_data[gi]['g_num'] not in used_g]
                if cands:
                    best_gi = max(cands, key=lambda gi: g_data[gi]['weight'])
                    g_item = g_data[best_gi]
                    orient = strategy_res['g_assignment'][best_gi][1][0]
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
            cov_gs = [g for g in sample_gs if g['g_idx'] in strategy_res['g_assignment']]
            assigned_es = sorted(set(strategy_res['g_assignment'][g['g_idx']][0] for g in cov_gs))
            
            assignments_str = []
            for g in sorted(cov_gs, key=lambda x: x['g_num']):
                lbl, hit = strategy_res['g_assignment'][g['g_idx']]
                orient = hit[0]
                pct = (hit[1] / g['count'] * 100) if g['count'] > 0 else 0
                assignments_str.append(f"G{g['g_num']}:{lbl}({orient})[{pct:.1f}%]")
            
            ws_det.append([
                r_id, s['n_cells'], len(sample_gs), len(cov_gs),
                round(len(cov_gs) / len(sample_gs) * 100, 2) if sample_gs else 0,
                "; ".join(assigned_es),
                ", ".join(assignments_str)
            ])

        # Sheet 5: Per_G_Coverage_Detail
        ws_perg = wb.create_sheet(title='Per_G_Coverage_Detail')
        ws_perg.append(['Sample', 'G_Key', 'G_Num', 'Total_Prots_In_G', 'Assigned_ES', 'Orientation', 'Matched_Prots_Count', 'Overlap_%', 'Variance_Weight', 'Status'])
        for g in g_data:
            if g['g_idx'] in strategy_res['g_assignment']:
                lbl, hit = strategy_res['g_assignment'][g['g_idx']]
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
        for r_id, (r2_f, r2_o, delta, sel_gs) in strategy_res['r2_per_sample'].items():
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
            cov_sample_gs = [g for g in sample_gs if g['g_idx'] in strategy_res['g_assignment']]
            var_cap = sum(g['weight'] for g in cov_sample_gs)
            ws_rois.append([
                r_id, s['n_cells'], len(sample_gs), len(cov_sample_gs),
                round(len(cov_sample_gs) / len(sample_gs) * 100, 2) if sample_gs else 0,
                round(var_cap * 100, 2)
            ])

        # Sheet 8: Coverage_Curve
        ws_curv = wb.create_sheet(title='Coverage_Curve')
        ws_curv.append(['Step', 'ES_Label', 'Cumulative_G_Covered', 'Cumulative_Coverage_%', 'Cumulative_Variance_%'])
        cum_g = 0
        cum_var = 0.0
        for rk, es in enumerate(strategy_res['selected'], 1):
            cum_g += len(es['new_g'])
            cum_var += es['var_gain']
            ws_curv.append([
                rk, es['lbl'], cum_g, round(cum_g / total_g * 100, 2), round(cum_var / total_rois * 100, 2)
            ])

        wb.save(wb_path)
        print(f"Exported: {wb_path}")

    # Export both master Excel deliverables
    out_signs_r2 = os.path.join(output_dir, f"{output_prefix}Optimal_ES_Result_Signs_R2.xlsx")
    out_optimal_es = os.path.join(output_dir, f"{output_prefix}Optimal_ES_Result.xlsx")

    export_master_workbook(out_signs_r2, hybrid_strat)
    export_master_workbook(out_optimal_es, hybrid_strat)

    print("=" * 75)
    print(f"ALL 4 MASTER DELIVERABLES CREATED SUCCESSFULLY IN {time.time()-t0:.2f}s!")
    print(f"  1. {out_optimal_es}")
    print(f"  2. {out_signs_r2}")
    print(f"  3. {es_combos_csv}")
    print(f"  4. {anchor_xlsx}")
    print("=" * 75)

if __name__ == '__main__':
    build_excel_deliverables()

