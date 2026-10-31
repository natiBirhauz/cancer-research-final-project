"""
optimal_es_finder.py
--------------------
Finds optimal ES combinations to maximize ROI coverage with minimum ES count.

Features:
1. Multi-anchor greedy search with backward pruning pass.
2. Conflict-free G assignment per ROI (at most ONE G-vector per cell, no duplicate Gs per ROI).
3. Exact per-G protein coverage calculation: matched / total_in_G * 100% (e.g. 3/8 = 37.5%).
4. Detailed per-ROI and global coverage metrics.
5. Rich 6-sheet Excel export:
   - Sheet 1: Combinations_Summary
   - Sheet 2: Optimal_Combination (Selected ES details)
   - Sheet 3: Best_Combination_Details (ROI x ES Conflict-Free Matrix)
   - Sheet 4: Per_G_Coverage_Detail (Every single G-vector with exact match & coverage %)
   - Sheet 5: Per_ROI_Coverage (Summary per ROI)
   - Sheet 6: Coverage_Curve (Step-by-step greedy progression)
"""

import os
import re
import time
from datetime import datetime
from collections import defaultdict
from typing import Dict, List, Set, Tuple
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TAILS_CSV  = os.path.join('Docs', 'All_Tails.csv')
ES_CSV     = os.path.join('Docs', 'ES_Combinations.csv')
OUTPUT_DIR = 'Docs'


def load_all_data(min_protein_count: int = 3):
    """
    Loads All_Tails.csv and ES_Combinations.csv, filtering for Protein_Count >= min_protein_count.
    Builds inverted index g_to_es for rapid greedy evaluation.
    """
    print(f"Loading data from {TAILS_CSV} and {ES_CSV} (min_protein_count={min_protein_count}) ...")
    df_tails = pd.read_csv(TAILS_CSV)
    df_es    = pd.read_csv(ES_CSV)

    # Filter Protein_Count >= min_protein_count
    df_es = df_es[df_es['Protein_Count'] >= min_protein_count].reset_index(drop=True)

    g_data = {}
    tail_g_map = defaultdict(set)
    for _, row in df_tails.iterrows():
        gk = row['G_Key']
        roi = row['ROI_Key']
        prots = frozenset(x.strip() for x in str(row['Protein_Set']).split(' | ') if x.strip())
        g_data[gk] = {
            'roi': roi,
            'patient': row['Patient'],
            'g_num': int(row['G_Num']),
            'prots': prots,
            'count': len(prots),
            'prots_str': str(row['Protein_Set'])
        }
        tail_g_map[roi].add(gk)

    universe = set(g_data.keys())

    # Build active ES dictionary and inverted index
    es_data = {}
    g_to_es = defaultdict(set)
    for _, row in df_es.iterrows():
        es_id = row['ES_ID']
        prots = frozenset(x.strip() for x in str(row['Protein_Set']).split(' | ') if x.strip())
        gks = set(row['Participating_G_Values'].split(' | ')) & universe
        if gks:
            es_data[es_id] = {
                'es_id': es_id,
                'prots': prots,
                'protein_count': int(row['Protein_Count']),
                'unique_patients': int(row['Unique_Patients_Count']),
                'rois_covered_count': int(row['ROIs_Covered_Count']),
                'gks': gks,
                'g_count': len(gks),
                'avg_prot_cov': float(row['Avg_Protein_Coverage_Per_G_%']),
                'protein_set_str': str(row['Protein_Set']),
                'tails_detail': str(row['Tails_Detail'])
            }
            for gk in gks:
                g_to_es[gk].add(es_id)

    return g_data, tail_g_map, universe, es_data, df_es, g_to_es


def run_greedy_search(
    es_data: dict,
    g_to_es: dict,
    tail_g_map: dict,
    universe: set,
    anchor_es_id: str = None,
    min_marginal_gain: int = 6,
    max_es_count: int = 100
) -> Tuple[List[str], List[dict]]:
    """
    Fast greedy search maximizing unique G coverage across ROIs.
    Stops when marginal gain < min_marginal_gain or max_es_count reached (eliminating the long tail).
    """
    selected = []
    covered_g = set()
    uncovered_g = set(universe)
    snapshots = []
    total_rois = len(tail_g_map)

    # If anchor specified
    if anchor_es_id and anchor_es_id in es_data:
        selected.append(anchor_es_id)
        new_g = es_data[anchor_es_id]['gks'] & uncovered_g
        covered_g.update(new_g)
        uncovered_g -= new_g

        roi_cov_pcts = [
            len(tail_g_map[r] & covered_g) / len(tail_g_map[r]) * 100
            for r in tail_g_map
        ]
        avg_roi_cov = float(np.mean(roi_cov_pcts))
        rois_with_cov = sum(1 for p in roi_cov_pcts if p > 0)

        snapshots.append({
            'n_ES': len(selected),
            'Added_ES_ID': anchor_es_id,
            'Protein_Count': es_data[anchor_es_id]['protein_count'],
            'Protein_Set': es_data[anchor_es_id]['protein_set_str'],
            'Delta_Unique_G': len(new_g),
            'Cumulative_Unique_G': len(covered_g),
            'Universe_Coverage_%': round(len(covered_g) / len(universe) * 100, 2),
            'ROIs_Covered': rois_with_cov,
            'Avg_ROI_Coverage_%': round(avg_roi_cov, 2),
        })

    while uncovered_g and len(selected) < max_es_count:
        # Candidate ES are those covering at least one uncovered G-vector
        cands = set()
        for gk in uncovered_g:
            cands.update(g_to_es[gk])
        cands -= set(selected)
        if not cands:
            break

        best_es = None
        best_gain = 0
        best_prot_cov = 0.0

        # Prioritize candidates with >= 2 patients first
        pool = [eid for eid in cands if es_data[eid]['unique_patients'] >= 2]
        if not pool:
            pool = list(cands)

        for es_id in pool:
            gain = len(es_data[es_id]['gks'] & uncovered_g)
            if gain > best_gain:
                best_gain = gain
                best_es = es_id
                best_prot_cov = es_data[es_id]['avg_prot_cov']
            elif gain == best_gain and gain > 0:
                # Tie-breaking: prefer higher average protein coverage
                if es_data[es_id]['avg_prot_cov'] > best_prot_cov:
                    best_es = es_id
                    best_prot_cov = es_data[es_id]['avg_prot_cov']

        # Stopping condition: cut off long tail
        if best_gain < min_marginal_gain or best_es is None:
            break

        selected.append(best_es)
        new_g = es_data[best_es]['gks'] & uncovered_g
        covered_g.update(new_g)
        uncovered_g -= new_g

        # ROI coverage %
        roi_cov_pcts = [
            len(tail_g_map[r] & covered_g) / len(tail_g_map[r]) * 100
            for r in tail_g_map
        ]
        avg_roi_cov = float(np.mean(roi_cov_pcts))
        rois_with_cov = sum(1 for p in roi_cov_pcts if p > 0)

        snapshots.append({
            'n_ES': len(selected),
            'Added_ES_ID': best_es,
            'Protein_Count': es_data[best_es]['protein_count'],
            'Protein_Set': es_data[best_es]['protein_set_str'],
            'Delta_Unique_G': len(new_g),
            'Cumulative_Unique_G': len(covered_g),
            'Universe_Coverage_%': round(len(covered_g) / len(universe) * 100, 2),
            'ROIs_Covered': rois_with_cov,
            'Avg_ROI_Coverage_%': round(avg_roi_cov, 2),
        })

    return selected, snapshots


def backward_pruning(combo: List[str], es_data: dict) -> List[str]:
    """
    Removes redundant ES whose elimination does not reduce unique G-vector coverage.
    """
    pruned = list(combo)
    changed = True
    while changed:
        changed = False
        for es in list(pruned):
            others_g = set()
            for o in pruned:
                if o != es:
                    others_g.update(es_data[o]['gks'])
            unique_contrib = es_data[es]['gks'] - others_g
            if len(unique_contrib) == 0:
                pruned.remove(es)
                changed = True
                break
    return pruned


def assign_conflict_free_for_combo(combo: List[str], tail_g_map: dict, g_data: dict, es_data: dict):
    """
    Assigns conflict-free G-values per ROI for the selected combination.
    Each ROI cell receives at most ONE G-vector.
    No duplicate G-vectors across ES for the same ROI.

    Returns:
      matrix_rows: list of dicts for the ROI x ES matrix table
      assigned_per_g: {gk: es_id} indicating which ES was assigned each G
    """
    matrix_rows = []
    assigned_per_g = {}

    for roi in sorted(tail_g_map.keys()):
        total_gs = tail_g_map[roi]
        assigned_in_roi = set()
        row_cells = {
            'ROI': roi,
            'Patient': roi.split('_')[0],
            'Total_G_in_ROI': len(total_gs),
        }

        # Build candidate Gs per ES in this ROI
        cands_per_es = []
        for es in combo:
            matching = [gk for gk in (es_data[es]['gks'] & total_gs)]
            cands_per_es.append((es, matching))

        # Sort by fewest candidates first (most constrained first)
        cands_per_es.sort(key=lambda x: len(x[1]))

        roi_assigned_covs = []
        for es, matching in cands_per_es:
            avail = [gk for gk in matching if gk not in assigned_in_roi]
            if avail:
                # Pick available G that maximizes protein coverage (fewest total proteins in G)
                best_gk = min(avail, key=lambda gk: g_data[gk]['count'])
                assigned_in_roi.add(best_gk)
                assigned_per_g[best_gk] = es

                gn = g_data[best_gk]['g_num']
                m_k = len(es_data[es]['prots'] & g_data[best_gk]['prots'])
                tot = g_data[best_gk]['count']
                pct = m_k / tot * 100
                roi_assigned_covs.append(pct)
                row_cells[es] = f"G{gn} ({pct:.1f}%)"
            else:
                row_cells[es] = ""

        # Pass 2: Assign any remaining unassigned Gs in this ROI to their best matching ES in combo
        unassigned = [gk for gk in total_gs if gk not in assigned_in_roi]
        for gk in unassigned:
            matching_es = [es for es in combo if gk in es_data[es]['gks']]
            if matching_es:
                best_es = matching_es[0]
                assigned_in_roi.add(gk)
                assigned_per_g[gk] = best_es
                gn = g_data[gk]['g_num']
                m_k = len(es_data[best_es]['prots'] & g_data[gk]['prots'])
                tot = g_data[gk]['count']
                pct = m_k / tot * 100
                roi_assigned_covs.append(pct)
                if row_cells[best_es]:
                    row_cells[best_es] += f", G{gn} ({pct:.1f}%)"
                else:
                    row_cells[best_es] = f"G{gn} ({pct:.1f}%)"

        # Summary columns for the ROI
        row_cells['Assigned_G_Count'] = len(assigned_in_roi)
        row_cells['ROI_G_Coverage_%'] = round(len(assigned_in_roi) / len(total_gs) * 100, 2)
        row_cells['Avg_Protein_Coverage_in_ROI_%'] = round(float(np.mean(roi_assigned_covs)), 2) if roi_assigned_covs else 0.0

        matrix_rows.append(row_cells)

    return matrix_rows, assigned_per_g


def build_per_g_detail_df(g_data: dict, es_data: dict, assigned_per_g: dict, combo: List[str]) -> pd.DataFrame:
    """
    Constructs the detailed table for EVERY single G-vector in the universe.
    Shows the exact matched proteins and coverage % (e.g. 3/8 = 37.5%).
    """
    rows = []
    combo_set = set(combo)

    # Sort G-keys by patient, roi, g_num
    sorted_gkeys = sorted(
        g_data.keys(),
        key=lambda gk: (g_data[gk]['patient'], g_data[gk]['roi'], g_data[gk]['g_num'])
    )

    for gk in sorted_gkeys:
        d = g_data[gk]
        assigned_es = assigned_per_g.get(gk, None)

        if assigned_es and assigned_es in es_data:
            es_info = es_data[assigned_es]
            matched = sorted(es_info['prots'] & d['prots'])
            cov_pct = round(len(matched) / d['count'] * 100, 2)
            rows.append({
                'Patient': d['patient'],
                'ROI': d['roi'],
                'G_Num': d['g_num'],
                'G_Key': gk,
                'Total_Proteins_in_G': d['count'],
                'G_Proteins': ' | '.join(sorted(d['prots'])),
                'Is_Covered': 'YES',
                'Assigned_ES': assigned_es,
                'ES_Proteins': es_info['protein_set_str'],
                'Matched_Proteins_Count': len(matched),
                'Matched_Proteins': ', '.join(matched),
                'Protein_Coverage_%': cov_pct
            })
        else:
            # Check if covered by any other ES in combo
            matching_in_combo = [es for es in combo if gk in es_data[es]['gks']]
            best_alt = matching_in_combo[0] if matching_in_combo else ''
            m_alt = len(es_data[best_alt]['prots'] & d['prots']) if best_alt else 0
            alt_cov = round(m_alt / d['count'] * 100, 2) if best_alt else 0.0
            matched_alt_list = sorted(es_data[best_alt]['prots'] & d['prots']) if best_alt else []
            rows.append({
                'Patient': d['patient'],
                'ROI': d['roi'],
                'G_Num': d['g_num'],
                'G_Key': gk,
                'Total_Proteins_in_G': d['count'],
                'G_Proteins': ' | '.join(sorted(d['prots'])),
                'Is_Covered': 'YES' if best_alt else 'NO',
                'Assigned_ES': best_alt if best_alt else '',
                'ES_Proteins': es_data[best_alt]['protein_set_str'] if best_alt else '',
                'Matched_Proteins_Count': m_alt,
                'Matched_Proteins': ', '.join(matched_alt_list),
                'Protein_Coverage_%': alt_cov
            })

    return pd.DataFrame(rows)


def build_per_roi_summary_df(tail_g_map: dict, matrix_rows: list, combo: List[str]) -> pd.DataFrame:
    """
    Builds the per-ROI summary table.
    """
    roi_summary_rows = []
    matrix_by_roi = {r['ROI']: r for r in matrix_rows}

    for roi in sorted(tail_g_map.keys()):
        m_row = matrix_by_roi.get(roi, {})
        assigned_es_list = [es for es in combo if m_row.get(es, '') != '']
        roi_summary_rows.append({
            'ROI': roi,
            'Patient': roi.split('_')[0],
            'Total_G_Count': m_row.get('Total_G_in_ROI', len(tail_g_map[roi])),
            'Assigned_G_Count': m_row.get('Assigned_G_Count', 0),
            'ROI_Coverage_%': m_row.get('ROI_G_Coverage_%', 0.0),
            'Avg_Protein_Coverage_%': m_row.get('Avg_Protein_Coverage_in_ROI_%', 0.0),
            'Contributing_ES_Count': len(assigned_es_list),
            'Contributing_ES_List': ', '.join(assigned_es_list)
        })

    return pd.DataFrame(roi_summary_rows)


def style_excel_sheet(ws, title_row_bg="1F497D", header_bg="203764", is_matrix=False):
    """
    Applies professional styling to an openpyxl worksheet.
    """
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color=header_bg, end_color=header_bg, fill_type='solid')
    data_font   = Font(name='Calibri', size=10)
    center_align = Alignment(horizontal='center', vertical='center')
    left_align   = Alignment(horizontal='left', vertical='center')

    ws.row_dimensions[1].height = 28

    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border

    # Format data rows
    for row_idx in range(2, ws.max_row + 1):
        ws.row_dimensions[row_idx].height = 20
        # Alternate shading
        row_fill = PatternFill(start_color="F9FAFB" if row_idx % 2 == 0 else "FFFFFF", fill_type="solid")
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.fill = row_fill
            cell.border = thin_border
            if isinstance(cell.value, (int, float)):
                cell.alignment = center_align
            elif not cell.value:
                cell.alignment = center_align
            else:
                cell.alignment = left_align

    # Column width auto-fit
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col[:100]:  # sample first 100 rows
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12) if not is_matrix else min(max(max_len + 3, 11), 35)


def generate_multiple_combinations(
    es_data: dict,
    g_to_es: dict,
    tail_g_map: dict,
    universe: set,
    top_anchors: int = 40,
    min_marginal_gain: int = 6,
    max_es_count: int = 100
) -> List[dict]:
    """
    Runs multi-anchor greedy searches to generate a diverse pool of combinations.
    """
    print(f"\nGenerating candidate combinations using {top_anchors} diverse anchors (min_gain={min_marginal_gain}, max_es={max_es_count}) ...")
    sorted_es_ids = sorted(
        es_data.keys(),
        key=lambda eid: (es_data[eid]['unique_patients'], es_data[eid]['g_count'], es_data[eid]['protein_count']),
        reverse=True
    )[:top_anchors]

    combos_pool = []
    seen_combos = set()

    for idx, anchor_id in enumerate(sorted_es_ids):
        combo_raw, _ = run_greedy_search(
            es_data, g_to_es, tail_g_map, universe,
            anchor_es_id=anchor_id,
            min_marginal_gain=min_marginal_gain,
            max_es_count=max_es_count
        )
        combo_pruned = backward_pruning(combo_raw, es_data)

        combo_key = tuple(sorted(combo_pruned))
        if combo_key in seen_combos:
            continue
        seen_combos.add(combo_key)

        # Evaluate combination
        all_g = set()
        for eid in combo_pruned:
            all_g.update(es_data[eid]['gks'])

        roi_covs = [len(tail_g_map[r] & all_g) / len(tail_g_map[r]) * 100 for r in tail_g_map]
        rois_covered = sum(1 for p in roi_covs if p > 0)
        avg_roi_cov  = float(np.mean(roi_covs))
        avg_prot_cov = float(np.mean([es_data[eid]['avg_prot_cov'] for eid in combo_pruned]))

        combos_pool.append({
            'anchor': anchor_id,
            'combo_list': combo_pruned,
            'ES_Count': len(combo_pruned),
            'ROIs_Covered': rois_covered,
            'Total_ROIs': len(tail_g_map),
            'Avg_ROI_G_Coverage_%': round(avg_roi_cov, 2),
            'Total_Unique_G_Covered': len(all_g),
            'Total_G_in_Universe': len(universe),
            'Universe_G_Coverage_%': round(len(all_g) / len(universe) * 100, 2),
            'Avg_Protein_Coverage_%': round(avg_prot_cov, 2),
            'Selected_ES_List': ', '.join(combo_pruned)
        })

    # Sort combinations
    combos_pool.sort(
        key=lambda c: (c['ROIs_Covered'], c['Avg_ROI_G_Coverage_%'], c['Total_Unique_G_Covered'], -c['ES_Count']),
        reverse=True
    )

    for idx, c in enumerate(combos_pool):
        c['Rank'] = idx + 1
        c['ES_COM'] = f"ES_COM_{idx + 1}"

    return combos_pool


def main():
    t0 = time.time()
    print("=" * 75)
    print("OPTIMAL ES COMBINATION FINDER (Length >= 3, Balanced & Conflict-Free)")
    print("=" * 75)

    MIN_PROTEIN_COUNT = 3
    MIN_MARGINAL_GAIN = 1
    MAX_ES_COUNT = 100
    TOP_ANCHORS = 40

    g_data, tail_g_map, universe, es_data, df_es_all, g_to_es = load_all_data(min_protein_count=MIN_PROTEIN_COUNT)
    print(f"  Total ROIs          : {len(tail_g_map)}")
    print(f"  Total G-vectors     : {len(universe):,}")
    print(f"  Active ES (size>=3) : {len(es_data):,}")
    print(f"  Min Marginal Gain   : {MIN_MARGINAL_GAIN} Gs/ES")
    print(f"  Max ES Target       : {MAX_ES_COUNT}")

    # Generate top combinations
    combos_pool = generate_multiple_combinations(
        es_data, g_to_es, tail_g_map, universe,
        top_anchors=TOP_ANCHORS,
        min_marginal_gain=MIN_MARGINAL_GAIN,
        max_es_count=MAX_ES_COUNT
    )
    print(f"  Generated {len(combos_pool)} distinct pruned combinations")

    best_combo_info = combos_pool[0]
    best_combo_name = best_combo_info['ES_COM']
    best_combo_list = best_combo_info['combo_list']

    print(f"\n[WINNING COMBINATION: {best_combo_name}]")
    print(f"  ES Sequences Count : {best_combo_info['ES_Count']}")
    print(f"  ROIs Covered       : {best_combo_info['ROIs_Covered']} / {best_combo_info['Total_ROIs']}")
    print(f"  Avg ROI G Coverage : {best_combo_info['Avg_ROI_G_Coverage_%']}%")
    print(f"  Universe Coverage  : {best_combo_info['Total_Unique_G_Covered']} / {best_combo_info['Total_G_in_Universe']} ({best_combo_info['Universe_G_Coverage_%']}%)")
    print(f"  Avg Prot Coverage  : {best_combo_info['Avg_Protein_Coverage_%']}%")

    # Run full coverage curve for best anchor
    _, curve_snapshots = run_greedy_search(
        es_data, g_to_es, tail_g_map, universe,
        anchor_es_id=best_combo_info['anchor'],
        min_marginal_gain=MIN_MARGINAL_GAIN,
        max_es_count=MAX_ES_COUNT
    )
    df_curve = pd.DataFrame(curve_snapshots)

    # 1. Combinations Summary DataFrame
    summary_cols = [
        'Rank', 'ES_COM', 'ES_Count', 'ROIs_Covered', 'Total_ROIs',
        'Avg_ROI_G_Coverage_%', 'Total_Unique_G_Covered', 'Total_G_in_Universe',
        'Universe_G_Coverage_%', 'Avg_Protein_Coverage_%', 'Selected_ES_List'
    ]
    df_summary = pd.DataFrame(combos_pool)[summary_cols]

    # 2. Optimal Combination Selected ES Detail
    optimal_es_rows = []
    for rank_idx, eid in enumerate(best_combo_list):
        d = es_data[eid]
        optimal_es_rows.append({
            'Rank': rank_idx + 1,
            'ES_ID': eid,
            'Protein_Count': d['protein_count'],
            'Protein_Set': d['protein_set_str'],
            'Unique_Patients': d['unique_patients'],
            'ROIs_Covered': d['rois_covered_count'],
            'Total_G_Values_Covered': d['g_count'],
            'Avg_Protein_Coverage_Per_G_%': d['avg_prot_cov'],
            'Tails_Detail': d['tails_detail']
        })
    df_optimal_es = pd.DataFrame(optimal_es_rows)

    # 3. Conflict-Free Assignment Matrix (ROI x ES)
    print("\nAssigning conflict-free G-values per ROI ...")
    matrix_rows, assigned_per_g = assign_conflict_free_for_combo(best_combo_list, tail_g_map, g_data, es_data)
    df_matrix = pd.DataFrame(matrix_rows)
    meta_cols = ['ROI', 'Patient', 'Total_G_in_ROI', 'Assigned_G_Count', 'ROI_G_Coverage_%', 'Avg_Protein_Coverage_in_ROI_%']
    df_matrix = df_matrix[meta_cols + best_combo_list]

    # 4. Detailed Per-G Coverage Table
    print("Building detailed per-G coverage table ...")
    df_per_g = build_per_g_detail_df(g_data, es_data, assigned_per_g, best_combo_list)

    # 5. Per-ROI Coverage Summary
    df_per_roi = build_per_roi_summary_df(tail_g_map, matrix_rows, best_combo_list)

    # Export to Excel
    print("\nExporting comprehensive 6-sheet Excel report ...")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_ts = os.path.join(OUTPUT_DIR, f'Optimal_ES_Result_{ts}.xlsx')
    out_main = os.path.join(OUTPUT_DIR, 'Optimal_ES_Result.xlsx')

    for out_path in [out_ts, out_main]:
        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            df_summary.to_excel(writer, sheet_name='Combinations_Summary', index=False)
            df_optimal_es.to_excel(writer, sheet_name='Optimal_Combination', index=False)
            df_matrix.to_excel(writer, sheet_name='Best_Combination_Details', index=False)
            df_per_g.to_excel(writer, sheet_name='Per_G_Coverage_Detail', index=False)
            df_per_roi.to_excel(writer, sheet_name='Per_ROI_Coverage', index=False)
            df_curve.to_excel(writer, sheet_name='Coverage_Curve', index=False)

        # Style with openpyxl
        wb = openpyxl.load_workbook(out_path)
        for sname in wb.sheetnames:
            style_excel_sheet(wb[sname], is_matrix=(sname == 'Best_Combination_Details'))
        wb.save(out_path)
        print(f"  Saved & styled: {out_path}")

    elapsed = time.time() - t0
    print(f"\nExecution finished in {elapsed:.2f} seconds!")
    print("=" * 75)


if __name__ == '__main__':
    main()
