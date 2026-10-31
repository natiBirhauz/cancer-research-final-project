"""
Optimal_ES_COM_Conflict_Free_Report.py
--------------------------------------
1. Calculates Average Section G-Coverage % across all 120 sections.
2. Optimizes Set Cover targeting maximum G-vector coverage across all 120 sections.
3. Applies Backward Pruning pass to eliminate redundant ES sequences.
4. Solves G-assignment conflicts per ROI:
   - Exactly ONE G-vector per cell.
   - NO duplicated G-vector across ES columns for the same ROI/patient.
5. Generates Excel output with two sheets:
   - '100_Combinations_Summary'
   - 'ES_COM_1_Details'
"""

import pandas as pd
import numpy as np
import re
import os
import glob
import time
from collections import defaultdict

def find_source_report():
    patterns = ['Docs/Exact_Sequences_Miner_*.csv', 'Docs/CoreESAnalysis_PartA_*.csv', '*.csv']
    files = []
    for p in patterns:
        for f in glob.glob(p):
            if "Combinations" not in f and "Analysis" not in f and "Report" not in f:
                files.append(f)
    return max(files, key=os.path.getctime) if files else None

def parse_g_level_coverage(roi_string):
    """
    Parses 'P12_ROI_1_G1 (+)' or 'P2_ROI_5(1G/23G 4%)' into structured tuples.
    """
    parsed_items = []
    if pd.isna(roi_string): 
        return parsed_items
    
    normalized_str = str(roi_string).replace("|", ";").replace(",", ";")
    parts = normalized_str.split(";")
    
    for part in parts:
        part = part.strip()
        if not part: continue
        
        m1 = re.search(r"(.+?)_G(\d+)\s*\(([-+])\)", part)
        m2 = re.search(r"([A-Za-z0-9_]+)\s*\(\s*(\d+)G", part)
        
        if m1:
            section_name, g_num, sign = m1.groups()
            section_name = section_name.strip()
            full_g_id = f"{section_name}_G{g_num}"
            g_val = f"G{g_num}({sign})"
            parsed_items.append((section_name, full_g_id, g_val))
        elif m2:
            section_name, g_num = m2.groups()
            section_name = section_name.strip()
            full_g_id = f"{section_name}_G{g_num}"
            g_val = f"G{g_num}"
            parsed_items.append((section_name, full_g_id, g_val))
            
    return parsed_items

def assign_conflict_free_gs(sec_name, es_list, es_sec_g_map):
    """
    Assigns at most ONE unique G-vector to each ES column for a given ROI/section.
    Ensures no duplicate G-vectors across ES columns for the same ROI.
    """
    assigned_gs = set()
    cell_assignments = {}
    
    es_candidates = []
    for es in es_list:
        candidates = es_sec_g_map.get((sec_name, es), [])
        if candidates:
            es_candidates.append((es, candidates))
            
    # Sort ES sequences with fewer candidate G-vectors first to give them priority
    es_candidates.sort(key=lambda x: len(x[1]))
    
    for es, candidates in es_candidates:
        chosen_g = None
        for g_val in candidates:
            if g_val not in assigned_gs:
                chosen_g = g_val
                break
        if chosen_g:
            assigned_gs.add(chosen_g)
            cell_assignments[es] = chosen_g
            
    return cell_assignments

def run_optimal_g_coverage_pipeline():
    report_path = find_source_report()
    if not report_path:
        print("Error: Could not find sequence dataset CSV in Docs/ folder.")
        return

    start_time = time.time()
    print(f"Loading dataset from: {report_path}...")

    try:
        df = pd.read_csv(report_path)
        if len(df.columns) == 1 and '\t' in df.columns[0]:
            df = pd.read_csv(report_path, sep='\t')
    except Exception:
        df = pd.read_excel(report_path)

    target_col = 'Participating ROIs' if 'Participating ROIs' in df.columns else 'Tails_Detail'
    
    print("Mapping ROIs, G-vectors, and ES candidate sets...")
    
    es_labels = []
    section_total_g = defaultdict(set)
    es_sec_covered_g = defaultdict(set)
    es_sec_g_map = defaultdict(list)
    all_sections = set()
    all_g_vectors = set()
    es_all_g_map = defaultdict(set)

    for idx, row in df.iterrows():
        label = row.get('ES_Label', None)
        if pd.isna(label) or str(label).strip() in ['', 'ES_0']:
            label = f"ES{idx + 1}"
        else:
            label = str(label).strip()

        roi_str = row[target_col]
        parsed_list = parse_g_level_coverage(roi_str)

        if parsed_list:
            es_labels.append(label)
            for sec_name, full_g_id, g_val in parsed_list:
                all_sections.add(sec_name)
                all_g_vectors.add(full_g_id)
                section_total_g[sec_name].add(full_g_id)
                es_sec_covered_g[(sec_name, label)].add(full_g_id)
                es_all_g_map[label].add(full_g_id)
                if g_val not in es_sec_g_map[(sec_name, label)]:
                    es_sec_g_map[(sec_name, label)].append(g_val)

    num_es = len(es_labels)
    total_sections_count = len(all_sections)
    total_g_count = len(all_g_vectors)
    print(f"Total unique sections: {total_sections_count}")
    print(f"Total unique G-vectors: {total_g_count}")
    print(f"Active ES sequences mapped: {num_es:,}")

    combinations = []
    num_combos = min(100, num_es)

    print("\nOptimizing ES combinations targeting MAX G-Vector Coverage...")

    for i in range(num_combos):
        anchor_es = es_labels[i]
        combo = [anchor_es]
        covered_g = set(es_all_g_map[anchor_es])
        uncovered_g = all_g_vectors - covered_g

        # G-Vector Level Greedy Loop (continues until max G-vector coverage is reached)
        while uncovered_g:
            best_es = None
            best_gain = 0
            for es in es_labels:
                if es in combo: continue
                gain = len(es_all_g_map[es] & uncovered_g)
                if gain > best_gain:
                    best_gain = gain
                    best_es = es

            if best_gain == 0:
                break

            combo.append(best_es)
            covered_g.update(es_all_g_map[best_es])
            uncovered_g -= es_all_g_map[best_es]

        # Backward Pruning: Eliminate redundant ES sequences with 0 unique contribution
        pruned_combo = list(combo)
        changed = True
        while changed:
            changed = False
            for es in list(pruned_combo):
                others = [x for x in pruned_combo if x != es]
                others_g = set()
                for x in others:
                    others_g.update(es_all_g_map[x])
                unique_g = es_all_g_map[es] - others_g
                if len(unique_g) == 0:
                    pruned_combo.remove(es)
                    changed = True
                    break

        # Calculate Section G-Coverage % for each section
        sec_coverages = []
        covered_sections_set = set()
        for sec in all_sections:
            sec_g_tot = len(section_total_g[sec])
            sec_g_cov = set()
            for es in pruned_combo:
                sec_g_cov.update(es_sec_covered_g[(sec, es)])
            if len(sec_g_cov) > 0:
                covered_sections_set.add(sec)
            pct = (len(sec_g_cov) / sec_g_tot * 100) if sec_g_tot > 0 else 0.0
            sec_coverages.append(pct)

        avg_section_g_coverage = float(np.mean(sec_coverages))

        combinations.append({
            'ES_com': f"ES_COM_{i+1}",
            'ES_Count': len(pruned_combo),
            'ROIs_Covered': len(covered_sections_set),
            'ROIs_Total': total_sections_count,
            'Avg_Section_G_Coverage_Pct': round(avg_section_g_coverage, 2),
            'combo_list': pruned_combo
        })

    # Sort combinations by (1) ROIs_Covered, (2) Max Avg_Section_G_Coverage_Pct, (3) Min ES_Count
    df_combos = pd.DataFrame(combinations).sort_values(
        by=['ROIs_Covered', 'Avg_Section_G_Coverage_Pct', 'ES_Count'],
        ascending=[False, False, True]
    ).reset_index(drop=True)

    for idx in range(len(df_combos)):
        df_combos.at[idx, 'ES_com'] = f"ES_COM_{idx+1}"

    table1 = df_combos.drop(columns=['combo_list']).copy()

    # Build Table 2 (ES_COM_1_Details) with Conflict-Free Single G Assignments
    best_combo_list = df_combos.iloc[0]['combo_list']
    best_combo_name = df_combos.iloc[0]['ES_com']

    print(f"\nWINNER: {best_combo_name} covers {df_combos.iloc[0]['ROIs_Covered']}/{total_sections_count} sections!")
    print(f"Average Section G-Coverage: {df_combos.iloc[0]['Avg_Section_G_Coverage_Pct']}% using {len(best_combo_list)} ES sequences.")
    print("Constructing conflict-free matrix table for ES_COM_1...")

    matrix_data = defaultdict(dict)

    for sec in sorted(list(all_sections)):
        assignments = assign_conflict_free_gs(sec, best_combo_list, es_sec_g_map)
        for es in best_combo_list:
            matrix_data[sec][es] = assignments.get(es, "")

    df_matrix = pd.DataFrame.from_dict(matrix_data, orient='index')
    df_matrix = df_matrix.reindex(columns=best_combo_list).fillna("")
    df_matrix.index.name = 'ROI'
    df_matrix.reset_index(inplace=True)

    # Save Output Excel
    os.makedirs('Docs', exist_ok=True)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    output_file = f'Docs/Optimal_ES_COM_Report_{timestamp}.xlsx'

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        table1.to_excel(writer, sheet_name='100_Combinations_Summary', index=False)
        df_matrix.to_excel(writer, sheet_name=f'{best_combo_name}_Details', index=False)

    elapsed = time.time() - start_time
    print(f"\nExecution completed in {elapsed:.2f} seconds!")
    print(f"Report successfully saved to: {output_file}")

if __name__ == "__main__":
    run_optimal_g_coverage_pipeline()