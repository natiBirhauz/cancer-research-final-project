"""
generate_es_combinations.py
----------------------------
Generates the comprehensive ES combinations table:
- ES = unique set of protein names (no signs, no order).
- Mines all valid protein subsets (size >= 2) that appear across all G-vectors.
- Maps each ES to every G-vector containing it.
- Computes per-G protein coverage: matched / total_proteins_in_G * 100% (e.g. 3/8 = 37.5%).
- Computes Tails_Detail per ROI and per G.
- Computes Avg_Protein_Coverage_Per_G_%, Unique_Patients_Count, ROIs_Covered_Count, G_Values_Covered.
- Exports to Docs/ES_Combinations.csv and Docs/ES_Combinations_<timestamp>.csv.
"""

import os
import re
import time
import glob
from datetime import datetime
from itertools import combinations
from collections import defaultdict
import pandas as pd

TAILS_CSV  = os.path.join('Docs', 'All_Tails.csv')
OUTPUT_DIR = 'Docs'


def load_g_vectors(tails_csv_path: str):
    """
    Loads G-vectors from All_Tails.csv.
    Returns:
      g_data: {gkey: {'roi': str, 'patient': str, 'g_num': int, 'prots': frozenset, 'count': int}}
      total_g_per_roi: {roi: int}
    """
    if not os.path.exists(tails_csv_path):
        raise FileNotFoundError(f"Could not find {tails_csv_path}. Run export_tails_table.py first.")

    df = pd.read_csv(tails_csv_path)
    g_data = {}
    g_by_roi = defaultdict(list)

    for _, row in df.iterrows():
        gk = row['G_Key']
        roi = row['ROI_Key']
        p = row['Patient']
        prots = frozenset(x.strip() for x in str(row['Protein_Set']).split(' | ') if x.strip())
        g_data[gk] = {
            'roi': roi,
            'patient': p,
            'g_num': int(row['G_Num']),
            'prots': prots,
            'count': len(prots),
            'protein_set_str': str(row['Protein_Set'])
        }
        g_by_roi[roi].append(gk)

    total_g_per_roi = {roi: len(gks) for roi, gks in g_by_roi.items()}
    return g_data, total_g_per_roi


def mine_es_combinations(g_data: dict, min_size: int = 3):
    """
    Finds all unique protein combinations appearing across G-vectors,
    and maps each ES to every G-vector it covers:
    - Covers G if at least 2 proteins match (for 2-protein Gs, covers 100%!).
    """
    pair_to_gks = defaultdict(set)
    triplets = set()
    for gk, d in g_data.items():
        plist = sorted(d['prots'])
        for p in combinations(plist, 2):
            pair_to_gks[frozenset(p)].add(gk)
        for t in combinations(plist, min_size):
            triplets.add(frozenset(t))

    subset_to_gks = {}
    for t in triplets:
        plist = sorted(t)
        p1 = frozenset([plist[0], plist[1]])
        p2 = frozenset([plist[0], plist[2]])
        p3 = frozenset([plist[1], plist[2]])
        matching_gks = pair_to_gks[p1] | pair_to_gks[p2] | pair_to_gks[p3]
        subset_to_gks[t] = sorted(matching_gks)

    return subset_to_gks


def build_es_dataframe(subset_to_gks: dict, g_data: dict, total_g_per_roi: dict) -> pd.DataFrame:
    """
    Computes all statistics and Tails_Detail for each ES.
    Sorts by dominance: Unique_Patients_Count DESC, ROIs_Covered_Count DESC, G_Values_Covered DESC.
    """
    raw_records = []
    for sub, g_list in subset_to_gks.items():
        k = len(sub)
        patients = set(g_data[gk]['patient'] for gk in g_list)
        rois = set(g_data[gk]['roi'] for gk in g_list)
        covs = [len(sub & g_data[gk]['prots']) / g_data[gk]['count'] * 100 for gk in g_list]
        avg_cov = sum(covs) / len(covs)

        raw_records.append({
            '_sub': sub,
            '_g_list': g_list,
            'Protein_Count': k,
            'Unique_Patients_Count': len(patients),
            'ROIs_Covered_Count': len(rois),
            'G_Values_Covered': len(g_list),
            'Avg_Protein_Coverage_Per_G_%': round(avg_cov, 2),
            'Protein_Set': ' | '.join(sorted(sub)),
        })

    # Sort dominant patterns first
    df_raw = pd.DataFrame(raw_records)
    df_raw = df_raw.sort_values(
        by=['Unique_Patients_Count', 'ROIs_Covered_Count', 'G_Values_Covered', 'Avg_Protein_Coverage_Per_G_%'],
        ascending=[False, False, False, False]
    ).reset_index(drop=True)

    # Assign ES_IDs and build Tails_Detail
    final_rows = []
    for idx, row in df_raw.iterrows():
        es_id = f"ES{idx + 1}"
        sub = row['_sub']
        g_list = row['_g_list']
        k = row['Protein_Count']

        # Group by ROI
        roi_to_gs = defaultdict(list)
        for gk in g_list:
            matched_k = len(sub & g_data[gk]['prots'])
            roi_to_gs[g_data[gk]['roi']].append((g_data[gk]['g_num'], matched_k, g_data[gk]['count']))

        td_parts = []
        for r_id, g_items in sorted(roi_to_gs.items(), key=lambda x: -len(x[1])):
            g_strs = [f"G{gn}:{m}/{tot}={m/tot*100:.1f}%" for gn, m, tot in sorted(g_items)]
            g_joined = ', '.join(g_strs)
            td_parts.append(f"{r_id}({len(g_items)}/{total_g_per_roi[r_id]}G | {g_joined})")

        final_rows.append({
            'ES_ID': es_id,
            'Protein_Count': k,
            'Unique_Patients_Count': row['Unique_Patients_Count'],
            'ROIs_Covered_Count': row['ROIs_Covered_Count'],
            'G_Values_Covered': row['G_Values_Covered'],
            'Avg_Protein_Coverage_Per_G_%': row['Avg_Protein_Coverage_Per_G_%'],
            'Protein_Set': row['Protein_Set'],
            'Tails_Detail': '; '.join(td_parts),
            'Participating_G_Values': ' | '.join(g_list)
        })

    return pd.DataFrame(final_rows)


def main():
    t0 = time.time()
    print("=" * 70)
    print("GENERATE ES COMBINATIONS (Unique Protein Subsets, Size >= 3, No Signs)")
    print("=" * 70)

    print(f"\n[1/3] Loading G-vectors from: {TAILS_CSV}")
    g_data, total_g_per_roi = load_g_vectors(TAILS_CSV)
    print(f"      Loaded {len(g_data):,} G-vectors across {len(total_g_per_roi)} ROIs")

    print(f"\n[2/3] Mining all unique ES combinations (size >= 3) ...")
    subset_to_gks = mine_es_combinations(g_data, min_size=3)
    print(f"      Mined {len(subset_to_gks):,} unique ES combinations in {time.time() - t0:.2f}s")

    print(f"      Building detailed statistics table ...")
    df_es = build_es_dataframe(subset_to_gks, g_data, total_g_per_roi)
    print(f"      Constructed table with {len(df_es):,} rows")

    print(f"\n[3/3] Exporting ...")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_ts = os.path.join(OUTPUT_DIR, f'ES_Combinations_{ts}.csv')
    out_main = os.path.join(OUTPUT_DIR, 'ES_Combinations.csv')

    df_es.to_csv(out_ts, index=False, encoding='utf-8-sig')
    df_es.to_csv(out_main, index=False, encoding='utf-8-sig')
    print(f"      Saved: {out_main}")
    print(f"      Saved: {out_ts}")

    # Summary
    print(f"\n{'=' * 70}")
    print("SUMMARY")
    print(f"{'=' * 70}")
    print(f"  Total Unique ES     : {len(df_es):,}")
    print(f"  Protein Count Range : {df_es['Protein_Count'].min()} - {df_es['Protein_Count'].max()}")
    print(f"  Patients Count Range: {df_es['Unique_Patients_Count'].min()} - {df_es['Unique_Patients_Count'].max()}")
    print(f"  G_Values Count Range: {df_es['G_Values_Covered'].min()} - {df_es['G_Values_Covered'].max()}")
    print(f"\n  Top 5 ES Combinations:")
    cols_preview = ['ES_ID', 'Protein_Count', 'Unique_Patients_Count', 'G_Values_Covered', 'Avg_Protein_Coverage_Per_G_%', 'Protein_Set']
    print(df_es.head(5)[cols_preview].to_string(index=False))

    # Highlight specific example if present
    example_target = "CD163 | CD21 | CXCR3"
    match = df_es[df_es['Protein_Set'] == example_target]
    if not match.empty:
        ex_row = match.iloc[0]
        print(f"\n  [Example Check: {example_target}]")
        print(f"    ES_ID                : {ex_row['ES_ID']}")
        print(f"    Unique Patients      : {ex_row['Unique_Patients_Count']}")
        print(f"    G_Values Covered     : {ex_row['G_Values_Covered']}")
        print(f"    Avg Protein Coverage : {ex_row['Avg_Protein_Coverage_Per_G_%']}%")
        print(f"    Tails_Detail sample  : {ex_row['Tails_Detail'][:120]}...")

    print(f"\nExecution time: {time.time() - t0:.2f}s")
    print(f"{'=' * 70}")


if __name__ == '__main__':
    main()
