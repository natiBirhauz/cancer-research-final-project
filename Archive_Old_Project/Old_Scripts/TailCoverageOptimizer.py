"""
TailCoverageOptimizer.py
------------------------
For each ROI (from All_Tail_Proteins_Tables), finds the set of sequences from
Exact_Sequences_Miner that covers the maximum number of tail proteins (G-components),
with the constraint that the same G-component cannot appear in more than one
selected sequence per ROI.

Logic:
- Each ROI has multiple G-components (G1, G2, G3...)
- Each G-component has tail proteins (e.g. G1+ has CD20, CD3; G1- has Ki67...)
- A sequence from Exact_Sequences_Miner "covers" a G-component if that G-component
  appears in the sequence's Participating ROIs column
- The same G-component (e.g. P1_ROI_03_G5) cannot be used by more than one selected sequence
"""

import os
import glob
import re
import pandas as pd
from datetime import datetime
from collections import defaultdict


# ─────────────────────────────────────────────
# 1. LOAD TAIL PROTEINS PER ROI
# ─────────────────────────────────────────────

def load_tail_proteins(tails_folder):
    """
    Returns a dict:
      roi_id  →  { 'G1': {'proteins': set(), 'sign': '+'/'- '}, ... }
      e.g. 'P1_ROI_03' → { 'G1': {'proteins': {'CD20','LAG3'}, 'dominant_sign': '+'}, ... }
    """
    files = glob.glob(os.path.join(tails_folder, '*.xlsx')) + \
            glob.glob(os.path.join(tails_folder, '*.csv'))

    roi_data = {}  # roi_id → { g_key → set of proteins }

    for f in sorted(files):
        roi_id = os.path.basename(f) \
                     .replace('_tail_proteins.xlsx', '') \
                     .replace('_tail_proteins.csv', '')

        try:
            df = pd.read_excel(f) if f.endswith('.xlsx') else pd.read_csv(f)
        except Exception as e:
            print(f"  Warning: could not read {f}: {e}")
            continue

        # Collect all G numbers present in this file
        g_cols = [c for c in df.columns if re.match(r'G\d+[+-]$', c)]
        g_numbers = sorted(set(int(re.search(r'G(\d+)', c).group(1)) for c in g_cols))

        g_dict = {}
        for g_num in g_numbers:
            proteins = set()
            col_plus  = f'G{g_num}+'
            col_minus = f'G{g_num}-'
            if col_plus in df.columns:
                proteins.update(df[col_plus].dropna().astype(str).str.strip().tolist())
            if col_minus in df.columns:
                proteins.update(df[col_minus].dropna().astype(str).str.strip().tolist())
            proteins.discard('nan')
            if proteins:
                g_dict[f'G{g_num}'] = proteins

        if g_dict:
            roi_data[roi_id] = g_dict

    print(f"Loaded {len(roi_data)} ROIs from tail protein tables.")
    return roi_data


# ─────────────────────────────────────────────
# 2. PARSE PARTICIPATING ROIs FROM SEQUENCES FILE
# ─────────────────────────────────────────────

def parse_roi_gs(roi_string):
    """
    'P1_ROI_03_G4 (+) | P2_ROI_5_G3 (-) | ...'
    Returns set of full ROI_G keys like {'P1_ROI_03_G4', 'P2_ROI_5_G3', ...}
    """
    if pd.isna(roi_string):
        return set()
    result = set()
    for item in str(roi_string).split(' | '):
        item = item.strip()
        # Remove trailing ' (+)' or ' (-)'
        roi_g = re.sub(r'\s+\([+-]\)\s*$', '', item).strip()
        if roi_g:
            result.add(roi_g)
    return result


# ─────────────────────────────────────────────
# 3. GREEDY COVER PER ROI
# ─────────────────────────────────────────────

def find_best_coverage_for_roi(roi_id, g_dict, seq_df, seq_roi_gs):
    """
    For one ROI, greedily pick sequences that cover its G-components
    (no G-component can appear in two selected sequences).

    Returns list of selected sequence dicts with coverage details.
    """
    # Build map: G-key (e.g. 'P1_ROI_03_G4') → G number (e.g. 'G4')
    roi_g_keys = {}  # full_key → G_label
    for g_label in g_dict:          # g_label = 'G1', 'G2', ...
        full_key = f"{roi_id}_{g_label}"
        roi_g_keys[full_key] = g_label

    total_g = len(roi_g_keys)
    if total_g == 0:
        return [], 0, 0

    covered_g_keys = set()        # G-keys already assigned to a sequence
    selected = []

    # Candidates: sequences that contain at least one G-key of this ROI
    candidate_mask = seq_df['_roi_g_set'].apply(
        lambda s: bool(s & set(roi_g_keys.keys()))
    )
    candidates = seq_df[candidate_mask].copy()

    if candidates.empty:
        return [], 0, total_g

    # Greedy: at each step pick the sequence covering the most NEW uncovered G-keys
    remaining = list(candidates.index)

    while remaining:
        best_idx   = None
        best_new   = 0
        best_g_set = set()

        for idx in remaining:
            roi_g_in_seq = seq_df.loc[idx, '_roi_g_set'] & set(roi_g_keys.keys())
            new_g = roi_g_in_seq - covered_g_keys
            if len(new_g) > best_new:
                best_new   = len(new_g)
                best_idx   = idx
                best_g_set = new_g

        if best_new == 0:
            break   # nothing new can be covered

        # Record what proteins these G-components contained
        covered_proteins = set()
        covered_g_labels = set()
        for gk in best_g_set:
            g_label = roi_g_keys[gk]
            covered_proteins |= g_dict.get(g_label, set())
            covered_g_labels.add(g_label)

        row = seq_df.loc[best_idx]
        selected.append({
            'ROI':                 roi_id,
            'Sequence_Rank':       len(selected) + 1,
            'Sequence Length':     row['Sequence Length'],
            'Unique Patients Count': row['Unique Patients Count'],
            'Exact Sign Sequence': row['Exact Sign Sequence'],
            'G_Components_Covered': ', '.join(sorted(covered_g_labels)),
            'Proteins_Covered':    ', '.join(sorted(covered_proteins)),
            'New_G_Count':         best_new,
            'Participating ROIs':  row['Participating ROIs'],
        })

        covered_g_keys |= best_g_set
        remaining.remove(best_idx)

        if covered_g_keys >= set(roi_g_keys.keys()):
            break   # full coverage achieved

    covered_count = len(covered_g_keys)
    return selected, covered_count, total_g


# ─────────────────────────────────────────────
# 4. MAIN
# ─────────────────────────────────────────────

def main():
    tails_folder   = 'All_Tail_Proteins_Tables'
    sequences_file = 'Docs/Exact_Sequences_Miner_20260630_090923.csv'

    print("=" * 65)
    print("TAIL COVERAGE OPTIMIZER")
    print("=" * 65)

    # ── Load sequences ──
    print(f"\nLoading sequences from: {sequences_file}")
    seq_df = pd.read_csv(sequences_file)
    print(f"  {len(seq_df):,} sequences loaded.")

    # Pre-parse ROI_G sets (once, for speed)
    print("  Pre-parsing Participating ROIs...")
    seq_df['_roi_g_set'] = seq_df['Participating ROIs'].apply(parse_roi_gs)

    # Build inverted index: roi_g_key → list of seq indices
    print("  Building inverted index (roi_G_key → sequence indices)...")
    inverted_index = defaultdict(list)
    for idx, roi_g_set in seq_df['_roi_g_set'].items():
        for key in roi_g_set:
            inverted_index[key].append(idx)
    print(f"  Index built: {len(inverted_index):,} unique ROI_G keys.")

    # ── Load tails ──
    print(f"\nLoading tail protein tables from: {tails_folder}")
    roi_data = load_tail_proteins(tails_folder)

    # ── Process each ROI ──
    all_results   = []
    summary_rows  = []

    total_rois = len(roi_data)
    for i, (roi_id, g_dict) in enumerate(sorted(roi_data.items()), 1):
        print(f"  [{i}/{total_rois}] {roi_id}  ({len(g_dict)} G-components)...")

        # Get only candidate sequence indices via the inverted index (fast!)
        all_full_keys = {f"{roi_id}_{g}" for g in g_dict}
        candidate_idx = set()
        for fk in all_full_keys:
            candidate_idx.update(inverted_index.get(fk, []))

        # Subset seq_df to candidates only
        if not candidate_idx:
            all_results.append({
                'ROI': roi_id, 'Sequence_Rank': 0,
                'Sequence Length': 0, 'Unique Patients Count': 0,
                'Exact Sign Sequence': 'NO MATCH',
                'G_Components_Covered': '', 'Proteins_Covered': '',
                'New_G_Count': 0, 'Participating ROIs': ''
            })
            summary_rows.append({
                'ROI': roi_id, 'Total_G_Components': len(g_dict),
                'Covered_G_Components': 0,
                'Uncovered_G_Components': len(g_dict),
                'Coverage_%': 0.0, 'Sequences_Used': 0
            })
            print(f"         → 0/{len(g_dict)} G-components covered (no matching sequences)")
            continue

        sub_df = seq_df.loc[sorted(candidate_idx)]

        selected, covered, total_g = find_best_coverage_for_roi(
            roi_id, g_dict, sub_df, None
        )

        pct = covered / total_g * 100 if total_g > 0 else 0
        print(f"         → {covered}/{total_g} G-components covered ({pct:.1f}%) "
              f"using {len(selected)} sequence(s)")

        all_results.extend(selected)
        summary_rows.append({
            'ROI':                    roi_id,
            'Total_G_Components':     total_g,
            'Covered_G_Components':   covered,
            'Uncovered_G_Components': total_g - covered,
            'Coverage_%':             round(pct, 2),
            'Sequences_Used':         len(selected),
        })

    # ── Export ──
    print("\nExporting results...")
    if not os.path.exists('Docs'):
        os.makedirs('Docs')

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    # Detailed results
    detail_df = pd.DataFrame(all_results)
    detail_csv   = f'Docs/Tail_Coverage_Detail_{ts}.csv'
    detail_xlsx  = f'Docs/Tail_Coverage_Detail_{ts}.xlsx'
    detail_df.to_csv(detail_csv, index=False)

    # Summary
    summary_df = pd.DataFrame(summary_rows)
    summary_df = summary_df.sort_values('Coverage_%', ascending=False)
    summary_csv  = f'Docs/Tail_Coverage_Summary_{ts}.csv'
    summary_xlsx = f'Docs/Tail_Coverage_Summary_{ts}.xlsx'
    summary_df.to_csv(summary_csv, index=False)

    # Excel with two sheets
    with pd.ExcelWriter(detail_xlsx, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        detail_df.to_excel(writer, sheet_name='Detail', index=False)

    print(f"\n  Detail  CSV  → {detail_csv}")
    print(f"  Summary CSV  → {summary_csv}")
    print(f"  Excel        → {detail_xlsx}  (Summary + Detail sheets)")

    # ── Overall stats ──
    total_g_all   = summary_df['Total_G_Components'].sum()
    covered_g_all = summary_df['Covered_G_Components'].sum()
    overall_pct   = covered_g_all / total_g_all * 100 if total_g_all > 0 else 0

    fully_covered = (summary_df['Coverage_%'] == 100).sum()

    print("\n" + "=" * 65)
    print("OVERALL SUMMARY")
    print("=" * 65)
    print(f"  ROIs processed       : {len(summary_df)}")
    print(f"  ROIs fully covered   : {fully_covered} / {len(summary_df)}")
    print(f"  Total G-components   : {total_g_all}")
    print(f"  Covered G-components : {covered_g_all}  ({overall_pct:.2f}%)")
    print(f"  Uncovered            : {total_g_all - covered_g_all}")
    print("=" * 65)


if __name__ == "__main__":
    main()
