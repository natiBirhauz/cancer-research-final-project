"""
MinimalSetCover.py
------------------
Finds the MINIMAL number of sequence motifs that:
1. Cover maximum ROIs (G-components) across all patients
2. Each sequence appears in as many ROIs as possible
3. Together define maximum biological processes

This is a Set Cover optimization problem.
"""

import os
import pandas as pd
import numpy as np
from collections import defaultdict
from datetime import datetime

def load_results_file():
    """Load the most recent Exact_Sequences_Miner output."""
    docs_path = 'Docs'
    if not os.path.exists(docs_path):
        print("Error: 'Docs' folder not found. Run GlobalClusterReverseSign.py first.")
        return None
    
    # Find all result files
    csv_files = [f for f in os.listdir(docs_path) if f.startswith('Exact_Sequences_Miner') and f.endswith('.csv')]
    xlsx_files = [f for f in os.listdir(docs_path) if f.startswith('Exact_Sequences_Miner') and f.endswith('.xlsx')]
    
    # Prefer CSV (has all data)
    if csv_files:
        latest_file = sorted(csv_files)[-1]
        file_path = os.path.join(docs_path, latest_file)
        print(f"Loading: {file_path}")
        return pd.read_csv(file_path)
    elif xlsx_files:
        latest_file = sorted(xlsx_files)[-1]
        file_path = os.path.join(docs_path, latest_file)
        print(f"Loading: {file_path}")
        return pd.read_excel(file_path)
    else:
        print("Error: No Exact_Sequences_Miner results found in Docs folder.")
        return None

def parse_participating_rois(roi_string):
    """
    Parse the 'Participating ROIs' column.
    Example: "P1_G1 (+) | P2_G3 (-) | P5_G2 (+)"
    Returns: set of ROI IDs (e.g., {'P1_G1', 'P2_G3', 'P5_G2'})
    """
    if pd.isna(roi_string):
        return set()
    
    rois = set()
    for item in str(roi_string).split(' | '):
        # Extract just the ROI ID (before the sign indicator)
        roi_id = item.strip().rsplit(' ', 1)[0]  # Remove ' (+)' or ' (-)'
        rois.add(roi_id)
    return rois

def greedy_set_cover(df, max_sequences=None, min_sequence_length=4, min_roi_coverage=3, target_coverage=95.0, length_weight=0.5):
    """
    Optimized greedy algorithm to find minimal set of sequences that cover maximum ROIs.
    Balances between ROI coverage and sequence length to prefer longer, more specific patterns.
    
    Args:
        df: DataFrame with columns ['Exact Sign Sequence', 'Participating ROIs', 'Sequence Length', 'Unique Patients Count']
        max_sequences: Maximum number of sequences to select (None = no limit)
        min_sequence_length: Only consider sequences with at least this many proteins
        min_roi_coverage: Only consider sequences that appear in at least this many ROIs
        target_coverage: Stop when reaching this percentage of ROI coverage (default 95%)
        length_weight: Weight for sequence length (0=ignore length, 1=prioritize length heavily)
    
    Returns:
        List of selected sequence indices and coverage statistics
    """
    # Filter sequences by length and ROI coverage
    df_filtered = df[
        (df['Sequence Length'] >= min_sequence_length) &
        (df['Unique Patients Count'] >= min_roi_coverage)
    ].copy()
    
    print(f"\nFiltered to {len(df_filtered)} sequences (length≥{min_sequence_length}, ROIs≥{min_roi_coverage})")
    
    if len(df_filtered) == 0:
        print("Warning: No sequences meet the filtering criteria.")
        return [], {}
    
    # Parse ROI coverage for each sequence
    print("Parsing ROI coverage...")
    df_filtered['ROI_Set'] = df_filtered['Participating ROIs'].apply(parse_participating_rois)
    
    # Get all unique ROIs in the dataset
    all_rois = set()
    for roi_set in df_filtered['ROI_Set']:
        all_rois.update(roi_set)
    
    total_rois = len(all_rois)
    print(f"Total unique ROIs in dataset: {total_rois}")
    print(f"Target coverage: {target_coverage}% ({int(total_rois * target_coverage / 100)} ROIs)")
    
    # Sort sequences by ROI count (descending) for faster convergence
    df_filtered['ROI_Count'] = df_filtered['ROI_Set'].apply(len)
    df_filtered = df_filtered.sort_values('ROI_Count', ascending=False)
    
    # Greedy Set Cover Algorithm (Optimized with Length Weighting)
    covered_rois = set()
    selected_sequences = []
    remaining_indices = list(df_filtered.index)
    
    # Pre-compute for faster lookup
    roi_sets = {idx: df_filtered.loc[idx, 'ROI_Set'] for idx in remaining_indices}
    seq_lengths = {idx: df_filtered.loc[idx, 'Sequence Length'] for idx in remaining_indices}
    
    iteration = 0
    while remaining_indices and (max_sequences is None or len(selected_sequences) < max_sequences):
        iteration += 1
        
        # Find sequence with best score (coverage + length bonus)
        best_idx = None
        best_score = -1
        best_new_coverage = 0
        best_roi_set = set()
        
        # Optimize: only check top candidates based on potential coverage
        check_limit = min(10000, len(remaining_indices))  # Only check top 10K candidates per iteration
        
        for idx in remaining_indices[:check_limit]:
            roi_set = roi_sets[idx]
            new_coverage = len(roi_set - covered_rois)  # How many NEW ROIs does this cover?
            
            # Calculate weighted score: coverage + (length bonus)
            # Length bonus scales with sequence length
            length_bonus = (seq_lengths[idx] - min_sequence_length) * length_weight
            score = new_coverage + length_bonus
            
            if score > best_score:
                best_idx = idx
                best_score = score
                best_new_coverage = new_coverage
                best_roi_set = roi_set
                
                # If this covers a lot, don't waste time checking others
                if new_coverage > 50:
                    break
        
        # If no new ROIs can be covered, stop
        if best_new_coverage == 0:
            print(f"\nStopping: No more ROIs can be covered after {len(selected_sequences)} sequences.")
            break
        
        # Add the best sequence
        selected_sequences.append(best_idx)
        covered_rois.update(best_roi_set)
        remaining_indices.remove(best_idx)
        del roi_sets[best_idx]  # Remove from cache
        
        seq_len = seq_lengths[best_idx]
        coverage_pct = len(covered_rois) / total_rois * 100
        print(f"  [{iteration}] Len={seq_len} Seq {best_idx}: +{best_new_coverage} ROIs → {len(covered_rois)}/{total_rois} ({coverage_pct:.2f}%)")
        
        # Stop if we've covered all ROIs or reached target
        if len(covered_rois) >= total_rois * target_coverage / 100:
            print(f"\n✓ Target coverage ({target_coverage}%) achieved with {len(selected_sequences)} sequences!")
            break
    
    stats = {
        'total_rois': total_rois,
        'covered_rois': len(covered_rois),
        'coverage_percentage': len(covered_rois) / total_rois * 100 if total_rois > 0 else 0,
        'num_sequences': len(selected_sequences),
        'uncovered_rois': all_rois - covered_rois
    }
    
    return selected_sequences, stats

def export_results(df, selected_indices, stats):
    """Export the minimal set cover results."""
    if not selected_indices:
        print("No sequences selected. Nothing to export.")
        return
    
    # Create result DataFrame
    result_df = df.loc[selected_indices].copy()
    result_df['Selection_Order'] = range(1, len(selected_indices) + 1)
    
    # Reorder columns
    cols = ['Selection_Order', 'Sequence Length', 'Unique Patients Count', 
            'Exact Sign Sequence', 'Participating ROIs']
    result_df = result_df[cols]
    
    # Create output folder if needed
    if not os.path.exists('Docs'):
        os.makedirs('Docs')
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # Export to CSV
    csv_filename = os.path.join('Docs', f'Minimal_Set_Cover_{timestamp}.csv')
    result_df.to_csv(csv_filename, index=False)
    print(f"\n✓ Results saved to: {csv_filename}")
    
    # Export to Excel
    excel_filename = os.path.join('Docs', f'Minimal_Set_Cover_{timestamp}.xlsx')
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
        result_df.to_excel(writer, sheet_name='Minimal Set Cover', index=False)
        
        # Add summary sheet
        summary_data = {
            'Metric': [
                'Total Sequences Selected',
                'Total ROIs in Dataset',
                'ROIs Covered',
                'Coverage Percentage',
                'Uncovered ROIs'
            ],
            'Value': [
                stats['num_sequences'],
                stats['total_rois'],
                stats['covered_rois'],
                f"{stats['coverage_percentage']:.2f}%",
                len(stats['uncovered_rois'])
            ]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        
        # Add uncovered ROIs sheet if any
        if stats['uncovered_rois']:
            uncovered_df = pd.DataFrame({
                'Uncovered ROI': sorted(list(stats['uncovered_rois']))
            })
            uncovered_df.to_excel(writer, sheet_name='Uncovered ROIs', index=False)
    
    print(f"✓ Results saved to: {excel_filename}")
    
    # Print summary
    print("\n" + "="*60)
    print("MINIMAL SET COVER SUMMARY")
    print("="*60)
    print(f"Selected Sequences: {stats['num_sequences']}")
    print(f"Coverage: {stats['covered_rois']}/{stats['total_rois']} ROIs ({stats['coverage_percentage']:.2f}%)")
    if stats['uncovered_rois']:
        print(f"Uncovered ROIs: {len(stats['uncovered_rois'])}")
    print("="*60)

def main():
    print("="*60)
    print("MINIMAL SET COVER FOR BIOLOGICAL SEQUENCES")
    print("="*60)
    
    # Load data
    df = load_results_file()
    if df is None:
        return
    
    print(f"\nLoaded {len(df)} sequences from previous analysis.")
    
    # Parameters
    min_sequence_length = 4   # Minimum proteins in sequence
    min_roi_coverage = 3       # Minimum ROIs that must contain the sequence
    max_sequences = None       # None = find minimal set automatically (will find all needed)
    target_coverage = 100.0    # Full coverage - find ALL sequences needed
    length_weight = 2.0        # Weight for preferring longer sequences (0=ignore, 1=moderate, 2=strong)
    
    print(f"\nParameters:")
    print(f"  - Minimum sequence length: {min_sequence_length} proteins")
    print(f"  - Minimum ROI coverage: {min_roi_coverage} ROIs")
    print(f"  - Maximum sequences: {'Unlimited (find minimal)' if max_sequences is None else max_sequences}")
    print(f"  - Target coverage: {target_coverage}%")
    print(f"  - Length weight: {length_weight} (higher = prefer longer sequences)")
    
    # Run greedy set cover
    print("\n" + "-"*60)
    print("Running Greedy Set Cover Algorithm...")
    print("-"*60)
    
    selected_indices, stats = greedy_set_cover(
        df, 
        max_sequences=max_sequences,
        min_sequence_length=min_sequence_length,
        min_roi_coverage=min_roi_coverage,
        target_coverage=target_coverage,
        length_weight=length_weight
    )
    
    # Export results
    if selected_indices:
        export_results(df, selected_indices, stats)
    else:
        print("\nNo sequences selected. Check your filtering parameters.")

if __name__ == "__main__":
    main()
