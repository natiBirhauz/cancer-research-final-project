"""
GlobalClusterReverseSign.py
---------------------------
Finds the LONGEST EXACT sequences of proteins with consistent signs (+/-).
Handles massive datasets by exporting to CSV and limiting Excel to Top 1M rows.
"""

import os
import glob
import re
from datetime import datetime
import pandas as pd
import numpy as np

def get_valid_paths():
    """Smart path finder to locate the data folders whether they are in the root or Docs folder."""
    tails_candidates = ['All_Tail_Proteins_Tables', os.path.join('Docs', 'All_Tail_Proteins_Tables')]
    g_candidates = ['All_G_Tables', os.path.join('Docs', 'All_G_Tables')]
    
    tails_path = next((p for p in tails_candidates if os.path.exists(p)), None)
    g_path = next((p for p in g_candidates if os.path.exists(p)), None)
    
    return tails_path, g_path

def load_all_vectors(tails_folder, g_folder, noise_cutoff=0.05):
    """Loads Tail proteins (candidates) and G-components (actual signs)."""
    g_files = glob.glob(os.path.join(g_folder, '*.csv')) + glob.glob(os.path.join(g_folder, '*.xlsx'))
    tail_files = glob.glob(os.path.join(tails_folder, '*.csv')) + glob.glob(os.path.join(tails_folder, '*.xlsx'))
    
    def extract_id(f): return os.path.basename(f).replace('_tail_proteins', '').replace('_G', '').replace('.xlsx', '').replace('.csv', '')
    g_file_map = {extract_id(f): f for f in g_files}
    
    vectors = []
    
    for tail_file in tail_files:
        p_name = extract_id(tail_file)
        if p_name not in g_file_map: continue
            
        df_tail = pd.read_csv(tail_file) if str(tail_file).lower().endswith('.csv') else pd.read_excel(tail_file)
        g_cols = [col for col in df_tail.columns if re.match(r'G\d+[+-]', col)]
        g_numbers = sorted(list(set([int(re.search(r'G(\d+)', col).group(1)) for col in g_cols])))
        
        df_g = pd.read_csv(g_file_map[p_name]) if str(g_file_map[p_name]).lower().endswith('.csv') else pd.read_excel(g_file_map[p_name])
        prot_col = df_g.columns[0]
        df_g.set_index(prot_col, inplace=True)
        df_g.index = df_g.index.astype(str).str.strip()
        
        for g_num in g_numbers:
            tail_proteins = set()
            col_plus = f"G{g_num}+"
            col_minus = f"G{g_num}-"
            
            if col_plus in df_tail.columns:
                tail_proteins.update(df_tail[col_plus].dropna().astype(str).str.strip().tolist())
            if col_minus in df_tail.columns:
                tail_proteins.update(df_tail[col_minus].dropna().astype(str).str.strip().tolist())

            g_col_name = f"G{g_num}"
            if g_col_name in df_g.columns:
                signs_dict = {}
                for prot in df_g.index:
                    val = df_g.loc[prot, g_col_name]
                    if isinstance(val, pd.Series): val = val.iloc[0]
                    val = float(val)
                    
                    if val > noise_cutoff:
                        signs_dict[prot] = 1
                    elif val < -noise_cutoff:
                        signs_dict[prot] = -1
                        
                if signs_dict:
                    vectors.append({
                        'id': f"{p_name}_G{g_num}",
                        'patient': p_name,
                        'tails': tail_proteins, 
                        'signs': signs_dict     
                    })
    return vectors

def canonicalize_motif(motif_dict):
    """Ensures an inverse sequence generates the exact same string."""
    sorted_prots = sorted(list(motif_dict.keys()))
    first_sign = motif_dict[sorted_prots[0]]
    
    if first_sign == -1:
        canon = {p: -motif_dict[p] for p in sorted_prots}
    else:
        canon = {p: motif_dict[p] for p in sorted_prots}
        
    motif_str = " | ".join([f"{p}({'+' if canon[p]==1 else '-'})" for p in sorted_prots])
    return motif_str, canon

def mine_exact_sequences(tails_folder, g_folder, min_sequence_length=4, min_patients=3):
    print("Loading biological data (Tails + G Tables)...")
    vectors = load_all_vectors(tails_folder, g_folder, noise_cutoff=0.05)
    
    if not vectors:
        print("Error: Could not load vectors. Check your data folders.")
        return
        
    print(f"Total G-components loaded: {len(vectors)}")
    print("Discovering maximal shared motifs using TAIL anchors (This may take a minute)...")
    unique_motifs_strs = set()
    
    for i in range(len(vectors)):
        s_i = vectors[i]['signs']
        t_i = vectors[i]['tails']
        
        for j in range(i + 1, len(vectors)):
            s_j = vectors[j]['signs']
            t_j = vectors[j]['tails']
            
            candidates = t_i | t_j 
            
            direct = {}
            inverse = {}
            
            for p in candidates:
                if p in s_i and p in s_j:
                    if s_i[p] == s_j[p]:
                        direct[p] = s_i[p]
                    elif s_i[p] == -s_j[p]:
                        inverse[p] = s_i[p]
            
            best_match = direct if len(direct) > len(inverse) else inverse
            
            if len(best_match) >= min_sequence_length:
                motif_str, _ = canonicalize_motif(best_match)
                unique_motifs_strs.add(motif_str)
                
    print(f"Found {len(unique_motifs_strs)} unique biological sequences. Scanning global population...")
    
    results = []
    
    for m_str in unique_motifs_strs:
        m_dict = {}
        for item in m_str.split(" | "):
            prot = item.split("(")[0]
            sign = 1 if "(+)" in item else -1
            m_dict[prot] = sign
            
        patients_with_motif = set()
        participating_gs = []
        
        for v in vectors:
            pat = v['patient']
            v_signs = v['signs']
            
            is_direct = all(p in v_signs and v_signs[p] == m_dict[p] for p in m_dict)
            is_inverse = all(p in v_signs and v_signs[p] == -m_dict[p] for p in m_dict)
            
            if is_direct:
                patients_with_motif.add(pat)
                participating_gs.append(f"{v['id']} (+)")
            elif is_inverse:
                patients_with_motif.add(pat)
                participating_gs.append(f"{v['id']} (-)")
                
        # Only save if it meets the minimum patient threshold
        if len(patients_with_motif) >= min_patients:
            results.append({
                'Sequence Length': len(m_dict),
                'Unique Patients Count': len(patients_with_motif),
                'Exact Sign Sequence': m_str,
                'Participating ROIs': " | ".join(sorted(participating_gs))
            })
            
    if not results:
        print(f"No shared sequences found for minimum {min_patients} patients.")
        return
        
    print("Sorting and exporting results...")
    df = pd.DataFrame(results)
    
    df = df.sort_values(by=['Sequence Length', 'Unique Patients Count'], ascending=[False, False])
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    
    if not os.path.exists('Docs'):
        os.makedirs('Docs')
        
    # 1. Export ALL data to CSV (No row limit)
    csv_filename = os.path.join('Docs', f'Exact_Sequences_Miner_{timestamp}.csv')
    df.to_csv(csv_filename, index=False)
    print(f"FULL dataset saved successfully to CSV: {csv_filename}")
    
    # 2. Export Top 1 Million to Excel
    max_excel_rows = 1000000
    if len(df) > max_excel_rows:
        print(f"\nWarning: Excel cannot handle {len(df)} rows.")
        print(f"Extracting the top {max_excel_rows} longest and most common sequences to Excel...")
        df_excel = df.head(max_excel_rows)
    else:
        df_excel = df
        
    excel_filename = os.path.join('Docs', f'Exact_Sequences_Miner_{timestamp}.xlsx')
    df_excel.to_excel(excel_filename, index=False)
    print(f"TOP results saved successfully to Excel: {excel_filename}")

if __name__ == "__main__":
    tails_path, g_path = get_valid_paths()
    
    if not tails_path or not g_path:
        print("ERROR: Could not find data folders.")
    else:
        print(f"Located Tails folder at: {tails_path}")
        print(f"Located G folder at: {g_path}\n")
        
        # INCREASED minimum patients from 2 to 3 to filter out noise
        mine_exact_sequences(tails_path, g_path, min_sequence_length=3, min_patients=3)