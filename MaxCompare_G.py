"""
MaxSharedCompare.py
-------------------
Identifies maximal shared biological sequences across patient profiles.
Uses FULL DENSE VECTORS, dynamically Normalizes each patient's data (MaxAbs),
and extracts proteins with the HIGHEST MATCH PERCENTAGE.
"""

import os
import glob
import re
from datetime import datetime
import pandas as pd
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
import openpyxl
from openpyxl.styles import Alignment

def get_unique_proteins(g_folder_path):
    all_files = glob.glob(os.path.join(g_folder_path, '*.csv')) + \
                glob.glob(os.path.join(g_folder_path, '*.xlsx'))
    unique_proteins = set()
    
    for file in all_files:
        try:
            df = pd.read_csv(file) if file.lower().endswith('.csv') else pd.read_excel(file)
            prot_col = df.columns[0] 
            proteins = df[prot_col].dropna().astype(str).str.strip().tolist()
            unique_proteins.update([p for p in proteins if p])
        except Exception as e:
            print(f"Warning: Could not process {file}. Error: {e}")
            
    return sorted(list(unique_proteins))

def parse_dense_patient_matrix(tail_file_path, g_file_path, unified_proteins):
    df_tail = pd.read_csv(tail_file_path) if str(tail_file_path).lower().endswith('.csv') else pd.read_excel(tail_file_path)
    g_cols = [col for col in df_tail.columns if re.match(r'G\d+[+-]', col)]
    g_numbers = sorted(list(set([int(re.search(r'G(\d+)', col).group(1)) for col in g_cols])))
    
    df_g = pd.read_csv(g_file_path) if str(g_file_path).lower().endswith('.csv') else pd.read_excel(g_file_path)
    prot_col = df_g.columns[0]
    df_g.set_index(prot_col, inplace=True)
    df_g.index = df_g.index.astype(str).str.strip()
    
    matrix = np.zeros((len(g_numbers), len(unified_proteins)))
    prot_idx = {prot: idx for idx, prot in enumerate(unified_proteins)}
    
    for i, g_num in enumerate(g_numbers):
        g_col_name = f"G{g_num}"
        if g_col_name in df_g.columns:
            for prot in unified_proteins:
                if prot in df_g.index:
                    val = df_g.loc[prot, g_col_name]
                    if isinstance(val, pd.Series):
                        val = val.iloc[0]
                    matrix[i, prot_idx[prot]] = float(val)
                    
    return matrix, g_numbers

def get_highest_matching_proteins(vec_a, vec_b, unified_proteins, is_direct, top_n=10, min_norm_threshold=0.05):
    """
    Normalizes both vectors to [-1.0, 1.0] to ensure fair comparison,
    then extracts proteins sorted by their MATCH PERCENTAGE.
    """
    matching_proteins = []
    
    # --- שלב הנרמול (MaxAbs Scaler) לכל חולה בנפרד ---
    max_a = np.max(np.abs(vec_a))
    max_b = np.max(np.abs(vec_b))
    
    # מונע חלוקה באפס אם כל הוקטור ריק
    norm_vec_a = vec_a / max_a if max_a > 0 else vec_a
    norm_vec_b = vec_b / max_b if max_b > 0 else vec_b
    
    for idx, prot in enumerate(unified_proteins):
        val_a = norm_vec_a[idx]
        val_b = norm_vec_b[idx]
        
        # סינון רעשי רקע: עובד כעת על האחוזים המנורמלים.
        # min_norm_threshold=0.05 אומר שלפחות אחד החולים מבטא את החלבון ב-5% מהעוצמה המקסימלית שלו
        if max(abs(val_a), abs(val_b)) >= min_norm_threshold:
            
            # אם ההתאמה היא הפוכה (Inverse), נשווה למינוס הערך של חולה ב'
            compare_b = val_b if is_direct else -val_b
            
            # חישוב אחוז ההתאמה בין שני הערכים המנורמלים
            max_val = max(abs(val_a), abs(compare_b))
            diff = abs(val_a - compare_b)
            
            if max_val > 0:
                match_pct = (1.0 - (diff / max_val)) * 100
            else:
                match_pct = 0.0
                
            if match_pct > 0:
                matching_proteins.append({
                    'protein': prot,
                    'match_pct': match_pct,
                    'val_a': val_a, 
                    'val_b': val_b
                })
                
    matching_proteins.sort(key=lambda x: (x['match_pct'], abs(x['val_a'])), reverse=True)
    
    formatted_prots = []
    for p in matching_proteins[:top_n]:
        # שיניתי את התצוגה ל- Norm_A / Norm_B כדי שיהיה ברור שמדובר בערכים מנורמלים
        formatted_prots.append(f"{p['protein']} ({p['match_pct']:.1f}% Match | Norm_A: {p['val_a']:+.3f}, Norm_B: {p['val_b']:+.3f})")
        
    return " | ".join(formatted_prots) if formatted_prots else "No significant matching proteins"

def extract_patient_id(filename):
    name = filename.replace('_tail_proteins', '').replace('_G', '')
    name = name.replace('.xlsx', '').replace('.csv', '')
    return name

def analyze_maximal_sequences(tails_folder, g_folder, similarity_threshold=0.80):
    print("Building universal protein dictionary from full G tables...")
    unified_proteins = get_unique_proteins(g_folder)
    
    tail_files = glob.glob(os.path.join(tails_folder, '*.csv')) + glob.glob(os.path.join(tails_folder, '*.xlsx'))
    g_files = glob.glob(os.path.join(g_folder, '*.csv')) + glob.glob(os.path.join(g_folder, '*.xlsx'))
    
    g_file_map = {extract_patient_id(os.path.basename(f)): f for f in g_files}
    
    print("Parsing full patient matrices...")
    patients_data = {}
    for tail_file in tail_files:
        p_name = extract_patient_id(os.path.basename(tail_file))
        if p_name in g_file_map:
            g_file = g_file_map[p_name]
            matrix, g_numbers = parse_dense_patient_matrix(tail_file, g_file, unified_proteins)
            patients_data[p_name] = {'g_nums': g_numbers, 'matrix': matrix}
            
    print(f"Executing sequence matching based on normalized percentage similarity...")
    patient_names = list(patients_data.keys())
    results = []
    
    for i in range(len(patient_names)):
        for j in range(i + 1, len(patient_names)):
            pA, pB = patient_names[i], patient_names[j]
            matA, gA = patients_data[pA]['matrix'], patients_data[pA]['g_nums']
            matB, gB = patients_data[pB]['matrix'], patients_data[pB]['g_nums']
            
            if len(matA) == 0 or len(matB) == 0: 
                continue
            
            sim_matrix = cosine_similarity(matA, matB)
            
            sim_list = []
            for r in range(len(gA)):
                for c in range(len(gB)):
                    sim_list.append((abs(sim_matrix[r, c]), sim_matrix[r, c], r, c))
            sim_list.sort(key=lambda x: x[0], reverse=True)
            
            used_A, used_B = set(), set()
            shared_components = []
            
            for abs_sim, raw_sim, r, c in sim_list:
                if abs_sim >= similarity_threshold:
                    if r not in used_A and c not in used_B:
                        used_A.add(r)
                        used_B.add(c)
                        
                        is_direct = (raw_sim > 0)
                        match_type = "Direct (+)" if is_direct else "Inverse (-)"
                        
                        # הפונקציה שלנו כעת מנרמלת אוטומטית לפני החישוב
                        top_prots = get_highest_matching_proteins(matA[r], matB[c], unified_proteins, is_direct)
                        
                        shared_components.append({
                            'G_A': f"G{gA[r]}",
                            'G_B': f"G{gB[c]}",
                            'Sim': abs_sim * 100,
                            'Type': match_type,
                            'Proteins': top_prots
                        })
            
            if shared_components:
                avg_sim = np.mean([x['Sim'] for x in shared_components])
                details_text = "\n".join([f"{x['G_A']} <-> {x['G_B']} ({x['Sim']:.1f}%, {x['Type']}) | {x['Proteins']}" for x in shared_components])
                
                results.append({
                    'Patient A': pA,
                    'Patient B': pB,
                    'Maximal Shared Sequences': len(shared_components),
                    'Average Similarity (%)': avg_sim,
                    'Match Details (Normalized Protein Match %)': details_text
                })

    if not results:
        print("No matches found.")
        return

    print("Compiling final report...")
    df_results = pd.DataFrame(results)
    df_results = df_results.sort_values(by=['Maximal Shared Sequences', 'Average Similarity (%)'], ascending=[False, False])
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f'Maximal_Sequences_Report_{timestamp}.xlsx'
    
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
        df_results.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']
        for row in worksheet.iter_rows(min_row=2, max_col=5, max_row=len(df_results)+1):
            row[4].alignment = Alignment(wrap_text=True)

    print(f"Success! Report generated: {excel_filename}")

if __name__ == "__main__":
    TAILS_FOLDER = 'All_Tail_Proteins_Tables'
    G_FOLDER = 'All_G_Tables' 
    analyze_maximal_sequences(TAILS_FOLDER, G_FOLDER, similarity_threshold=0.80)