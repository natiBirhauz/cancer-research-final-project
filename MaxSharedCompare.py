import os
import glob
import re
from datetime import datetime
import pandas as pd
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
import openpyxl
from openpyxl.styles import Alignment

def get_unique_proteins(folder_path):
    """
    Scans all dataset files to build a universal list of unique proteins.
    This creates the standardized feature space for vector comparisons.
    """
    all_files = glob.glob(os.path.join(folder_path, '*.csv')) + \
                glob.glob(os.path.join(folder_path, '*.xlsx'))
    unique_proteins = set()
    
    for file in all_files:
        try:
            df = pd.read_csv(file) if file.lower().endswith('.csv') else pd.read_excel(file)
            # Find columns that represent protein features (excluding the numeric 'Value' columns)
            protein_cols = [col for col in df.columns if 'value' not in col.lower()]
            
            for col in protein_cols:
                proteins = df[col].dropna().astype(str).str.strip().tolist()
                unique_proteins.update([p for p in proteins if p])
        except Exception as e:
            print(f"Warning: Could not process {file}. Error: {e}")
            
    return sorted(list(unique_proteins))

def parse_patient_matrix(file_path, unified_proteins):
    """
    Converts a patient's raw data file into a mathematical matrix based on the unified protein list.
    Preserves negative values to allow for anti-correlation detection.
    """
    df = pd.read_csv(file_path) if file_path.lower().endswith('.csv') else pd.read_excel(file_path)
    
    # Identify all G-components (e.g., G1, G2, etc.)
    g_cols = [col for col in df.columns if re.match(r'G\d+[+-]', col)]
    g_numbers = sorted(list(set([int(re.search(r'G(\d+)', col).group(1)) for col in g_cols])))
    
    matrix = np.zeros((len(g_numbers), len(unified_proteins)))
    prot_idx = {prot: idx for idx, prot in enumerate(unified_proteins)}
    
    for i, g_num in enumerate(g_numbers):
        plus_col = f"G{g_num}+"
        plus_val_col = [c for c in df.columns if c.startswith(f"G{g_num}+") and 'val' in c.lower()]
        minus_col = f"G{g_num}-"
        minus_val_col = [c for c in df.columns if c.startswith(f"G{g_num}-") and 'val' in c.lower()]
        
        # Process positive drivers
        if plus_col in df.columns and plus_val_col:
            for _, row in df[[plus_col, plus_val_col[0]]].dropna().iterrows():
                prot, val = str(row[plus_col]).strip(), float(row[plus_val_col[0]])
                if prot in prot_idx: 
                    matrix[i, prot_idx[prot]] += val
                    
        # Process negative drivers (values are appended as-is to preserve raw data signs)
        if minus_col in df.columns and minus_val_col:
            for _, row in df[[minus_col, minus_val_col[0]]].dropna().iterrows():
                prot, val = str(row[minus_col]).strip(), float(row[minus_val_col[0]])
                if prot in prot_idx: 
                    matrix[i, prot_idx[prot]] += val 
                    
    return matrix, g_numbers

def get_shared_driving_proteins(vec_a, vec_b, unified_proteins, top_n=7):
    """
    Extracts the most influential proteins that are actively expressed in BOTH vectors.
    Formats the output for the final report.
    """
    shared_proteins = []
    for idx, prot in enumerate(unified_proteins):
        val_a, val_b = vec_a[idx], vec_b[idx]
        
        # Protein must be active in both patients to be considered a shared driver
        if val_a != 0 and val_b != 0:
            shared_proteins.append({
                'protein': prot,
                'impact': abs(val_a) + abs(val_b),
                'val_a': val_a, 
                'val_b': val_b
            })
            
    shared_proteins.sort(key=lambda x: x['impact'], reverse=True)
    return " | ".join([f"{p['protein']} ({p['val_a']:.2f}, {p['val_b']:.2f})" for p in shared_proteins[:top_n]])

def has_sufficient_overlap(vec_a, vec_b, threshold_ratio=0.10):
    """
    Validates that the vectors share a minimum percentage of active proteins,
    preventing high cosine similarity scores driven by empty features.
    """
    active_a = set(np.where(vec_a != 0)[0])
    active_b = set(np.where(vec_b != 0)[0])
    
    if not active_a or not active_b:
        return False
        
    common_proteins = active_a.intersection(active_b)
    min_len = min(len(active_a), len(active_b))
    
    return (len(common_proteins) / min_len) >= threshold_ratio

def extract_patient_id(filename):
    """Cleans the filename to keep just the patient identifier."""
    return filename.replace('tail_proteins_', '').replace('.xlsx', '').replace('.csv', '')

def analyze_maximal_sequences(folder_path, similarity_threshold=0.80, overlap_threshold=0.10):
    """
    Main pipeline: Analyzes all patient profiles and extracts the maximal 
    shared biological sequences using a greedy 1-to-1 matching algorithm.
    """
    print("Initializing environment and building unified protein dictionary...")
    unified_proteins = get_unique_proteins(folder_path)
    
    all_files = glob.glob(os.path.join(folder_path, '*.csv')) + glob.glob(os.path.join(folder_path, '*.xlsx'))
    
    print("Parsing patient matrices...")
    patients_data = {}
    for file in all_files:
        p_name = extract_patient_id(os.path.basename(file))
        matrix, g_numbers = parse_patient_matrix(file, unified_proteins)
        patients_data[p_name] = {'g_nums': g_numbers, 'matrix': matrix}
        
    print(f"Executing sequence matching (Sim >= {similarity_threshold*100}%, Overlap >= {overlap_threshold*100}%)...")
    patient_names = list(patients_data.keys())
    results = []
    
    for i in range(len(patient_names)):
        for j in range(i + 1, len(patient_names)):
            pA, pB = patient_names[i], patient_names[j]
            matA, gA = patients_data[pA]['matrix'], patients_data[pA]['g_nums']
            matB, gB = patients_data[pB]['matrix'], patients_data[pB]['g_nums']
            
            if len(matA) == 0 or len(matB) == 0: 
                continue
            
            # Calculate cosine similarity matrix between all components of Patient A and Patient B
            sim_matrix = cosine_similarity(matA, matB)
            
            # Flatten and sort correlations
            sim_list = []
            for r in range(len(gA)):
                for c in range(len(gB)):
                    sim_list.append((abs(sim_matrix[r, c]), sim_matrix[r, c], r, c))
            sim_list.sort(key=lambda x: x[0], reverse=True)
            
            # Greedy 1-to-1 matching to avoid double counting components
            used_A, used_B = set(), set()
            shared_components = []
            
            for abs_sim, raw_sim, r, c in sim_list:
                if abs_sim >= similarity_threshold:
                    if r not in used_A and c not in used_B:
                        if has_sufficient_overlap(matA[r], matB[c], overlap_threshold):
                            used_A.add(r)
                            used_B.add(c)
                            
                            top_prots = get_shared_driving_proteins(matA[r], matB[c], unified_proteins)
                            match_type = "Direct (+)" if raw_sim > 0 else "Inverse (-)"
                            
                            shared_components.append({
                                'G_A': f"G{gA[r]}",
                                'G_B': f"G{gB[c]}",
                                'Sim': abs_sim * 100,
                                'Type': match_type,
                                'Proteins': top_prots
                            })
            
            # Aggregate findings for this patient pair
            if shared_components:
                avg_sim = np.mean([x['Sim'] for x in shared_components])
                details_text = "\n".join([f"{x['G_A']} <-> {x['G_B']} ({x['Sim']:.1f}%, {x['Type']}) | {x['Proteins']}" for x in shared_components])
                
                results.append({
                    'Patient A': pA,
                    'Patient B': pB,
                    'Maximal Shared Sequences': len(shared_components),
                    'Average Similarity (%)': avg_sim,
                    'Match Details (Components & Shared Proteins)': details_text
                })

    if not results:
        print("No matches found meeting the current threshold criteria.")
        return

    print("Compiling final report...")
    df_results = pd.DataFrame(results)
    df_results = df_results.sort_values(by=['Maximal Shared Sequences', 'Average Similarity (%)'], ascending=[False, False])
    
    # Generate timestamped filename to avoid PermissionError if an older file is open
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f'Maximal_Sequences_Report_{timestamp}.xlsx'
    
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
        df_results.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']
        # Enable text wrapping for the details column
        for row in worksheet.iter_rows(min_row=2, max_col=5, max_row=len(df_results)+1):
            row[4].alignment = Alignment(wrap_text=True)

    print(f"Success! Report generated: {excel_filename}")

if __name__ == "__main__":
    # Define dataset location and parameters
    DATASET_FOLDER = 'tails' 
    analyze_maximal_sequences(DATASET_FOLDER, similarity_threshold=0.80, overlap_threshold=0.10)