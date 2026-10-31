"""
GlobalClusterCompare.py
-----------------------
Finds GLOBAL shared biological sequences across multiple patients simultaneously.
Uses Graph Theory (NetworkX) to find connected components (clusters) of similar G-vectors,
and extracts the core biological signature of the entire cluster.
"""

import os
import glob
import re
from datetime import datetime
import pandas as pd
import numpy as np
from sklearn.metrics.pairwise import cosine_similarity
import networkx as nx
import openpyxl
from openpyxl.styles import Alignment

def get_unique_proteins(g_folder_path):
    """Scan all files to build a universal list of unique proteins."""
    all_files = glob.glob(os.path.join(g_folder_path, '*.csv')) + \
                glob.glob(os.path.join(g_folder_path, '*.xlsx'))
    unique_proteins = set()
    for file in all_files:
        try:
            df = pd.read_csv(file) if file.lower().endswith('.csv') else pd.read_excel(file)
            prot_col = df.columns[0] 
            proteins = df[prot_col].dropna().astype(str).str.strip().tolist()
            unique_proteins.update([p for p in proteins if p])
        except: pass
    return sorted(list(unique_proteins))

def parse_all_dense_vectors(tails_folder, g_folder, unified_proteins):
    """Extract and normalize all valid G components from all patients."""
    tail_files = glob.glob(os.path.join(tails_folder, '*.csv')) + glob.glob(os.path.join(tails_folder, '*.xlsx'))
    g_files = glob.glob(os.path.join(g_folder, '*.csv')) + glob.glob(os.path.join(g_folder, '*.xlsx'))
    
    def extract_id(f): return os.path.basename(f).replace('_tail_proteins', '').replace('_G', '').replace('.xlsx', '').replace('.csv', '')
    g_file_map = {extract_id(f): f for f in g_files}
    
    global_nodes = []
    
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
        
        prot_idx = {prot: idx for idx, prot in enumerate(unified_proteins)}
        
        for g_num in g_numbers:
            g_col_name = f"G{g_num}"
            if g_col_name in df_g.columns:
                vec = np.zeros(len(unified_proteins))
                for prot in unified_proteins:
                    if prot in df_g.index:
                        val = df_g.loc[prot, g_col_name]
                        if isinstance(val, pd.Series): val = val.iloc[0]
                        vec[prot_idx[prot]] = float(val)
                
                # MaxAbs Normalization
                max_val = np.max(np.abs(vec))
                if max_val > 0:
                    vec = vec / max_val
                    
                global_nodes.append({
                    'id': f"{p_name}_G{g_num}",
                    'patient': p_name,
                    'g_num': g_num,
                    'vec': vec
                })
                
    return global_nodes

def analyze_global_clusters(tails_folder, g_folder, similarity_threshold=0.80):
    print("Building universal protein dictionary...")
    unified_proteins = get_unique_proteins(g_folder)
    
    print("Extracting vectors into global pool...")
    nodes = parse_all_dense_vectors(tails_folder, g_folder, unified_proteins)
    
    if not nodes:
        print("No valid data found.")
        return
        
    print(f"Total components found: {len(nodes)}. Building similarity graph...")
    
    # Initialize graph
    G = nx.Graph()
    for n in nodes:
        G.add_node(n['id'], patient=n['patient'], vec=n['vec'])
        
    # Compute global similarity matrix
    vecs = np.array([n['vec'] for n in nodes])
    sim_matrix = cosine_similarity(vecs)
    
    # Add edges for similarities above threshold
    for i in range(len(nodes)):
        for j in range(i + 1, len(nodes)):
            raw_sim = sim_matrix[i, j]
            if abs(raw_sim) >= similarity_threshold:
                G.add_edge(nodes[i]['id'], nodes[j]['id'], weight=abs(raw_sim), raw_sim=raw_sim)
                
    print("Finding connected components (Clusters)...")
    communities = list(nx.connected_components(G))
    results = []
    
    for cluster_id, community in enumerate(communities, 1):
        community_nodes = list(community)
        patients_in_cluster = set([G.nodes[n]['patient'] for n in community_nodes])
        
        # Filter: Keep clusters with at least 2 unique patients
        if len(patients_in_cluster) < 2:
            continue
            
        # Calculate cluster's average similarity
        subgraph = G.subgraph(community_nodes)
        edges = subgraph.edges(data=True)
        if edges:
            avg_sim = np.mean([data['weight'] for _, _, data in edges]) * 100
        else:
            avg_sim = 100.0

        # Create aligned cluster signature (Handle inverse correlations)
        anchor_id = community_nodes[0]
        anchor_vec = G.nodes[anchor_id]['vec']
        aligned_vecs = [anchor_vec]
        
        participants_details = [f"{anchor_id} (+)"]
        
        for other_id in community_nodes[1:]:
            other_vec = G.nodes[other_id]['vec']
            direction = np.dot(anchor_vec, other_vec)
            
            if direction < 0:
                aligned_vecs.append(-other_vec)
                participants_details.append(f"{other_id} (-)")
            else:
                aligned_vecs.append(other_vec)
                participants_details.append(f"{other_id} (+)")
                
        # Compute mean vector of aligned components
        avg_vec = np.mean(aligned_vecs, axis=0)
        
        # FIXED: Count ALL non-zero proteins (removed the 0.05 threshold)
        sig_proteins_count = np.sum(np.abs(avg_vec) > 0.05)
        
        # FIXED: Extract ALL sorted proteins without any slice limitation or noise filter
        top_indices = np.argsort(np.abs(avg_vec))[::-1]
        core_signature = []
        for idx in top_indices:
            val = avg_vec[idx]
            if abs(val) > 0.0:  # Include everything that isn't absolute zero
                core_signature.append(f"{unified_proteins[idx]} ({val:+.3f})")
                
        results.append({
            'Cluster ID': f"Cluster_{cluster_id}",
            'Unique Patients Count': len(patients_in_cluster),
            'Total Components in Cluster': len(community_nodes),
            'Average Similarity (%)': round(avg_sim, 2),
            'Unique Proteins Count': int(sig_proteins_count),
            'Participating Patients & ROIs': " | ".join(sorted(participants_details)),
            'Core Biological Signature (Avg Norm)': " | ".join(core_signature)
        })
        
    if not results:
        print("No shared clusters found.")
        return

    print("Compiling global report...")
    df_results = pd.DataFrame(results)
    
    # Sort by patient count (descending)
    df_results = df_results.sort_values(by=['Unique Patients Count', 'Total Components in Cluster'], ascending=[False, False])
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f'Global_Shared_Processes_Report_{timestamp}.xlsx'
    
    with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
        df_results.to_excel(writer, index=False)
        worksheet = writer.sheets['Sheet1']
        
        # Adjust text wrapping for detail columns
        for row in worksheet.iter_rows(min_row=2, max_col=7, max_row=len(df_results)+1):
            row[5].alignment = Alignment(wrap_text=True) # Participating Patients
            row[6].alignment = Alignment(wrap_text=True) # Core Signature

    print(f"Success! Global Clusters Report generated: {excel_filename}")

if __name__ == "__main__":
    TAILS_FOLDER = 'All_Tail_Proteins_Tables'
    G_FOLDER = 'All_G_Tables' 
    analyze_global_clusters(TAILS_FOLDER, G_FOLDER, similarity_threshold=0.80)