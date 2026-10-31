"""
CohortTrendVisualizer.py
------------------------
Interactive visualizer that plots the RAW G-values of specific proteins 
across multiple patients (X-axis = Patient/ROI, Y-axis = G-Value).
Perfect for proving anti-correlation (e.g., CD3 goes up while TGFB goes down).
"""

import os
import glob
import re
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

def find_latest_report():
    """Finds the most recent Exact Sequences report."""
    patterns = ['Exact_Sequences_Miner_*.xlsx', os.path.join('Docs', 'Exact_Sequences_Miner_*.xlsx')]
    files = []
    for p in patterns:
        files.extend(glob.glob(p))
    if not files:
        print("Error: Could not find Exact_Sequences_Miner report.")
        return None
    return max(files, key=os.path.getctime)

def get_g_folder():
    """Finds the raw G-tables folder."""
    candidates = ['All_G_Tables', os.path.join('Docs', 'All_G_Tables')]
    return next((p for p in candidates if os.path.exists(p)), None)

def extract_proteins_from_sequence(seq_string):
    """Extracts just the protein names from the sequence string."""
    parts = str(seq_string).split(" | ")
    proteins = []
    for part in parts:
        match = re.search(r"([A-Za-z0-9_\-]+)\s*\(", part.strip())
        if match:
            proteins.append(match.group(1))
    return proteins

def fetch_patient_data(participating_string, proteins, g_folder):
    """
    Goes into the original G-files and extracts the exact numeric G-values 
    for the requested proteins across all participating patients.
    """
    # Example format: "P20_ROI_110_G1 (+) | P21_ROI_112_G3 (-)"
    participants = str(participating_string).split(" | ")
    
    plot_data = []
    
    for p in participants:
        # Extract base name (P20_ROI_110) and G-column (G1)
        match = re.match(r"(.+)_G(\d+)\s*\(([-+])\)", p.strip())
        if not match: continue
        
        base_name, g_num, sign = match.groups()
        g_col = f"G{g_num}"
        
        # Find the patient's G file
        # It could be .csv or .xlsx
        search_pattern = os.path.join(g_folder, f"{base_name}_G.*")
        patient_files = glob.glob(search_pattern)
        
        if not patient_files:
            continue
            
        file_path = patient_files[0]
        try:
            df = pd.read_csv(file_path) if file_path.endswith('.csv') else pd.read_excel(file_path)
            prot_col = df.columns[0]
            df.set_index(prot_col, inplace=True)
            df.index = df.index.astype(str).str.strip()
            
            if g_col in df.columns:
                patient_record = {'Patient_ROI': base_name, 'Alignment': sign}
                for prot in proteins:
                    if prot in df.index:
                        val = df.loc[prot, g_col]
                        if isinstance(val, pd.Series): val = val.iloc[0]
                        
                        # INVERSE ALIGNMENT LOGIC: 
                        # If the patient is marked as (-), their entire process is mathematically flipped.
                        # To show consistent biological behavior on the graph, we flip them back.
                        val = float(val)
                        if sign == '-':
                            val = -val
                            
                        patient_record[prot] = val
                    else:
                        patient_record[prot] = 0.0
                plot_data.append(patient_record)
        except Exception as e:
            print(f"Skipping {base_name} due to read error.")
            
    return pd.DataFrame(plot_data)

def plot_cohort_trends(df_plot, row_index):
    """Plots the extracted G-values with X = Patients, Y = G-Values."""
    if df_plot.empty:
        print("No data could be extracted for this sequence.")
        return
        
    sns.set_theme(style="whitegrid")
    plt.figure(figsize=(14, 7))
    
    # We don't need these columns for the Y-axis
    proteins = [c for c in df_plot.columns if c not in ['Patient_ROI', 'Alignment']]
    
    # Plot a line with markers for each protein
    markers = ['o', 's', '^', 'D', 'v', 'p', '*']
    
    for idx, prot in enumerate(proteins):
        plt.plot(df_plot['Patient_ROI'], df_plot[prot], 
                 marker=markers[idx % len(markers)], linewidth=2, markersize=8, label=prot)
                 
    plt.axhline(0, color='black', linewidth=1.5, linestyle='--')
    
    plt.title(f"Protein Expression Trends Across Patients (Row {row_index})\nProof of Correlation / Anti-Correlation", 
              fontsize=16, fontweight='bold', pad=15)
    plt.xlabel("Patient & ROI", fontsize=12, fontweight='bold')
    plt.ylabel("G-Value (Weight)", fontsize=12, fontweight='bold')
    
    # Rotate X labels so they don't overlap if there are many patients
    plt.xticks(rotation=45, ha='right', fontsize=9)
    plt.legend(title="Proteins", bbox_to_anchor=(1.05, 1), loc='upper left')
    
    plt.tight_layout()
    plt.show()

def interactive_menu():
    report_file = find_latest_report()
    g_folder = get_g_folder()
    
    if not report_file or not g_folder:
        return
        
    df = pd.read_excel(report_file)
    
    print("=" * 70)
    print(" 📈 Cross-Patient Trend Visualizer (X=Patient, Y=G-Value) 📈")
    print("=" * 70)
    
    preview_limit = min(15, len(df))
    for idx in range(preview_limit):
        seq_snippet = str(df.loc[idx, 'Exact Sign Sequence'])[:60] + "..."
        p_count = df.loc[idx, 'Unique Patients Count']
        print(f"[{idx}] (Shared by {p_count} patients) -> {seq_snippet}")
        
    print("-" * 70)
    
    while True:
        user_input = input(f"\nEnter row index to plot (0 to {len(df)-1}) or 'q' to quit: ").strip()
        if user_input.lower() == 'q':
            break
            
        try:
            row_idx = int(user_input)
            if 0 <= row_idx < len(df):
                seq_str = df.loc[row_idx, 'Exact Sign Sequence']
                participants_str = df.loc[row_idx, 'Participating ROIs']
                
                print(f"Extracting raw data for row {row_idx}... Please wait.")
                proteins = extract_proteins_from_sequence(seq_str)
                
                # If there are too many proteins, the graph will be messy. Let's limit to top 8.
                if len(proteins) > 8:
                    print(f"Note: This sequence has {len(proteins)} proteins. Plotting the first 8 for clarity.")
                    proteins = proteins[:8]
                    
                df_plot = fetch_patient_data(participants_str, proteins, g_folder)
                
                print("Plotting... Close plot window to select another row.")
                plot_cohort_trends(df_plot, row_idx)
            else:
                print("Invalid index.")
        except ValueError:
            print("Invalid input.")

if __name__ == "__main__":
    interactive_menu()

    