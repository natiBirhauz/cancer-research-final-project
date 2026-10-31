"""
count.py
--------
Scans the entire dataset folder and generates a global summary 
of the top 20 most frequently occurring proteins across all files.
"""

import os
import glob
import pandas as pd

def get_proteins_from_file(file_path):
    """Extracts all protein names from a single file, ignoring value columns."""
    try:
        df = pd.read_csv(file_path) if file_path.lower().endswith('.csv') else pd.read_excel(file_path)
        
        # Filter out the 'Value' columns to isolate protein name columns
        protein_cols = [col for col in df.columns if 'value' not in col.lower()]
        
        file_proteins = []
        for col in protein_cols:
            proteins = df[col].dropna().astype(str).tolist()
            # Clean accidental whitespaces
            proteins = [p.strip() for p in proteins if p.strip()]
            file_proteins.extend(proteins)
            
        return file_proteins
    except Exception as e:
        print(f"[Error] Failed to read {os.path.basename(file_path)}: {e}")
        return []

def print_global_summary(folder_path, top_n=20):
    """Iterates through the directory and aggregates the total protein counts."""
    print(f"Scanning directory: '{folder_path}'...\n")
    
    all_files = glob.glob(os.path.join(folder_path, '*.csv')) + \
                glob.glob(os.path.join(folder_path, '*.xlsx'))
    
    if not all_files:
        print(f"No valid data files found in folder: {folder_path}")
        return
        
    global_proteins_list = []
    
    # Process each file and aggregate the data
    for file in all_files:
        proteins_in_file = get_proteins_from_file(file)
        global_proteins_list.extend(proteins_in_file)
            
    # Print the Global Summary
    print("=" * 50)
    print(f" GLOBAL SUMMARY ACROSS ALL {len(all_files)} FILES ")
    print("=" * 50)
    
    if global_proteins_list:
        global_counts = pd.Series(global_proteins_list).value_counts()
        print(f"\nTop {top_n} most frequent proteins in the entire dataset:\n")
        print(global_counts.head(top_n).to_string())
        print("\n" + "=" * 50)
    else:
        print("No proteins found in any of the files.")

if __name__ == "__main__":
    DATASET_FOLDER = 'tails' 
    print_global_summary(DATASET_FOLDER, top_n=20)