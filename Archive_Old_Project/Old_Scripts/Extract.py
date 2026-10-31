import os
import shutil
from pathlib import Path

def gather_and_rename_patient_data(source_base_path, output_base_path):
    """
    Recursively scans patient subdirectories, extracts both 'tail_proteins' 
    and 'G' tables, renames them using Patient and ROI IDs, and centralizes 
    them inside the Docs folder.
    """
    base_dir = Path(source_base_path)
    output_base = Path(output_base_path)
    
    # Define and create dedicated output directories inside Docs
    tails_out_dir = output_base / "All_Tail_Proteins_Tables"
    g_out_dir = output_base / "All_G_Tables"
    
    tails_out_dir.mkdir(parents=True, exist_ok=True)
    g_out_dir.mkdir(parents=True, exist_ok=True)
    
    print(f"Scanning for patient data inside: {base_dir}")
    print(f"Tails target directory: {tails_out_dir}")
    print(f"G-Tables target directory: {g_out_dir}\n")
    
    tail_count = 0
    g_count = 0
    
    # Scan recursively for all spreadsheet extensions
    for ext in ["*.csv", "*.xlsx"]:
        for file_path in base_dir.rglob(ext):
            try:
                # Get path relative to the source root to easily map structures
                rel_parts = file_path.relative_to(base_dir).parts
                
                # Dynamic mapping based on structure: Patient/ROI/surprisal analysis/file
                if len(rel_parts) >= 3:
                    patient_id = rel_parts[0]  # e.g., P20
                    roi_id = rel_parts[1]      # e.g., ROI_110
                    
                    # Validation check to ensure correct folder naming alignment
                    if patient_id.lower().startswith('p') and roi_id.lower().startswith('roi'):
                        stem = file_path.stem.lower()
                        
                        # Process tail proteins tables
                        if stem == "tail_proteins":
                            new_filename = f"{patient_id}_{roi_id}_tail_proteins{file_path.suffix}"
                            shutil.copy2(file_path, tails_out_dir / new_filename)
                            print(f"[TAIL COPIED] {patient_id} -> {roi_id} | Name: {new_filename}")
                            tail_count += 1
                            
                        # Process G tables (matches files named exactly 'G' or 'g')
                        elif stem == "g":
                            new_filename = f"{patient_id}_{roi_id}_G{file_path.suffix}"
                            shutil.copy2(file_path, g_out_dir / new_filename)
                            print(f"[G COPIED]    {patient_id} -> {roi_id} | Name: {new_filename}")
                            g_count += 1
                            
            except Exception as e:
                # Safe fallback to skip non-matching system or unreadable files
                continue

    print("\n" + "="*40)
    print("Extraction process completed successfully!")
    print(f"Total Tail Protein tables copied: {tail_count}")
    print(f"Total G tables copied: {g_count}")
    print("="*40)

if __name__ == "__main__":
    # Core project directory containing all source files
    SOURCE_PATH = r"C:\Users\nati4\Desktop\פרויקט גמר\all patients (without isotope)"
    
    # Centralized documentation output directory
    OUTPUT_PATH = r"C:\Users\nati4\Desktop\פרויקט גמר\Docs"
    
    gather_and_rename_patient_data(SOURCE_PATH, OUTPUT_PATH)