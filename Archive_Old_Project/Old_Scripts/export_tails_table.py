"""
export_tails_table.py
---------------------
Reads every tail_proteins.xlsx from:
  all patients (without isotope)\P*\ROI_*\surprisal analysis\tail_proteins.xlsx

Outputs a unified TAILS table to:
  Docs/All_Tails.xlsx (and Docs/All_Tails_<timestamp>.xlsx)
with one row per G-vector per ROI.

Columns:
  Patient | ROI | ROI_Key | G_Key | G_Num | Protein_Set (sorted, no signs) |
  Proteins_Plus | Proteins_Minus | Protein_Count
"""

import os
import re
import glob
import pandas as pd
from datetime import datetime

BASE_DIR   = r'all patients (without isotope)'
OUTPUT_DIR = 'Docs'


def extract_g_vectors(df: pd.DataFrame, patient: str, roi: str) -> list:
    """
    From a tail_proteins DataFrame (4 rows x G*+- cols),
    extract each G-vector as a frozenset of protein names (no signs).
    Returns list of dicts: one per G-vector that has at least one protein.
    """
    rows = []

    # Find all G numbers present in columns
    g_nums = sorted(set(
        int(m.group(1))
        for col in df.columns
        if (m := re.match(r'^G(\d+)[+-]$', col))
    ))

    for g_num in g_nums:
        col_plus  = f'G{g_num}+'
        col_minus = f'G{g_num}-'

        # Collect protein names from + column (all non-NaN rows)
        plus_proteins = []
        if col_plus in df.columns:
            plus_proteins = [
                str(v).strip()
                for v in df[col_plus].dropna()
                if str(v).strip() not in ('', 'nan')
            ]

        # Collect protein names from - column
        minus_proteins = []
        if col_minus in df.columns:
            minus_proteins = [
                str(v).strip()
                for v in df[col_minus].dropna()
                if str(v).strip() not in ('', 'nan')
            ]

        all_proteins = plus_proteins + minus_proteins
        if not all_proteins:
            continue  # empty G-vector, skip

        protein_set = frozenset(all_proteins)
        protein_set_str = ' | '.join(sorted(protein_set))

        rows.append({
            'Patient'       : patient,
            'ROI'           : roi,
            'ROI_Key'       : f'{patient}_{roi}',
            'G_Key'         : f'{patient}_{roi}_G{g_num}',
            'G_Num'         : g_num,
            'Protein_Set'   : protein_set_str,
            'Proteins_Plus' : ' | '.join(sorted(plus_proteins)),
            'Proteins_Minus': ' | '.join(sorted(minus_proteins)),
            'Protein_Count' : len(protein_set),
            '_frozenset'    : protein_set,
        })

    return rows


def main():
    print('=' * 70)
    print('EXPORT TAILS TABLE')
    print('=' * 70)

    # Scan all tail_proteins.xlsx files
    pattern = os.path.join(BASE_DIR, '*', '*', 'surprisal analysis', 'tail_proteins.xlsx')
    files = sorted(glob.glob(pattern))
    print(f'\n[1/2] Found {len(files)} tail_proteins.xlsx files')

    all_rows = []
    skipped  = 0

    for filepath in files:
        parts = os.path.normpath(filepath).split(os.sep)
        try:
            sa_idx = parts.index('surprisal analysis')
            roi = parts[sa_idx - 1]
            patient = parts[sa_idx - 2]
        except (ValueError, IndexError):
            skipped += 1
            continue

        try:
            df = pd.read_excel(filepath, engine='openpyxl')
        except Exception as e:
            print(f'  WARNING: could not read {filepath}: {e}')
            skipped += 1
            continue

        g_rows = extract_g_vectors(df, patient, roi)
        all_rows.extend(g_rows)

    print(f'      {len(all_rows):,} G-vectors extracted  ({skipped} files skipped)')

    # Build DataFrame
    df_all = pd.DataFrame(all_rows)

    # Reorder columns
    cols_order = [
        'Patient', 'ROI', 'ROI_Key', 'G_Key', 'G_Num',
        'Protein_Set', 'Proteins_Plus', 'Proteins_Minus', 'Protein_Count'
    ]
    df_out = df_all[cols_order].copy()

    # Export
    print(f'\n[2/2] Exporting TAILS table ...')
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')

    tails_ts_path = os.path.join(OUTPUT_DIR, f'All_Tails_{ts}.xlsx')
    tails_main_path = os.path.join(OUTPUT_DIR, 'All_Tails.xlsx')
    tails_csv_path = os.path.join(OUTPUT_DIR, 'All_Tails.csv')

    df_out.to_excel(tails_ts_path, index=False)
    df_out.to_excel(tails_main_path, index=False)
    df_out.to_csv(tails_csv_path, index=False, encoding='utf-8-sig')

    print(f'      TAILS  -> {tails_main_path}  ({len(df_out):,} rows)')
    print(f'      TAILS  -> {tails_ts_path}')
    print(f'      TAILS  -> {tails_csv_path}')

    # Summary
    print(f'\n{"=" * 70}')
    print('SUMMARY')
    print(f'{"=" * 70}')
    print(f'  ROIs processed    : {df_out["ROI_Key"].nunique()}')
    print(f'  Patients          : {df_out["Patient"].nunique()}')
    print(f'  Total G-vectors   : {len(df_out):,}')
    print(f'  Protein cnt range : {df_out["Protein_Count"].min()} - {df_out["Protein_Count"].max()}')
    print(f'  Avg proteins per G: {df_out["Protein_Count"].mean():.2f}')
    print(f'\n  Sample TAILS rows:')
    print(df_out.head(5).to_string(index=False))
    print(f'{"=" * 70}')


if __name__ == '__main__':
    main()
