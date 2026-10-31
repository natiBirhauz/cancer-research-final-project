"""
03_mine_all_es_combinations.py
------------------------------
Step 3 of the Cancer Research TME Pipeline:
Loads all 2,778 G-vectors from All_Tail_Proteins_Tables/ and mines ALL possible
Exact Sign Sequence (ES) combinations from length 3 to 8 (120,791 unique combinations).
Computes exact occurrences, recurring 2+ G-vector signatures, and exports:
- Docs/All_Possible_ES_Combinations_From_Length_3.csv
- Docs/All_Possible_ES_Combinations_From_Length_3.xlsx
"""

import os
import glob
import re
import time
from collections import Counter, defaultdict
from itertools import combinations
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TAILS_DIR = 'All_Tail_Proteins_Tables'
OUTPUT_CSV = os.path.join('Docs', 'All_Possible_ES_Combinations_From_Length_3.csv')
OUTPUT_XLSX = os.path.join('Docs', 'All_Possible_ES_Combinations_From_Length_3.xlsx')

def canon_tuple(d):
    s = sorted(d.items(), key=lambda x: x[0])
    if s[0][1] == -1:
        return tuple((p, -v) for p, v in s)
    return tuple(s)

def tuple_to_str(t):
    return " | ".join(f"{p}({'+' if s==1 else '-'})" for p, s in t)

def main():
    t0 = time.time()
    print("====================================================================")
    print(" [Step 3/5] Mining All ES Combinations from Length 3 to 8")
    print("====================================================================")

    files = sorted(glob.glob(os.path.join(TAILS_DIR, '*.xlsx')))
    print(f"Loading G-vectors with exact signs from {len(files)} ROI files...")

    g_vectors = []
    all_prots = set()

    for f in files:
        roi_id = os.path.basename(f).replace('_tail_proteins.xlsx', '')
        df = pd.read_excel(f)
        g_nums = sorted(set(int(re.search(r'G(\d+)', c).group(1)) for c in df.columns if re.match(r'G\d+[+-]$', c)))
        for n in g_nums:
            signs = {}
            for s, col in [('+', f'G{n}+'), ('-', f'G{n}-')]:
                if col in df.columns:
                    for p in df[col].dropna():
                        p = str(p).strip()
                        if p:
                            signs[p] = 1 if s == '+' else -1
                            all_prots.add(p)
            if len(signs) >= 3:
                g_key = f"{roi_id}_G{n}"
                g_vectors.append((g_key, signs))

    print(f"Loaded {len(g_vectors)} G-vectors with >= 3 proteins across {len(all_prots)} unique proteins.")
    print("Mining all canonical combinations of lengths 3 to 8...")

    combo_g_map = defaultdict(list)
    for g_key, signs in g_vectors:
        n = len(signs)
        items = list(signs.items())
        for k in range(3, min(n + 1, 9)):
            for combo in combinations(items, k):
                ct = canon_tuple(dict(combo))
                combo_g_map[ct].append(g_key)

    total_unique = len(combo_g_map)
    print(f"\nMined {total_unique:,} unique canonical ES combinations!")

    # Summary by length
    by_len = defaultdict(list)
    for ct, glist in combo_g_map.items():
        by_len[len(ct)].append((ct, glist))

    summary_records = []
    for l in sorted(by_len.keys()):
        combos = by_len[l]
        tot = len(combos)
        rec2 = sum(1 for ct, glist in combos if len(glist) >= 2)
        max_freq = max(len(glist) for ct, glist in combos)
        top_combo = max(combos, key=lambda x: len(x[1]))
        top_str = f"{tuple_to_str(top_combo[0])} (in {len(top_combo[1])} Gs)"
        summary_records.append({
            'Length': l,
            'Total_Unique_Combinations': tot,
            'Combinations_in_2+_Gs': rec2,
            'Max_G_Frequency': max_freq,
            'Top_Combination': top_str
        })
        print(f"  Length {l}: {tot:>6,} combinations | 2+ Gs: {rec2:>5,} | Max Freq: {max_freq:>2} Gs")

    sum_df = pd.DataFrame(summary_records)

    # Export CSV
    print(f"\nExporting full CSV to '{OUTPUT_CSV}'...")
    csv_rows = []
    for ct, glist in combo_g_map.items():
        csv_rows.append({
            'Length': len(ct),
            'Frequency_Count': len(glist),
            'In_2+_Gs': len(glist) >= 2,
            'Exact_Sign_Sequence': tuple_to_str(ct),
            'Covered_G_Vectors': "; ".join(glist)
        })
    csv_df = pd.DataFrame(csv_rows)
    csv_df.sort_values(by=['Length', 'Frequency_Count'], ascending=[True, False], inplace=True)
    csv_df.to_csv(OUTPUT_CSV, index=False, encoding='utf-8-sig')

    # Export Excel with styled sheets
    print(f"Exporting styled Excel workbook to '{OUTPUT_XLSX}'...")
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    f_title = Font(name='Segoe UI', size=14, bold=True, color='1F497D')
    f_header = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    f_bold = Font(name='Segoe UI', size=10, bold=True)
    f_regular = Font(name='Segoe UI', size=10)
    fill_header = PatternFill(start_color='1F497D', end_color='1F497D', fill_type='solid')
    fill_zebra = PatternFill(start_color='F5F7FA', end_color='F5F7FA', fill_type='solid')
    border_thin = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    # Sheet 1: Summary_By_Length
    ws1 = wb.create_sheet('Summary_By_Length')
    ws1.views.sheetView[0].showGridLines = True
    ws1.cell(row=1, column=1, value="ALL CANDIDATE ES COMBINATIONS (LENGTHS 3 TO 8) SUMMARY").font = f_title
    headers1 = ["Length", "Total_Unique_Combinations", "Combinations_in_2+_Gs", "Max_G_Frequency", "Top_Combination"]
    for c, h in enumerate(headers1, 1):
        cell = ws1.cell(row=3, column=c, value=h)
        cell.font = f_header; cell.fill = fill_header; cell.alignment = Alignment(horizontal='center')
    for r_idx, row in enumerate(summary_records, 4):
        vals = [row['Length'], row['Total_Unique_Combinations'], row['Combinations_in_2+_Gs'], row['Max_G_Frequency'], row['Top_Combination']]
        for c, v in enumerate(vals, 1):
            cell = ws1.cell(row=r_idx, column=c, value=v)
            cell.font = f_regular; cell.border = border_thin
            if c <= 4: cell.alignment = Alignment(horizontal='center')
            if r_idx % 2 == 1: cell.fill = fill_zebra
    ws1.column_dimensions['A'].width = 12
    ws1.column_dimensions['B'].width = 28
    ws1.column_dimensions['C'].width = 25
    ws1.column_dimensions['D'].width = 18
    ws1.column_dimensions['E'].width = 80

    # Sheet 2: Frequent_Combinations_2plus_Gs
    ws2 = wb.create_sheet('Frequent_Combinations_2plus_Gs')
    ws2.views.sheetView[0].showGridLines = True
    headers2 = ["Length", "G_Count", "Exact_Sign_Sequence", "Covered_G_Vectors"]
    for c, h in enumerate(headers2, 1):
        cell = ws2.cell(row=1, column=c, value=h)
        cell.font = f_header; cell.fill = fill_header; cell.alignment = Alignment(horizontal='center')
    freq_combos = [r for r in csv_rows if r['In_2+_Gs']]
    for r_idx, r in enumerate(freq_combos, 2):
        vals = [r['Length'], r['Frequency_Count'], r['Exact_Sign_Sequence'], r['Covered_G_Vectors']]
        for c, v in enumerate(vals, 1):
            cell = ws2.cell(row=r_idx, column=c, value=v)
            cell.font = f_regular; cell.border = border_thin
            if c <= 2: cell.alignment = Alignment(horizontal='center')
            if r_idx % 2 == 1: cell.fill = fill_zebra
    ws2.column_dimensions['A'].width = 10
    ws2.column_dimensions['B'].width = 12
    ws2.column_dimensions['C'].width = 75
    ws2.column_dimensions['D'].width = 60

    # Sheet 3: Top_500_Per_Length
    ws3 = wb.create_sheet('Top_500_Per_Length')
    ws3.views.sheetView[0].showGridLines = True
    for c, h in enumerate(headers2, 1):
        cell = ws3.cell(row=1, column=c, value=h)
        cell.font = f_header; cell.fill = fill_header; cell.alignment = Alignment(horizontal='center')
    top500 = []
    for l in sorted(by_len.keys()):
        l_combos = [r for r in csv_rows if r['Length'] == l]
        top500.extend(l_combos[:500])
    for r_idx, r in enumerate(top500, 2):
        vals = [r['Length'], r['Frequency_Count'], r['Exact_Sign_Sequence'], r['Covered_G_Vectors']]
        for c, v in enumerate(vals, 1):
            cell = ws3.cell(row=r_idx, column=c, value=v)
            cell.font = f_regular; cell.border = border_thin
            if c <= 2: cell.alignment = Alignment(horizontal='center')
            if r_idx % 2 == 1: cell.fill = fill_zebra
    ws3.column_dimensions['A'].width = 10
    ws3.column_dimensions['B'].width = 12
    ws3.column_dimensions['C'].width = 75
    ws3.column_dimensions['D'].width = 60

    wb.save(OUTPUT_XLSX)
    print(f"\nStep 3 Complete: Mined all combinations and saved outputs in {time.time()-t0:.2f}s.\n")
    return 0

if __name__ == '__main__':
    main()
