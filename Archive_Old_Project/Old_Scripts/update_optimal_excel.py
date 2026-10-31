import glob
import os
import re
import pandas as pd
import numpy as np
from collections import Counter
import openpyxl

# 1. Load tail proteins and distinct G numbers per ROI
files = sorted(glob.glob('All_Tail_Proteins_Tables/*.xlsx') + glob.glob('All_Tail_Proteins_Tables/*.csv'))
tail_g_map = {}

for f in files:
    if os.path.basename(f).startswith('~$'):
        continue
    roi_id = os.path.basename(f).replace('_tail_proteins.xlsx', '').replace('_tail_proteins.csv', '')
    df = pd.read_excel(f) if f.endswith('.xlsx') else pd.read_csv(f)
    g_cols = [c for c in df.columns if re.match(r'^G\d+[+-]$', c)]
    g_nums = sorted(set(int(re.search(r'G(\d+)', c).group(1)) for c in g_cols))
    tail_g_map[roi_id] = [f"{roi_id}_G{n}" for n in g_nums]

print(f"Loaded {len(tail_g_map)} ROIs with {sum(len(v) for v in tail_g_map.values())} total G-vectors.")

# 2. Read Excel
excel_path = 'Docs/Optimal_ES_Combination_20260803_195423_2_updated.xlsx'
xl = pd.ExcelFile(excel_path)
sheets_dict = {s: xl.parse(s) for s in xl.sheet_names}

df_opt = sheets_dict['Optimal_Combination']

# 3. Compute Tails_Detail and Avg_Tail_Coverage_% (average of all percentages in Tails_Detail)
new_tails_detail = []
new_avg_tail_cov = []

for _, row in df_opt.iterrows():
    parts = [re.sub(r'\s*\([+-]\)\s*$', '', p).strip() for p in str(row['Participating_ROIs']).split(' | ') if p.strip()]
    g_counts_per_roi = Counter([re.sub(r'_G\d+$', '', p) for p in parts])
    
    tail_pcts = []
    tail_items = []
    for roi_id in sorted(tail_g_map.keys()):
        cnt = g_counts_per_roi.get(roi_id, 0)
        tot = len(tail_g_map[roi_id])
        if cnt > 0:
            pct = cnt / tot * 100
            tail_pcts.append(pct)
            tail_items.append((roi_id, cnt, tot, pct))
            
    # Sort tails by count descending
    tail_items.sort(key=lambda x: -x[1])
    td_str = '; '.join([f"{r}({c}G/{t}G {p:.0f}%)" for r, c, t, p in tail_items])
    avg_tail_cov = round(np.mean(tail_pcts), 2) if tail_pcts else 0.0
    
    new_tails_detail.append(td_str)
    new_avg_tail_cov.append(avg_tail_cov)

# 4. Update dataframe: remove Avg_Protein_Coverage_Per_G_% and old Avg_G_Coverage_%_Per_Section
# Keep clean column structure
cols_to_keep = ['Rank', 'ES_Label', 'Sequence_Length', 'Unique_Patients_Count', 'Total_G_Values', 'Exact_Sign_Sequence', 'Participating_ROIs']
clean_df = df_opt[cols_to_keep].copy()
clean_df['Tails_Detail'] = new_tails_detail
clean_df['Avg_Tail_Coverage_%'] = new_avg_tail_cov

print("\nUpdated columns in Optimal_Combination:")
print(clean_df.columns.tolist())

print("\nSample top 5 rows:")
print(clean_df[['Rank', 'ES_Label', 'Sequence_Length', 'Total_G_Values', 'Avg_Tail_Coverage_%']].head(5).to_string())

# 5. Save to both files
for target in ['Docs/Optimal_ES_Combination_20260803_195423_2_updated.xlsx', 'Docs/Optimal_ES_Combination_20260803_195423_2_updated_new.xlsx']:
    with pd.ExcelWriter(target, engine='openpyxl') as writer:
        for sheet_name, df_sheet in sheets_dict.items():
            if sheet_name == 'Optimal_Combination':
                clean_df.to_excel(writer, sheet_name=sheet_name, index=False)
            else:
                df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
    print(f"Successfully saved to {target}")
