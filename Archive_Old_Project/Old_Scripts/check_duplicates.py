import pandas as pd

df = pd.read_excel('Docs/Non_Redundant_ES_COM_Report_20260726_230802.xlsx', sheet_name='ES_COM_1_Details')
row = df[df['ROI'] == 'P10_ROI_09']

print('P10_ROI_09 sample columns:')
print(row[['ROI', 'ES463181', 'ES463216']].to_string())

print('\n--- Check for G16(-) duplicates ---')
for col in df.columns[1:]:
    val = row[col].values[0]
    if pd.notna(val) and 'G16(-)' in str(val):
        print(f'{col}: {val}')

# Check for cells with multiple G values
print('\n--- Cells with multiple G values ---')
for idx, row_data in df.head(10).iterrows():
    roi = row_data['ROI']
    for col in df.columns[1:]:
        val = row_data[col]
        if pd.notna(val):
            val_str = str(val)
            # Count commas or semicolons or "G" occurrences
            g_count = val_str.count('G')
            if g_count > 1:
                print(f'{roi} | {col}: {val_str}')
                break
