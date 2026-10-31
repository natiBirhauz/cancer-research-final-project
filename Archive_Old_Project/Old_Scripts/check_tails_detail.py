import pandas as pd
xl = pd.ExcelFile('Docs/CoreESAnalysis_20260713_184847.xlsx')
df = xl.parse('PartA All ES Ranked')
print('Columns:', df.columns.tolist())
print()
for _, row in df.head(3).iterrows():
    print(f"ES:           {row['ES_Label']}")
    print(f"Sequence:     {row['Exact Sign Sequence']}")
    print(f"G_Values:     {row['G_Values_Covered']}")
    print(f"Tails:        {row['Tails_Covered']}")
    print(f"Tails_Detail: {row['Tails_Detail']}")
    print()
