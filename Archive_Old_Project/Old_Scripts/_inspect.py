import pandas as pd, re
df = pd.read_excel(r"All_Tail_Proteins_Tables/P10_ROI_48_tail_proteins.xlsx")
print("SHAPE:", df.shape)
print("COLUMNS:", list(df.columns))
print()
for col in df.columns[:6]:
    vals=df[col].dropna().tolist()
    print(f"{col}: {vals[:5]}")
print()
g_cols=[c for c in df.columns if re.match(r"G\\d+[+-]$",c)]
print("G-cols:",g_cols)
print()
g12p=df["G12+"].dropna().tolist() if "G12+" in df.columns else []
g12m=df["G12-"].dropna().tolist() if "G12-" in df.columns else []
print("G12+ proteins:",g12p)
print("G12- proteins:",g12m)
