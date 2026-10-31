import pandas as pd

# Load data
df = pd.read_csv('Docs/Exact_Sequences_Miner_20260628_184709.csv')

# Search for sequences containing CD20, CD3, and Ki67
mask = (df['Exact Sign Sequence'].str.contains('CD20', case=False) & 
        df['Exact Sign Sequence'].str.contains('CD3', case=False) & 
        df['Exact Sign Sequence'].str.contains('Ki67', case=False))

result = df[mask].sort_values(by=['Sequence Length', 'Unique Patients Count'], ascending=[False, False])

print(f'Found {len(result)} sequences with CD20, CD3, and Ki67')
print('\n' + '='*100)

if len(result) > 0:
    for idx, row in result.head(20).iterrows():
        print(f"\nLength: {row['Sequence Length']} proteins | Patients: {row['Unique Patients Count']} patients")
        print(f"Sequence: {row['Exact Sign Sequence']}")
        print('-'*100)
else:
    print("\nNo sequences found containing all three proteins (CD20, CD3, Ki67)")
    print("\nLet me check each protein individually:")
    
    cd20 = df[df['Exact Sign Sequence'].str.contains('CD20', case=False)]
    cd3 = df[df['Exact Sign Sequence'].str.contains('CD3', case=False)]
    ki67 = df[df['Exact Sign Sequence'].str.contains('Ki67', case=False)]
    
    print(f"\n- Sequences with CD20: {len(cd20)}")
    print(f"- Sequences with CD3: {len(cd3)}")
    print(f"- Sequences with Ki67: {len(ki67)}")
    
    # Check pairs
    cd20_cd3 = df[(df['Exact Sign Sequence'].str.contains('CD20', case=False) & 
                   df['Exact Sign Sequence'].str.contains('CD3', case=False))]
    cd20_ki67 = df[(df['Exact Sign Sequence'].str.contains('CD20', case=False) & 
                    df['Exact Sign Sequence'].str.contains('Ki67', case=False))]
    cd3_ki67 = df[(df['Exact Sign Sequence'].str.contains('CD3', case=False) & 
                   df['Exact Sign Sequence'].str.contains('Ki67', case=False))]
    
    print(f"\n- Sequences with CD20 + CD3: {len(cd20_cd3)}")
    print(f"- Sequences with CD20 + Ki67: {len(cd20_ki67)}")
    print(f"- Sequences with CD3 + Ki67: {len(cd3_ki67)}")
    
    if len(cd20_cd3) > 0:
        print("\nTop 5 sequences with CD20 + CD3:")
        for idx, row in cd20_cd3.sort_values(by=['Sequence Length', 'Unique Patients Count'], ascending=[False, False]).head(5).iterrows():
            print(f"  Length {row['Sequence Length']}: {row['Exact Sign Sequence'][:120]}")
