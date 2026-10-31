"""
generate_es_from_tails.py
--------------------------
Generates Exact_Sequences_Miner_<timestamp>.csv from All_Tail_Proteins_Tables.

Logic:
  - Each G-vector in each tail table → one ES sequence candidate.
  - ES sequence = all proteins in that G-vector, written as:
        ProteinName(sign) | ProteinName(sign) | ...
    sorted alphabetically, using the raw table sign (G{n}+ → '+', G{n}- → '-').
  - Identical sequences that appear across multiple ROIs/G-vectors are grouped
    into one row.

Output columns (same schema as original Exact_Sequences_Miner):
  Sequence Length        — number of proteins in the ES
  Unique Patients Count  — number of distinct P-numbers
  Exact Sign Sequence    — "CD163(-) | CD21(-) | ..."
  Participating ROIs     — "P10_ROI_48_G12 (-) | P7_ROI_32_G6 (+) | ..."

  The orientation sign after each ROI_G entry is always (+) because we read
  the proteins directly from the table (no flip needed — this IS the canonical
  form of the sequence).
"""

import os
import re
import glob
import pandas as pd
from datetime import datetime
from collections import defaultdict

TAILS_FOLDER = 'All_Tail_Proteins_Tables'
OUTPUT_DIR   = 'Docs'


# ─────────────────────────────────────────────
# EXTRACT
# ─────────────────────────────────────────────

def extract_all_g_vectors(tails_folder: str) -> list:
    """
    Returns a list of dicts:
      { 'gkey': 'P10_ROI_48_G12',
        'roi_id': 'P10_ROI_48',
        'patient': 'P10',
        'es_seq': 'CD163(-) | CD21(-) | ...' }
    """
    files = sorted(
        glob.glob(os.path.join(tails_folder, '*.xlsx')) +
        glob.glob(os.path.join(tails_folder, '*.csv'))
    )
    vectors = []

    for f in files:
        if os.path.basename(f).startswith('~$'):
            continue
        roi_id = (os.path.basename(f)
                  .replace('_tail_proteins.xlsx', '')
                  .replace('_tail_proteins.csv', ''))

        # Extract patient number: first P-token e.g. "P10" from "P10_ROI_48"
        m_patient = re.match(r'^(P\d+)', roi_id)
        patient   = m_patient.group(1) if m_patient else roi_id

        try:
            df = (pd.read_excel(f, engine='openpyxl')
                  if f.endswith('.xlsx') else pd.read_csv(f))
        except Exception as e:
            print(f"  Warning: could not read {f}: {e}")
            continue

        # Detect G-numbers from columns like G1+, G1-, G12+, G12-
        g_nums = sorted(set(
            int(re.search(r'G(\d+)', c).group(1))
            for c in df.columns
            if re.match(r'G\d+[+-]$', c)
        ))

        for n in g_nums:
            proteins = {}          # name → sign
            for sign in ('+', '-'):
                col = f'G{n}{sign}'
                if col in df.columns:
                    for pname in df[col].dropna():
                        pname = str(pname).strip()
                        if pname:
                            proteins[pname] = sign

            if not proteins:
                continue

            # Build canonical sequence: alphabetical by protein name
            es_seq = ' | '.join(
                f"{p}({s})"
                for p, s in sorted(proteins.items(), key=lambda x: x[0])
            )

            vectors.append({
                'gkey'    : f"{roi_id}_G{n}",
                'roi_id'  : roi_id,
                'patient' : patient,
                'es_seq'  : es_seq,
            })

    return vectors


# ─────────────────────────────────────────────
# DEDUPLICATE & BUILD TABLE
# ─────────────────────────────────────────────

def build_es_table(vectors: list) -> pd.DataFrame:
    """
    Group identical ES sequences.
    Each group → one row with all ROI_G appearances.
    """
    # seq → list of (gkey, roi_id, patient)
    seq_map: dict = defaultdict(list)
    for v in vectors:
        seq_map[v['es_seq']].append((v['gkey'], v['roi_id'], v['patient']))

    records = []
    for es_seq, appearances in seq_map.items():
        seq_len        = len(es_seq.split(' | '))
        patients       = sorted(set(p for _, _, p in appearances))
        unique_patients = len(patients)

        # Sort appearances: by roi_id then G-number
        appearances_sorted = sorted(
            appearances,
            key=lambda x: (x[1], int(re.search(r'_G(\d+)$', x[0]).group(1)))
        )
        participating = ' | '.join(f"{gkey} (+)" for gkey, _, _ in appearances_sorted)

        records.append({
            'Sequence Length'      : seq_len,
            'Unique Patients Count': unique_patients,
            'Exact Sign Sequence'  : es_seq,
            'Participating ROIs'   : participating,
        })

    df = pd.DataFrame(records)
    df = df.sort_values(
        by=['Sequence Length', 'Unique Patients Count'],
        ascending=[False, False]
    ).reset_index(drop=True)
    return df


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    print("=" * 65)
    print("GENERATE EXACT SEQUENCES MINER  (from tail tables)")
    print("=" * 65)

    print(f"\n[1/3] Scanning tail tables in: {TAILS_FOLDER}")
    vectors = extract_all_g_vectors(TAILS_FOLDER)
    print(f"      {len(vectors):,} G-vectors extracted")

    print(f"\n[2/3] Deduplicating and building table ...")
    df = build_es_table(vectors)
    print(f"      {len(df):,} unique ES sequences")
    print(f"      Sequence length range: "
          f"{df['Sequence Length'].min()} – {df['Sequence Length'].max()}")
    print(f"      Unique patients range: "
          f"{df['Unique Patients Count'].min()} – "
          f"{df['Unique Patients Count'].max()}")

    print(f"\n[3/3] Exporting ...")
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(OUTPUT_DIR, f'Exact_Sequences_Miner_{ts}.csv')
    df.to_csv(out, index=False, encoding='utf-8-sig')
    print(f"      → {out}")

    print(f"\n{'='*65}")
    print("SAMPLE (first 5 rows):")
    print(f"{'='*65}")
    for _, row in df.head(5).iterrows():
        print(f"\n  Len={row['Sequence Length']}  "
              f"Patients={row['Unique Patients Count']}")
        print(f"  Seq : {row['Exact Sign Sequence'][:80]}")
        print(f"  ROIs: {str(row['Participating ROIs'])[:100]}")
    print(f"\n{'='*65}")
    print(f"Output: {out}")
    print(f"{'='*65}")

    return out


if __name__ == '__main__':
    main()
