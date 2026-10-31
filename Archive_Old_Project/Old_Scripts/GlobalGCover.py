"""
GlobalGCover.py
---------------
Finds the MINIMUM number of Exact Sign Sequences (ES1..ES491174) that cover
the MAXIMUM number of G-components across ALL ROIs.

Rules:
  - Universe  = every G-component that appears in Participating ROIs
                e.g.  P1_ROI_03_G4,  P9_ROI_51_G13, ...
  - Each ES covers the G-components listed in its 'Participating ROIs' column.
  - The SAME G-component cannot be assigned to two different ES.
  - Goal: greedy set-cover — at every step pick the ES that covers the most
          still-uncovered G-components, until no new G-component can be added.

Output (saved to Docs/):
  GlobalGCover_<timestamp>.xlsx   — two sheets: Summary + Detail
  GlobalGCover_<timestamp>.csv    — Detail rows only
"""

import os
import re
import pandas as pd
from datetime import datetime


# ──────────────────────────────────────────────
# helpers
# ──────────────────────────────────────────────

def parse_g_keys(participating_rois: str) -> frozenset:
    """
    'P1_ROI_03_G4 (+) | P9_ROI_51_G13 (-) | ...'
    → frozenset({'P1_ROI_03_G4', 'P9_ROI_51_G13', ...})
    """
    if pd.isna(participating_rois):
        return frozenset()
    keys = set()
    for part in str(participating_rois).split(' | '):
        key = re.sub(r'\s*\([+-]\)\s*$', '', part).strip()
        if key:
            keys.add(key)
    return frozenset(keys)


# ──────────────────────────────────────────────
# main
# ──────────────────────────────────────────────

def main():
    sequences_file = 'Docs/Exact_Sequences_Miner_20260630_090923.csv'

    print("=" * 65)
    print("GLOBAL G-COMPONENT SET COVER")
    print("=" * 65)

    # ── 1. Load ──────────────────────────────────────────────────────
    print(f"\n[1/4] Loading sequences from:\n      {sequences_file}")
    df = pd.read_csv(sequences_file)
    total_es = len(df)
    print(f"      {total_es:,} sequences (ES1 … ES{total_es})")

    # ── 2. Parse G-key sets ──────────────────────────────────────────
    print("\n[2/4] Parsing G-components from Participating ROIs...")
    df['_g_set'] = df['Participating ROIs'].apply(parse_g_keys)

    # Universe = all unique G-components
    universe: set = set()
    for gs in df['_g_set']:
        universe.update(gs)
    total_g = len(universe)
    print(f"      Universe: {total_g:,} unique G-components across all ROIs")

    # Remove sequences that cover ZERO G-components
    mask = df['_g_set'].apply(len) > 0
    df = df[mask].copy()
    print(f"      Sequences with ≥1 G-component: {len(df):,}")

    # Sort by coverage size descending — greedy will converge faster
    df['_g_count'] = df['_g_set'].apply(len)
    df = df.sort_values('_g_count', ascending=False).reset_index(drop=True)
    # ES label = original 1-based row number (before sort, stored separately)
    # We track original index from the CSV order
    df['ES_label'] = [f"ES{i+1}" for i in range(total_es)]   # reassign after reload
    # Actually reload original order for correct ES labels
    df_orig = pd.read_csv(sequences_file)
    df_orig['ES_label'] = [f"ES{i+1}" for i in range(len(df_orig))]
    df_orig['_g_set']   = df_orig['Participating ROIs'].apply(parse_g_keys)
    df_orig['_g_count'] = df_orig['_g_set'].apply(len)
    df_orig = df_orig[df_orig['_g_count'] > 0].copy()
    df_orig = df_orig.sort_values('_g_count', ascending=False).reset_index(drop=True)

    # ── 3. Greedy Set Cover ──────────────────────────────────────────
    print("\n[3/4] Running Greedy Set Cover …")
    print(f"      (stops when no more G-components can be covered)\n")

    covered:  set = set()
    selected: list = []          # list of result dicts
    remaining = list(df_orig.index)

    iteration = 0
    while remaining:
        iteration += 1

        # Find the ES that covers the most NEW G-components
        best_idx      = None
        best_new_count = 0
        best_new_set  : set = set()

        for idx in remaining:
            new = df_orig.loc[idx, '_g_set'] - covered
            n   = len(new)
            if n > best_new_count:
                best_new_count = n
                best_idx       = idx
                best_new_set   = new
                if n > 200:          # early-exit hint for very large covers
                    break

        if best_new_count == 0:
            print(f"  → Stopped at iteration {iteration}: "
                  f"no more G-components can be covered.")
            break

        # Commit
        row = df_orig.loc[best_idx]
        covered.update(best_new_set)
        remaining.remove(best_idx)

        pct = len(covered) / total_g * 100
        selected.append({
            'Rank':                   len(selected) + 1,
            'ES_Label':               row['ES_label'],
            'Sequence Length':        row['Sequence Length'],
            'Unique Patients Count':  row['Unique Patients Count'],
            'New_G_Covered':          best_new_count,
            'Total_G_Covered_So_Far': len(covered),
            'Coverage_%':             round(pct, 4),
            'Exact Sign Sequence':    row['Exact Sign Sequence'],
            'G_Keys_Newly_Covered':   ' | '.join(sorted(best_new_set)),
            'Participating ROIs':     row['Participating ROIs'],
        })

        print(f"  [{len(selected):>4}]  {row['ES_label']}  "
              f"Len={row['Sequence Length']}  "
              f"+{best_new_count} G  →  "
              f"{len(covered)}/{total_g}  ({pct:.2f}%)")

        if len(covered) == total_g:
            print(f"\n  ✓ Full coverage achieved with {len(selected)} sequences!")
            break

    # ── 4. Export ────────────────────────────────────────────────────
    print("\n[4/4] Exporting results …")
    os.makedirs('Docs', exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    detail_df = pd.DataFrame(selected)

    # Summary card
    summary = {
        'Total ES in file':              total_es,
        'ES Selected':                   len(selected),
        'Total G-components (universe)': total_g,
        'G-components covered':          len(covered),
        'G-components uncovered':        total_g - len(covered),
        'Overall coverage %':            round(len(covered) / total_g * 100, 4),
    }
    summary_df = pd.DataFrame([summary]).T.reset_index()
    summary_df.columns = ['Metric', 'Value']

    # Uncovered G-components
    uncovered_df = pd.DataFrame({'Uncovered_G_Component': sorted(universe - covered)})

    xlsx = f'Docs/GlobalGCover_{ts}.xlsx'
    csv  = f'Docs/GlobalGCover_{ts}.csv'

    with pd.ExcelWriter(xlsx, engine='openpyxl') as writer:
        summary_df.to_excel(writer,    sheet_name='Summary',    index=False)
        detail_df.to_excel(writer,     sheet_name='Selected ES', index=False)
        uncovered_df.to_excel(writer,  sheet_name='Uncovered',   index=False)

    detail_df.to_csv(csv, index=False)

    print(f"\n  Excel → {xlsx}")
    print(f"  CSV   → {csv}")

    print("\n" + "=" * 65)
    print("FINAL SUMMARY")
    print("=" * 65)
    for k, v in summary.items():
        print(f"  {k:<40}: {v}")
    print("=" * 65)


if __name__ == '__main__':
    main()
