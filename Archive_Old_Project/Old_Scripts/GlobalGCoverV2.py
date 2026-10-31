"""
GlobalGCoverV2.py
-----------------
Two-part analysis using all 491K Exact Sign Sequences (ES):

PART 1 — Global Greedy Set Cover
  Universe : all G-values across all tail tables
  Pool     : all 491K ES sequences
  Priority : longest Sequence Length first;
             tie → most new G-values covered;
             tie → first encountered
  Constraint: each G-value assigned to at most one ES
  Result   : minimum ES set that covers maximum G-values globally

PART 2 — Per-Tail Greedy Set Cover
  For each tail ROI independently:
    Universe : G-values of that tail only
    Pool     : all ES sequences that overlap with that tail
    Same priority as Part 1
    Same no-repeat constraint
    Computes maximum achievable coverage % for that tail
  Only tails with achievable coverage >= TAIL_COVERAGE_THRESHOLD (80%) are
  included in the "All Possible Covers per Tail" output sheet.

Excel Output (3 sheets):
  1. Summary          — global stats + per-tail coverage list
  2. Selected ES      — global greedy result, ranked
  3. Per-Tail Covers  — per-tail greedy results (only tails >= 80%)
"""

import os
import re
import glob
import pandas as pd
from datetime import datetime
from collections import defaultdict

# ─────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────
TAILS_FOLDER          = 'All_Tail_Proteins_Tables'
SEQUENCES_FILE        = 'Docs/Exact_Sequences_Miner_20260630_090923.csv'
TAIL_COVERAGE_THRESHOLD = 80.0   # only show tails with >= this % achievable coverage


# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────

def parse_g_keys(s: str) -> frozenset:
    """'P1_ROI_03_G4 (+) | P9_ROI_51_G13 (-)' → frozenset of G-key strings."""
    if pd.isna(s):
        return frozenset()
    keys = set()
    for part in str(s).split(' | '):
        key = re.sub(r'\s*\([+-]\)\s*$', '', part).strip()
        if key:
            keys.add(key)
    return frozenset(keys)


def load_tail_g_values(tails_folder: str) -> dict:
    """
    Returns { roi_id → frozenset of G-key strings }
    e.g. 'P1_ROI_03' → frozenset({'P1_ROI_03_G1', 'P1_ROI_03_G2', ...})
    """
    files = sorted(
        glob.glob(os.path.join(tails_folder, '*.xlsx')) +
        glob.glob(os.path.join(tails_folder, '*.csv'))
    )
    tail_g_map = {}
    for f in files:
        # skip Excel lock files
        if os.path.basename(f).startswith('~$'):
            continue
        roi_id = (os.path.basename(f)
                  .replace('_tail_proteins.xlsx', '')
                  .replace('_tail_proteins.csv', ''))
        try:
            df = pd.read_excel(f) if f.endswith('.xlsx') else pd.read_csv(f)
        except Exception as e:
            print(f"  Warning: could not read {f}: {e}")
            continue

        g_nums = sorted(set(
            int(re.search(r'G(\d+)', c).group(1))
            for c in df.columns
            if re.match(r'G\d+[+-]$', c)
        ))
        g_keys = frozenset(f"{roi_id}_G{n}" for n in g_nums)
        if g_keys:
            tail_g_map[roi_id] = g_keys

    return tail_g_map


def greedy_cover(candidates_df: pd.DataFrame,
                 universe: frozenset,
                 label: str = '') -> tuple[list, set]:
    """
    Greedy set cover over `universe` using rows in `candidates_df`.
    Each row must have columns: '_g_set', 'Sequence Length', 'ES_label',
    'Unique Patients Count', 'Exact Sign Sequence', 'Participating ROIs'.

    Selection priority:
      1. Must cover at least 1 NEW G-value
      2. Longest 'Sequence Length' first
      3. Tie → most new G-values covered
      4. Tie → first encountered (stable sort order)

    Returns (selected_list, covered_set).
    selected_list: list of dicts with coverage metadata per selected ES.
    """
    covered: set = set()
    selected: list = []
    # Work on a list of (idx, g_set, seq_len) sorted by seq_len desc (stable)
    rows = sorted(
        [(idx, row['_g_set'], int(row['Sequence Length']))
         for idx, row in candidates_df.iterrows()],
        key=lambda x: -x[2]   # longest first
    )

    while rows:
        best_pos      = None
        best_new_count = 0
        best_new_set  : frozenset = frozenset()
        best_len      = -1

        for pos, (idx, g_set, seq_len) in enumerate(rows):
            new = g_set - covered
            n   = len(new)
            if n == 0:
                continue
            # Priority: longest length, then most new G-values
            if (seq_len > best_len) or \
               (seq_len == best_len and n > best_new_count):
                best_pos       = pos
                best_new_count = n
                best_new_set   = new
                best_len       = seq_len

        if best_pos is None:
            break   # nothing more to cover

        idx, g_set, seq_len = rows.pop(best_pos)
        covered.update(best_new_set)

        row = candidates_df.loc[idx]
        pct = len(covered) / len(universe) * 100 if universe else 0
        selected.append({
            'Rank'                   : len(selected) + 1,
            'ES_Label'               : row['ES_label'],
            'Sequence Length'        : seq_len,
            'Unique Patients Count'  : row['Unique Patients Count'],
            'New_G_Covered'          : best_new_count,
            'Total_G_Covered_So_Far' : len(covered),
            'Coverage_%'             : round(pct, 4),
            'Exact Sign Sequence'    : row['Exact Sign Sequence'],
            'G_Keys_Newly_Covered'   : ' | '.join(sorted(best_new_set)),
            'Participating ROIs'     : row['Participating ROIs'],
        })

        if label:
            print(f"    [{len(selected):>4}]  {row['ES_label']}"
                  f"  Len={seq_len:>2}  +{best_new_count} G"
                  f"  → {len(covered)}/{len(universe)} ({pct:.2f}%)")

        if covered >= universe:
            if label:
                print(f"  ✓ Full coverage with {len(selected)} ES!")
            break

    return selected, covered


# ─────────────────────────────────────────────
# PART 1 — GLOBAL GREEDY
# ─────────────────────────────────────────────

def run_global_cover(df: pd.DataFrame, universe: frozenset) -> tuple[list, set]:
    print("\n[PART 1] Global Greedy Set Cover")
    print(f"  Universe: {len(universe):,} G-values")
    print(f"  Pool    : {len(df):,} ES sequences\n")

    selected, covered = greedy_cover(df, universe, label='global')

    print(f"\n  → Selected {len(selected)} ES  |  "
          f"Covered {len(covered)}/{len(universe)} "
          f"({len(covered)/len(universe)*100:.2f}%)")
    return selected, covered


# ─────────────────────────────────────────────
# PART 2 — PER-TAIL GREEDY
# ─────────────────────────────────────────────

def run_per_tail_cover(df: pd.DataFrame,
                       tail_g_map: dict,
                       inverted_index: dict) -> list:
    """
    For each tail, independently run greedy cover.
    Returns list of per-tail result dicts (only those >= TAIL_COVERAGE_THRESHOLD).
    """
    print(f"\n[PART 2] Per-Tail Greedy Set Cover  (threshold ≥ {TAIL_COVERAGE_THRESHOLD}%)")
    tail_results = []
    total = len(tail_g_map)

    for i, (roi_id, g_keys) in enumerate(sorted(tail_g_map.items()), 1):
        print(f"  [{i:>3}/{total}] {roi_id}  ({len(g_keys)} G-values) ...", end='')

        # Candidate ES indices: only those covering ≥1 G-value of this tail
        cand_indices = set()
        for gk in g_keys:
            cand_indices.update(inverted_index.get(gk, []))

        if not cand_indices:
            print(f"  → 0.00%  (no ES covers this tail) — skipped")
            continue

        sub_df = df.loc[sorted(cand_indices)]
        selected, covered = greedy_cover(sub_df, g_keys, label='')

        n_covered = len(covered)
        pct       = n_covered / len(g_keys) * 100

        print(f"  → {pct:.2f}%  ({n_covered}/{len(g_keys)} G)  "
              f"using {len(selected)} ES")

        if pct < TAIL_COVERAGE_THRESHOLD:
            print(f"       (below {TAIL_COVERAGE_THRESHOLD}% threshold — excluded from output)")
            continue

        # Build per-ES detail rows for this tail
        es_detail_rows = []
        for s in selected:
            # Count how many of THIS tail's G-values the ES covers (in its new set)
            tail_g_in_new = [gk for gk in s['G_Keys_Newly_Covered'].split(' | ')
                             if gk.strip().startswith(roi_id)]
            es_detail_rows.append({
                'Tail_ROI'             : roi_id,
                'Tail_Total_G_Values'  : len(g_keys),
                'ES_Label'             : s['ES_Label'],
                'ES_Sequence_Length'   : s['Sequence Length'],
                'G_Values_Covered_Here': len(tail_g_in_new),
                'G_Keys_Covered_Here'  : ' | '.join(sorted(gk.strip()
                                          for gk in s['G_Keys_Newly_Covered'].split(' | ')
                                          if gk.strip().startswith(roi_id))),
                'Cumulative_G_Covered' : s['Total_G_Covered_So_Far'],
                'Cumulative_%'         : s['Coverage_%'],
                'Exact Sign Sequence'  : s['Exact Sign Sequence'],
                'Unique Patients Count': s['Unique Patients Count'],
            })

        tail_results.append({
            'roi_id'         : roi_id,
            'total_g'        : len(g_keys),
            'covered_g'      : n_covered,
            'pct'            : round(pct, 2),
            'n_es'           : len(selected),
            'es_labels'      : ', '.join(s['ES_Label'] for s in selected),
            'detail_rows'    : es_detail_rows,
        })

    return tail_results


# ─────────────────────────────────────────────
# BUILD SUMMARY DATAFRAME
# ─────────────────────────────────────────────

def build_summary(global_selected: list,
                  global_covered: set,
                  universe: frozenset,
                  tail_g_map: dict,
                  tail_results: list) -> pd.DataFrame:

    n_tails_above = len(tail_results)
    avg_cover     = (sum(t['pct'] for t in tail_results) / n_tails_above
                     if n_tails_above else 0)

    es_used_str   = ', '.join(s['ES_Label'] for s in global_selected)

    # Tails in global cover: tails that have at least one G covered globally
    g_to_es = {}
    for s in global_selected:
        for gk in s['G_Keys_Newly_Covered'].split(' | '):
            gk = gk.strip()
            if gk:
                g_to_es[gk] = s['ES_Label']

    tail_cover_strs = []
    n_tails_covered_global = 0
    for roi_id, g_keys in sorted(tail_g_map.items()):
        cov = sum(1 for gk in g_keys if gk in g_to_es)
        pct = round(cov / len(g_keys) * 100, 2) if g_keys else 0
        if cov > 0:
            n_tails_covered_global += 1
        tail_cover_strs.append(f"{roi_id}({pct}%)")

    rows = [
        ('Max coverage % (global)',           round(len(global_covered)/len(universe)*100, 4)),
        ('Number of tails covered >0% (global)', n_tails_covered_global),
        ('Total tails in dataset',             len(tail_g_map)),
        ('Number of ES sequences in global cover', len(global_selected)),
        ('Average cover % (per-tail greedy, ≥80% tails)', round(avg_cover, 4)),
        ('Number of tails with ≥80% achievable cover',    n_tails_above),
        ('Total G-values (universe)',          len(universe)),
        ('Total G-values covered (global)',    len(global_covered)),
        ('Total G-values uncovered (global)',  len(universe) - len(global_covered)),
        ('ES sequences used in global cover',  es_used_str),
        ('Tails included in global cover (with %)', ', '.join(tail_cover_strs)),
    ]
    return pd.DataFrame(rows, columns=['Metric', 'Value'])


# ─────────────────────────────────────────────
# EXPORT
# ─────────────────────────────────────────────

def export_results(summary_df, global_detail_df, per_tail_rows):
    os.makedirs('Docs', exist_ok=True)
    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx = f'Docs/GlobalGCoverV2_{ts}.xlsx'
    csv  = f'Docs/GlobalGCoverV2_{ts}.csv'

    # Flatten per-tail detail rows
    all_tail_detail = []
    for t in per_tail_rows:
        all_tail_detail.extend(t['detail_rows'])
    tail_detail_df = pd.DataFrame(all_tail_detail) if all_tail_detail else pd.DataFrame()

    # Per-tail summary (one row per tail)
    tail_summary_rows = [{
        'Tail_ROI'         : t['roi_id'],
        'Total_G_Values'   : t['total_g'],
        'Covered_G_Values' : t['covered_g'],
        'Coverage_%'       : t['pct'],
        'ES_Sequences_Used': t['n_es'],
        'ES_Labels'        : t['es_labels'],
    } for t in per_tail_rows]
    tail_summary_df = pd.DataFrame(tail_summary_rows).sort_values('Coverage_%', ascending=False)

    with pd.ExcelWriter(xlsx, engine='openpyxl') as writer:
        summary_df.to_excel(writer,      sheet_name='Summary',            index=False)
        global_detail_df.to_excel(writer, sheet_name='Selected ES (Global)', index=False)
        tail_summary_df.to_excel(writer, sheet_name='Per-Tail Summary',   index=False)
        if not tail_detail_df.empty:
            tail_detail_df.to_excel(writer, sheet_name='Per-Tail Covers', index=False)

    global_detail_df.to_csv(csv, index=False)

    print(f"\n  Excel → {xlsx}")
    print(f"  CSV   → {csv}")
    return xlsx


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    print("=" * 65)
    print("GLOBAL G-VALUE COVERAGE — V2  (longest-first greedy)")
    print("=" * 65)

    # ── 1. Load sequences ─────────────────────────────────────
    print(f"\n[1/5] Loading sequences from:\n      {SEQUENCES_FILE}")
    df = pd.read_csv(SEQUENCES_FILE)
    total_es = len(df)
    df['ES_label'] = [f"ES{i+1}" for i in range(total_es)]
    print(f"      {total_es:,} sequences loaded")

    # Parse G-key sets
    print("      Parsing Participating ROIs ...")
    df['_g_set']   = df['Participating ROIs'].apply(parse_g_keys)
    df['_g_count'] = df['_g_set'].apply(len)
    df = df[df['_g_count'] > 0].copy()
    df.reset_index(drop=True, inplace=True)
    print(f"      {len(df):,} ES with ≥1 G-value retained")

    # ── 2. Load tail tables ───────────────────────────────────
    print(f"\n[2/5] Loading tail tables from: {TAILS_FOLDER}")
    tail_g_map = load_tail_g_values(TAILS_FOLDER)
    print(f"      {len(tail_g_map)} tail tables loaded")

    universe: frozenset = frozenset().union(*tail_g_map.values())
    print(f"      Universe: {len(universe):,} unique G-values")

    # ── 3. Build inverted index  g_key → [es indices] ─────────
    print(f"\n[3/5] Building inverted index ...")
    inverted: dict = defaultdict(list)
    for idx, row in df.iterrows():
        for gk in row['_g_set']:
            if gk in universe:
                inverted[gk].append(idx)
    print(f"      Index built: {len(inverted):,} G-keys mapped")

    # Restrict df to ES that cover at least one tail G-value
    tail_es_indices = set()
    for idxs in inverted.values():
        tail_es_indices.update(idxs)
    df_tail = df.loc[sorted(tail_es_indices)].copy()
    print(f"      ES covering ≥1 tail G-value: {len(df_tail):,}")

    # ── 4. Part 1 — global greedy ─────────────────────────────
    print(f"\n[4/5] Running global greedy ...")
    global_selected, global_covered = run_global_cover(df_tail, universe)
    global_detail_df = pd.DataFrame(global_selected)

    # ── 5. Part 2 — per-tail greedy ───────────────────────────
    print(f"\n[5/5] Running per-tail greedy ...")
    tail_results = run_per_tail_cover(df, tail_g_map, inverted)

    # ── Build summary & export ─────────────────────────────────
    summary_df = build_summary(global_selected, global_covered,
                               universe, tail_g_map, tail_results)

    print(f"\nExporting results ...")
    xlsx = export_results(summary_df, global_detail_df, tail_results)

    # ── Print final summary ────────────────────────────────────
    print("\n" + "=" * 65)
    print("FINAL SUMMARY")
    print("=" * 65)
    for _, row in summary_df.head(10).iterrows():
        print(f"  {row['Metric']:<52}: {row['Value']}")
    print("=" * 65)
    print(f"\nOutput saved to: {xlsx}")


if __name__ == '__main__':
    main()
