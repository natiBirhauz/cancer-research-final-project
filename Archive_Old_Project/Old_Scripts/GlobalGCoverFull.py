"""
GlobalGCoverFull.py
-------------------
Finds the MINIMUM set of Exact Sign Sequences (ES) that covers the MAXIMUM
percentage of G-values across ALL tail tables.

Definitions:
  - G-value   = one G-component of one tail table, e.g. P1_ROI_03_G4
  - Universe  = all G-values found across every tail table file
  - ES covers a G-value if that G-value appears in the ES's Participating ROIs
  - Constraint: the same G-value cannot be assigned to two different ES
  - Goal: greedy set-cover — pick ES with the most still-uncovered G-values
          until no further G-value can be added

Output columns added to the summary:
  max % cover, number of tails cover, number of ES sequences in cover,
  average cover for tail tables, ES sequences used in cover,
  tails included in cover (with % of cover)
"""

import os, re, glob
import pandas as pd
from datetime import datetime
from collections import defaultdict


# ─────────────────────────────────────────────────────────────
# 1.  Read all tail tables → build universe of G-values
# ─────────────────────────────────────────────────────────────

def load_tail_g_values(tails_folder: str) -> dict:
    """
    Returns:
        tail_g_map : { roi_id → set of full G-keys }
                     e.g. 'P1_ROI_03' → {'P1_ROI_03_G1','P1_ROI_03_G2',...}
    """
    files = sorted(
        glob.glob(os.path.join(tails_folder, '*.xlsx')) +
        glob.glob(os.path.join(tails_folder, '*.csv'))
    )
    tail_g_map = {}
    for f in files:
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
        g_keys = {f"{roi_id}_G{n}" for n in g_nums}
        if g_keys:
            tail_g_map[roi_id] = g_keys

    return tail_g_map


# ─────────────────────────────────────────────────────────────
# 2.  Parse ES participating ROIs  (once)
# ─────────────────────────────────────────────────────────────

def parse_g_keys(s: str) -> frozenset:
    if pd.isna(s):
        return frozenset()
    keys = set()
    for part in str(s).split(' | '):
        key = re.sub(r'\s*\([+-]\)\s*$', '', part).strip()
        if key:
            keys.add(key)
    return frozenset(keys)


# ─────────────────────────────────────────────────────────────
# 3.  Greedy global set-cover
# ─────────────────────────────────────────────────────────────

def greedy_cover(df: pd.DataFrame, universe: set) -> list:
    """
    Returns list of dicts, one per selected ES, in selection order.
    Each dict contains all original ES columns + coverage metadata.
    """
    covered   : set  = set()
    selected  : list = []
    remaining        = list(df.index)

    while remaining:
        best_idx  = None
        best_new  = 0
        best_set  : set = set()

        for idx in remaining:
            new = df.at[idx, '_g_set'] - covered
            n   = len(new)
            if n > best_new:
                best_new  = n
                best_idx  = idx
                best_set  = new
                if n > 200:
                    break          # early-exit: clearly dominant

        if best_new == 0:
            break                  # nothing more to cover

        covered.update(best_set)
        remaining.remove(best_idx)

        row = df.loc[best_idx]
        pct = len(covered) / len(universe) * 100
        selected.append({
            'Rank'                    : len(selected) + 1,
            'ES_Label'                : row['ES_label'],
            'Sequence Length'         : row['Sequence Length'],
            'Unique Patients Count'   : row['Unique Patients Count'],
            'New_G_Covered'           : best_new,
            'Total_G_Covered_So_Far'  : len(covered),
            'Coverage_%'              : round(pct, 4),
            'Exact Sign Sequence'     : row['Exact Sign Sequence'],
            'G_Keys_Newly_Covered'    : ' | '.join(sorted(best_set)),
            'Participating ROIs'      : row['Participating ROIs'],
        })

        print(f"  [{len(selected):>4}]  {row['ES_label']}  "
              f"Len={row['Sequence Length']}  "
              f"+{best_new} G  →  "
              f"{len(covered)}/{len(universe)}  ({pct:.2f}%)")

        if covered >= universe:
            print(f"\n  ✓ Full coverage with {len(selected)} sequences!")
            break

    return selected, covered


# ─────────────────────────────────────────────────────────────
# 4.  Per-tail coverage stats
# ─────────────────────────────────────────────────────────────

def per_tail_stats(tail_g_map: dict, covered: set, selected: list) -> pd.DataFrame:
    """
    For each tail, compute:
      - total G-values in that tail
      - how many are covered
      - cover %
      - which ES sequences cover it (and how many G-values each covers in it)
    """
    # Build: g_key → ES label (which ES covers it)
    g_to_es = {}
    for s in selected:
        for gk in s['G_Keys_Newly_Covered'].split(' | '):
            gk = gk.strip()
            if gk:
                g_to_es[gk] = s['ES_Label']

    rows = []
    for roi_id, g_keys in sorted(tail_g_map.items()):
        total = len(g_keys)
        cov   = g_keys & covered
        n_cov = len(cov)
        pct   = round(n_cov / total * 100, 2) if total else 0

        # Which ES cover this tail, and how many G-values each
        es_counts = defaultdict(int)
        for gk in cov:
            es_label = g_to_es.get(gk, '?')
            es_counts[es_label] += 1

        es_detail = '; '.join(
            f"{es}({cnt}G)"
            for es, cnt in sorted(es_counts.items(),
                                   key=lambda x: -x[1])
        )

        rows.append({
            'Tail_ROI'            : roi_id,
            'Total_G_Values'      : total,
            'Covered_G_Values'    : n_cov,
            'Uncovered_G_Values'  : total - n_cov,
            'Cover_%'             : pct,
            'ES_Sequences_Used'   : '; '.join(sorted(es_counts.keys())),
            'ES_Detail (G count)' : es_detail,
        })

    return pd.DataFrame(rows).sort_values('Cover_%', ascending=False)


# ─────────────────────────────────────────────────────────────
# 5.  Main
# ─────────────────────────────────────────────────────────────

def main():
    tails_folder   = 'All_Tail_Proteins_Tables'
    sequences_file = 'Docs/Exact_Sequences_Miner_20260630_090923.csv'

    print("=" * 65)
    print("GLOBAL G-VALUE COVERAGE OPTIMIZER (full tail stats)")
    print("=" * 65)

    # ── load sequences ────────────────────────────────────────
    print(f"\n[1/5] Loading sequences …")
    df_orig = pd.read_csv(sequences_file)
    total_es = len(df_orig)
    df_orig['ES_label'] = [f"ES{i+1}" for i in range(total_es)]
    print(f"      {total_es:,} sequences (ES1 … ES{total_es})")

    # ── parse G-key sets ──────────────────────────────────────
    print(f"\n[2/5] Parsing Participating ROIs …")
    df_orig['_g_set']   = df_orig['Participating ROIs'].apply(parse_g_keys)
    df_orig['_g_count'] = df_orig['_g_set'].apply(len)
    df_work = df_orig[df_orig['_g_count'] > 0].copy()
    df_work = df_work.sort_values('_g_count', ascending=False).reset_index(drop=True)

    # ── load tail tables ──────────────────────────────────────
    print(f"\n[3/5] Loading tail tables from: {tails_folder}")
    tail_g_map = load_tail_g_values(tails_folder)
    print(f"      {len(tail_g_map)} tail tables loaded.")

    # Universe = only G-values that actually appear in a tail table
    # AND also appear in at least one ES (otherwise uncoverable)
    universe_from_tails: set = set()
    for g_keys in tail_g_map.values():
        universe_from_tails.update(g_keys)

    # Which of those are reachable by any ES?
    reachable: set = set()
    for gs in df_work['_g_set']:
        reachable.update(gs & universe_from_tails)

    universe = universe_from_tails          # full universe
    print(f"      Universe (G-values in tail tables)  : {len(universe):,}")
    print(f"      Reachable by at least one ES         : {len(reachable):,}")
    print(f"      Unreachable (no ES covers them)      : {len(universe - reachable):,}")

    # ── greedy cover ──────────────────────────────────────────
    print(f"\n[4/5] Running Greedy Set Cover …")
    print(f"      (covers only reachable G-values)\n")

    # Restrict ES to those covering at least one TAIL G-value
    mask = df_work['_g_set'].apply(lambda s: bool(s & universe))
    df_cover = df_work[mask].copy()
    print(f"      ES with ≥1 tail G-value: {len(df_cover):,}\n")

    selected, covered = greedy_cover(df_cover, reachable)

    # ── per-tail stats ────────────────────────────────────────
    print(f"\n[5/5] Computing per-tail statistics & exporting …")
    tail_df = per_tail_stats(tail_g_map, covered, selected)

    n_tails_covered    = (tail_df['Cover_%'] > 0).sum()
    n_tails_full       = (tail_df['Cover_%'] == 100).sum()
    avg_cover          = tail_df['Cover_%'].mean()
    max_cover          = tail_df['Cover_%'].max()
    total_g_universe   = len(universe)
    total_g_covered    = len(covered)
    overall_pct        = total_g_covered / total_g_universe * 100

    # ── build summary card ────────────────────────────────────
    es_labels_used = ', '.join(
        f"ES{s['ES_Label'][2:]}" if s['ES_Label'].startswith('ES') else s['ES_Label']
        for s in selected
    )
    tails_in_cover = '; '.join(
        f"{r['Tail_ROI']}({r['Cover_%']}%)"
        for _, r in tail_df[tail_df['Cover_%'] > 0].iterrows()
    )

    summary_rows = [
        ('Max coverage %',                    round(overall_pct, 4)),
        ('Number of tails covered (>0%)',      n_tails_covered),
        ('Number of tails fully covered (100%)', n_tails_full),
        ('Total tails in dataset',             len(tail_g_map)),
        ('Number of ES sequences used',        len(selected)),
        ('Average tail cover %',               round(avg_cover, 4)),
        ('Total G-values (universe)',           total_g_universe),
        ('Total G-values covered',             total_g_covered),
        ('Total G-values uncovered',           total_g_universe - total_g_covered),
        ('ES sequences used in cover',         es_labels_used),
        ('Tails included in cover (with %)',   tails_in_cover),
    ]
    summary_df = pd.DataFrame(summary_rows, columns=['Metric', 'Value'])

    # ── export ────────────────────────────────────────────────
    os.makedirs('Docs', exist_ok=True)
    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx = f'Docs/GlobalGCoverFull_{ts}.xlsx'
    csv  = f'Docs/GlobalGCoverFull_{ts}.csv'

    detail_df = pd.DataFrame(selected)

    with pd.ExcelWriter(xlsx, engine='openpyxl') as writer:
        summary_df.to_excel(writer,  sheet_name='Summary',         index=False)
        detail_df.to_excel(writer,   sheet_name='Selected ES',      index=False)
        tail_df.to_excel(writer,     sheet_name='Per-Tail Stats',   index=False)

    detail_df.to_csv(csv, index=False)

    print(f"\n  Excel → {xlsx}")
    print(f"  CSV   → {csv}")

    print("\n" + "=" * 65)
    print("FINAL SUMMARY")
    print("=" * 65)
    for label, val in summary_rows[:10]:
        print(f"  {label:<45}: {val}")
    print("=" * 65)


if __name__ == '__main__':
    main()
