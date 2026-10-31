"""
CoreESAnalysis.py
-----------------
Two complementary analyses to find the dominant ("core") ES sequences:

PART A — Independent Scoring (overlap allowed)
  For every ES sequence, independently compute:
    - How many G-values from the tail tables it covers
    - How many different tail tables it appears in
    - What % of the total universe it covers
  Result: ranking of ALL 491K ES by dominance score.
  This is the true "core" — no combination logic, just raw coverage per ES.

PART B — Greedy Minimum Set Cover (no G-value repeated)
  Standard greedy Set Cover (NP-hard optimal approximation): 
    - Tie: most G-values covered
  Result: minimum ES set that covers maximum G-values with no overlap.

Excel Output (5 sheets):
  1. Summary          — global stats for both parts
  2. Part A — All ES Ranked   — all 491K ES scored, sorted by dominance
  3. Part A — Top 500         — top 500 dominant ES for easy viewing
  4. Part B — Greedy Cover    — minimum cover set, ranked by selection order
  5. Per-Tail Coverage        — for each tail: Part A best ES + Part B coverage
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
TAILS_FOLDER     = 'All_Tail_Proteins_Tables'
SEQUENCES_FILE   = 'Docs/Exact_Sequences_Miner_20260915_132012.csv'
TOP_N            = 500      # how many top ES to show in the "Top N" sheet


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


def parse_g_keys_with_orient(s: str) -> dict:
    """
    'P1_ROI_03_G4 (+) | P9_ROI_51_G13 (-)'
    → { 'P1_ROI_03_G4': '+', 'P9_ROI_51_G13': '-' }

    The trailing (+/-) is the G-vector orientation relative to this ES:
      '+' → table columns are already aligned  (G{n}+ proteins have '+' sign)
      '-' → table columns are flipped          (G{n}+ proteins have '-' sign)
    """
    if pd.isna(s):
        return {}
    result = {}
    for part in str(s).split(' | '):
        part = part.strip()
        m = re.match(r'^(.+?)\s*\(([+-])\)\s*$', part)
        if m:
            result[m.group(1).strip()] = m.group(2)
    return result


def parse_es_proteins(s: str) -> dict:
    """
    'CD163(+) | CD21(+) | CXCR3(-)'
    → { 'CD163': '+', 'CD21': '+', 'CXCR3': '-' }
    """
    if pd.isna(s):
        return {}
    result = {}
    for part in str(s).split(' | '):
        part = part.strip()
        m = re.match(r'^(.+?)\(([+-])\)\s*$', part)
        if m:
            result[m.group(1).strip()] = m.group(2)
    return result


def load_tail_g_values(tails_folder: str) -> tuple:
    """
    Returns:
      tail_g_map     : { roi_id → frozenset of G-key strings }
      g_protein_count: { g_key → int }   total proteins in that G-vector
      g_proteins     : { g_key → dict{protein_name → sign('+'/'-')} }
                       canonical (un-flipped) mapping — '+' means the protein
                       lives in the G{n}+ column, '-' in G{n}-.
    """
    files = sorted(
        glob.glob(os.path.join(tails_folder, '*.xlsx')) +
        glob.glob(os.path.join(tails_folder, '*.csv'))
    )
    tail_g_map      = {}
    g_protein_count = {}
    g_proteins      = {}   # NEW: g_key → {protein → raw_sign}
    for f in files:
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
        if not g_keys:
            continue
        tail_g_map[roi_id] = g_keys
        for n in g_nums:
            gkey   = f"{roi_id}_G{n}"
            pmap   = {}   # protein_name → raw sign in the table ('+' or '-')
            for sign in ('+', '-'):
                col = f'G{n}{sign}'
                if col in df.columns:
                    for name in df[col].dropna():
                        pmap[str(name).strip()] = sign
            g_proteins[gkey]      = pmap
            g_protein_count[gkey] = len(pmap)
    return tail_g_map, g_protein_count, g_proteins


# ─────────────────────────────────────────────
# PART A — INDEPENDENT SCORING
# ─────────────────────────────────────────────

def score_all_es(df: pd.DataFrame,
                 tail_g_map: dict,
                 universe: frozenset,
                 g_protein_count: dict,
                 g_proteins: dict) -> pd.DataFrame:
    """
    For every ES, independently compute:

      Per-G protein coverage
      ──────────────────────
      For each G-vector the ES appears in:
          matched = |ES_protein_names ∩ G_protein_names|
          per_G_coverage = matched / g_protein_count[gkey] * 100
      Coverage is capped at 100% (matched can't exceed g_total by definition,
      but we cap defensively).

      Avg_Protein_Coverage_Per_G_% = mean(per_G_coverage) over all G-vectors

      Tails_Detail
      ────────────
      Per-ROI breakdown listing every G-vector the ES touches:
        ROI(N/totalG G | GX:matched/g_total=pct%; ...)

      G_Values_Covered : # G-vectors from tail tables this ES appears in
      Tails_Covered    : # distinct ROI tail-tables this ES appears in
      Pct_Universe_Covered : G_Values_Covered / total_tail_G * 100

    Sorted descending by (Sequence Length, G_Values_Covered, Tails_Covered).
    """
    print("  Scoring all ES sequences independently ...")

    tail_list      = list(tail_g_map.items())
    total_universe = len(universe)

    records = []
    n = len(df)
    report_every = max(1, n // 20)

    for i, (idx, row) in enumerate(df.iterrows()):
        if i % report_every == 0:
            print(f"    {i:>7}/{n}  ({i/n*100:.0f}%) ...")

        g_set = row['_g_set']
        if not g_set:
            continue

        tail_g_covered = g_set & universe
        n_covered = len(tail_g_covered)
        if n_covered == 0:
            continue

        # ES protein names (signs ignored for name-match coverage)
        es_prot_map   = parse_es_proteins(row['Exact Sign Sequence'])  # {name→sign}
        es_prot_names = set(es_prot_map.keys())                        # just names
        if not es_prot_names:
            continue

        # ── Per-G coverage: |ES_names ∩ G_names| / g_total ──────────────
        per_g_ratios = []
        g_per_roi: dict[str, list] = defaultdict(list)  # roi → [(gnum, matched, g_total, pct)]

        for gkey in tail_g_covered:
            g_pmap  = g_proteins.get(gkey, {})           # {protein_name → sign}
            g_total = g_protein_count.get(gkey, 0)
            if g_total == 0:
                continue

            matched = len(es_prot_names & set(g_pmap.keys()))
            pct_g   = min(matched / g_total * 100, 100.0)  # cap at 100%
            per_g_ratios.append(pct_g)

            roi_id = re.sub(r'_G\d+$', '', gkey)
            gnum   = re.search(r'_G(\d+)$', gkey).group(1)
            g_per_roi[roi_id].append((int(gnum), matched, g_total, pct_g))

        if not per_g_ratios:
            continue

        avg_prot_cov = round(sum(per_g_ratios) / len(per_g_ratios), 4)

        # ── Per-tail G-count (G_Values_Covered / Tails_Covered) ──────────
        tail_g_hits: dict[str, int] = {}
        for roi_id, g_keys in tail_list:
            hit = len(g_set & g_keys)
            if hit > 0:
                tail_g_hits[roi_id] = hit

        n_tails  = len(tail_g_hits)
        pct_univ = n_covered / total_universe * 100

        # ── Build Tails_Detail ────────────────────────────────────────────
        tails_detail_parts = []
        for roi_id, g_cnt in sorted(tail_g_hits.items(), key=lambda x: -x[1]):
            g_entries       = sorted(g_per_roi.get(roi_id, []))  # sort by G-number
            total_g_in_tail = len(tail_g_map[roi_id])
            g_str = ', '.join(
                f"G{gnum}:{matched}/{gtot}={pct:.1f}%"
                for gnum, matched, gtot, pct in g_entries
            )
            tails_detail_parts.append(
                f"{roi_id}({g_cnt}/{total_g_in_tail}G | {g_str})"
            )
        tails_str = '; '.join(tails_detail_parts)

        records.append({
            'ES_Label'                      : row['ES_label'],
            'Sequence Length'               : int(row['Sequence Length']),
            'Unique Patients Count'         : row['Unique Patients Count'],
            'G_Values_Covered'              : n_covered,
            'Tails_Covered'                 : n_tails,
            'Pct_Universe_Covered'          : round(pct_univ, 4),
            'Avg_Protein_Coverage_Per_G_%'  : avg_prot_cov,
            'Exact Sign Sequence'           : row['Exact Sign Sequence'],
            'Tails_Detail'                  : tails_str,
            'Participating ROIs'            : row['Participating ROIs'],
        })

    result_df = pd.DataFrame(records)
    result_df.sort_values(
        by=['Sequence Length', 'G_Values_Covered', 'Tails_Covered'],
        ascending=[False, False, False],
        inplace=True
    )
    result_df.insert(0, 'Rank', range(1, len(result_df) + 1))
    result_df.reset_index(drop=True, inplace=True)
    return result_df

# ─────────────────────────────────────────────
# PART B — GREEDY MINIMUM COVER
# ─────────────────────────────────────────────

def greedy_cover(df: pd.DataFrame, universe: frozenset) -> tuple:
    """
    Greedy Set Cover over `universe`.
    Priority: longest Sequence Length first → most G-values covered → first found.
    No G-value assigned to more than one ES.
    Returns (selected_list, covered_set).
    """
    print("  Running greedy set cover ...")

    covered : set  = set()
    selected: list = []

    # Build a plain list of dicts for fast access — avoids column name issues
    rows = []
    for idx, row in df.iterrows():
        rows.append({
            'idx'       : idx,
            'g_set'     : row['_g_set'],
            'seq_len'   : int(row['Sequence Length']),
            'es_label'  : row['ES_label'],
            'patients'  : row['Unique Patients Count'],
            'sequence'  : row['Exact Sign Sequence'],
            'part_rois' : row['Participating ROIs'],
        })

    # Pre-sort by seq_len DESC then g_set size DESC (stable)
    rows.sort(key=lambda r: (-r['seq_len'], -len(r['g_set'])))

    while rows:
        best_pos       = None
        best_new_count = 0
        best_new_set   : set = set()
        best_len       = -1

        for pos, r in enumerate(rows):
            new   = r['g_set'] - covered
            n_new = len(new)
            if n_new == 0:
                continue
            if r['seq_len'] > best_len or \
               (r['seq_len'] == best_len and n_new > best_new_count):
                best_pos       = pos
                best_new_count = n_new
                best_new_set   = new
                best_len       = r['seq_len']

        if best_pos is None:
            break

        r = rows.pop(best_pos)
        covered.update(best_new_set)
        pct = len(covered) / len(universe) * 100

        selected.append({
            'Rank'                   : len(selected) + 1,
            'ES_Label'               : r['es_label'],
            'Sequence Length'        : best_len,
            'Unique Patients Count'  : r['patients'],
            'New_G_Covered'          : best_new_count,
            'Total_G_Covered_So_Far' : len(covered),
            'Coverage_%'             : round(pct, 4),
            'Exact Sign Sequence'    : r['sequence'],
            'G_Keys_Newly_Covered'   : ' | '.join(sorted(best_new_set)),
            'Participating ROIs'     : r['part_rois'],
        })

        print(f"    [{len(selected):>4}]  {r['es_label']}"
              f"  Len={best_len:>2}  +{best_new_count} G"
              f"  → {len(covered)}/{len(universe)}  ({pct:.2f}%)")

        if covered >= universe:
            print(f"  ✓ Full coverage with {len(selected)} ES!")
            break

    return selected, covered


# ─────────────────────────────────────────────
# PER-TAIL SUMMARY
# ─────────────────────────────────────────────

def build_per_tail_summary(tail_g_map: dict,
                            part_a_df: pd.DataFrame,
                            greedy_selected: list,
                            greedy_covered: set) -> pd.DataFrame:
    """
    For each tail:
      - Best single ES from Part A (highest G_Values_Covered for that tail)
      - Part B coverage: how many of its G-values are covered by the greedy set
    """
    # Map G-key → ES label from greedy
    g_to_es = {}
    for s in greedy_selected:
        for gk in s['G_Keys_Newly_Covered'].split(' | '):
            gk = gk.strip()
            if gk:
                g_to_es[gk] = s['ES_Label']

    rows = []
    for roi_id, g_keys in sorted(tail_g_map.items()):
        total_g = len(g_keys)

        # Part B: greedy coverage of this tail
        cov_b = [gk for gk in g_keys if gk in g_to_es]
        n_b   = len(cov_b)
        pct_b = round(n_b / total_g * 100, 2) if total_g else 0
        es_b  = '; '.join(sorted(set(g_to_es[gk] for gk in cov_b)))

        rows.append({
            'Tail_ROI'                      : roi_id,
            'Total_G_Values'                : total_g,
            'PartB_G_Covered'               : n_b,
            'PartB_Coverage_%'              : pct_b,
            'PartB_ES_Labels'               : es_b,
        })

    return pd.DataFrame(rows).sort_values('PartB_Coverage_%', ascending=False)


# ─────────────────────────────────────────────
# BUILD SUMMARY
# ─────────────────────────────────────────────

def build_summary(universe, tail_g_map,
                  part_a_df, greedy_selected, greedy_covered) -> pd.DataFrame:

    total_g   = len(universe)
    total_es  = len(part_a_df)

    # Part A stats
    a_top1    = part_a_df.iloc[0] if len(part_a_df) > 0 else None
    a_top_pct = a_top1['Pct_Universe_Covered'] if a_top1 is not None else 0

    # Part B stats
    b_pct     = len(greedy_covered) / total_g * 100 if total_g else 0
    b_n_es    = len(greedy_selected)

    # Tails covered by Part B
    g_to_es = {}
    for s in greedy_selected:
        for gk in s['G_Keys_Newly_Covered'].split(' | '):
            gk = gk.strip()
            if gk: g_to_es[gk] = s['ES_Label']

    tail_pct_strs = []
    n_tails_cov   = 0
    for roi_id, g_keys in sorted(tail_g_map.items()):
        n = sum(1 for gk in g_keys if gk in g_to_es)
        p = round(n / len(g_keys) * 100, 1) if g_keys else 0
        if n > 0: n_tails_cov += 1
        tail_pct_strs.append(f"{roi_id}({p}%)")

    avg_tail_pct = (sum(float(s.split('(')[1].rstrip('%)'))
                        for s in tail_pct_strs) / len(tail_pct_strs)
                    if tail_pct_strs else 0)

    es_b_labels = ', '.join(s['ES_Label'] for s in greedy_selected)

    rows = [
        ('── PART A (independent scoring) ──────────────────', ''),
        ('Total ES sequences scored',               total_es),
        ('Total G-values in universe',              total_g),
        ('Best single ES (most G-values covered)',  a_top1['ES_Label'] if a_top1 is not None else ''),
        ('Best single ES sequence',                 a_top1['Exact Sign Sequence'] if a_top1 is not None else ''),
        ('Best single ES % universe covered',       a_top_pct),
        ('Best single ES tails covered',            a_top1['Tails_Covered'] if a_top1 is not None else 0),
        ('── PART B (greedy minimum cover) ─────────────────', ''),
        ('Max % coverage (greedy)',                 round(b_pct, 4)),
        ('Number of ES sequences in greedy cover',  b_n_es),
        ('Number of tails covered >0% (greedy)',    n_tails_cov),
        ('Total tails in dataset',                  len(tail_g_map)),
        ('Average tail coverage % (greedy)',        round(avg_tail_pct, 2)),
        ('G-values covered (greedy)',               len(greedy_covered)),
        ('G-values uncovered (greedy)',             total_g - len(greedy_covered)),
        ('ES sequences used in greedy cover',       es_b_labels),
        ('Tails included in greedy cover (with %)', ', '.join(tail_pct_strs)),
    ]
    return pd.DataFrame(rows, columns=['Metric', 'Value'])


# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    print("=" * 65)
    print("CORE ES ANALYSIS  (Part A: scoring  |  Part B: greedy cover)")
    print("=" * 65)

    # ── Load sequences ────────────────────────────────────────
    print(f"\n[1/6] Loading sequences from:\n      {SEQUENCES_FILE}")
    df = pd.read_csv(SEQUENCES_FILE)
    total_es = len(df)
    df['ES_label'] = [f"ES{i+1}" for i in range(total_es)]
    print(f"      {total_es:,} sequences loaded")

    print("      Parsing Participating ROIs ...")
    df['_g_set']   = df['Participating ROIs'].apply(parse_g_keys)
    df['_g_count'] = df['_g_set'].apply(len)
    df = df[df['_g_count'] > 0].copy()
    df.reset_index(drop=True, inplace=True)
    print(f"      {len(df):,} ES with ≥1 G-value retained")

    # ── Load tail tables ──────────────────────────────────────
    print(f"\n[2/6] Loading tail tables from: {TAILS_FOLDER}")
    tail_g_map, g_protein_count, g_proteins = load_tail_g_values(TAILS_FOLDER)
    universe   = frozenset().union(*tail_g_map.values())
    print(f"      {len(tail_g_map)} tail tables  |  {len(universe):,} unique G-values")

    # Restrict to ES that touch at least one tail G-value
    df_tail = df[df['_g_set'].apply(lambda s: bool(s & universe))].copy()
    df_tail.reset_index(drop=True, inplace=True)
    print(f"      ES covering ≥1 tail G-value: {len(df_tail):,}")

    # ── Part A: score every ES independently ─────────────────
    print(f"\n[3/6] Part A — scoring every ES independently ...")
    part_a_df = score_all_es(df_tail, tail_g_map, universe, g_protein_count, g_proteins)
    print(f"      Scored {len(part_a_df):,} ES sequences")

    # ── Part B: greedy minimum cover ─────────────────────────
    print(f"\n[4/6] Part B — greedy minimum set cover ...")
    greedy_selected, greedy_covered = greedy_cover(df_tail, universe)
    greedy_df = pd.DataFrame(greedy_selected)
    print(f"      → {len(greedy_selected)} ES cover "
          f"{len(greedy_covered)}/{len(universe)} G-values "
          f"({len(greedy_covered)/len(universe)*100:.2f}%)")

    # ── Per-tail summary ──────────────────────────────────────
    print(f"\n[5/6] Building per-tail summary ...")
    per_tail_df = build_per_tail_summary(
        tail_g_map, part_a_df, greedy_selected, greedy_covered)

    # ── Build summary & export ────────────────────────────────
    print(f"\n[6/6] Building summary & exporting ...")
    summary_df = build_summary(
        universe, tail_g_map, part_a_df, greedy_selected, greedy_covered)

    os.makedirs('Docs', exist_ok=True)
    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx = f'Docs/CoreESAnalysis_{ts}.xlsx'
    csv  = f'Docs/CoreESAnalysis_PartA_{ts}.csv'

    with pd.ExcelWriter(xlsx, engine='openpyxl') as writer:
        summary_df.to_excel(writer,
                            sheet_name='Summary',          index=False)
        part_a_df.to_excel(writer,
                            sheet_name='PartA All ES Ranked', index=False)
        part_a_df.head(TOP_N).to_excel(writer,
                            sheet_name=f'PartA Top{TOP_N}',  index=False)
        greedy_df.to_excel(writer,
                            sheet_name='PartB Greedy Cover', index=False)
        per_tail_df.to_excel(writer,
                            sheet_name='Per-Tail Coverage',  index=False)

    part_a_df.to_csv(csv, index=False)

    print(f"\n  Excel → {xlsx}")
    print(f"  CSV   → {csv}")

    # ── Print summary ─────────────────────────────────────────
    print("\n" + "=" * 65)
    print("FINAL SUMMARY")
    print("=" * 65)
    for _, row in summary_df.iterrows():
        v = str(row['Value'])
        if len(v) > 80:
            v = v[:77] + '...'
        print(f"  {str(row['Metric']):<52}: {v}")
    print("=" * 65)
    print(f"\nOutput: {xlsx}")


if __name__ == '__main__':
    main()
