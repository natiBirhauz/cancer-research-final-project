"""
PureGreedyCover.py
------------------
Greedy Set Cover אחד נקי — ללא anchor, סורק את כל 491K ES.

אלגוריתם:
  בכל צעד: בחר את ה-ES שמוסיף הכי הרבה G-values חדשים לכיסוי הגלובלי.
  עצור כשאין עוד G חדש לכסות.
  אחר כך: Backward Pruning — הסר ES מיותרים.

אופטימיזציה לביצועים:
  Inverted index: G-key → רשימת ES-ים שמכסים אותו.
  בכל צעד: רק ES-ים שמכסים G-values שעדיין לא כוסו הם רלוונטיים.
  לכן לא צריך לסרוק את כל 491K בכל צעד.

פלט Excel (4 גליונות):
  1. Summary          — סטטיסטיקות
  2. Greedy_Cover     — ES-ים שנבחרו לפי סדר הבחירה + delta coverage
  3. Per_ROI_Coverage — כיסוי ייחודי לכל ROI
  4. ROI_ES_Matrix    — מטריצה: ROI × ES (G אחד לתא, ללא כפילויות)
"""

import pandas as pd
import numpy as np
import re
import os
import glob
import time
from collections import defaultdict
from typing import Dict, FrozenSet, Set, List, Tuple

# ─────────────────────────────────────────────────────────────────
TAILS_FOLDER = 'All_Tail_Proteins_Tables'


# ─────────────────────────────────────────────────────────────────
# LOADING
# ─────────────────────────────────────────────────────────────────

def find_source_file():
    patterns = ['Docs/Exact_Sequences_Miner_*.csv',
                'Docs/CoreESAnalysis_PartA_*.csv']
    files = []
    for p in patterns:
        files.extend(glob.glob(p))
    return max(files, key=os.path.getctime) if files else None


def parse_g_keys(roi_string: str) -> FrozenSet[str]:
    if pd.isna(roi_string):
        return frozenset()
    keys = set()
    for part in str(roi_string).split(' | '):
        key = re.sub(r'\s*\([+-]\)\s*$', '', part).strip()
        if key:
            keys.add(key)
    return frozenset(keys)


def parse_g_sign_map(roi_string: str) -> Dict[str, str]:
    result = {}
    if pd.isna(roi_string):
        return result
    for part in str(roi_string).split(' | '):
        m = re.match(r'(.+?)_G(\d+)\s*\(([-+])\)', part.strip())
        if m:
            roi_prefix, g_num, sign = m.groups()
            result[f"{roi_prefix.strip()}_G{g_num}"] = f"G{g_num}({sign})"
    return result


def load_tail_g_values(folder: str) -> Dict[str, FrozenSet[str]]:
    files = sorted(glob.glob(os.path.join(folder, '*.xlsx')) +
                   glob.glob(os.path.join(folder, '*.csv')))
    result = {}
    for f in files:
        if os.path.basename(f).startswith('~$'):
            continue
        roi_id = (os.path.basename(f)
                  .replace('_tail_proteins.xlsx', '')
                  .replace('_tail_proteins.csv', ''))
        try:
            df = pd.read_excel(f) if f.endswith('.xlsx') else pd.read_csv(f)
        except Exception as e:
            print(f"  Warning: {f}: {e}")
            continue
        g_nums = sorted(set(
            int(re.search(r'G(\d+)', c).group(1))
            for c in df.columns if re.match(r'G\d+[+-]$', c)
        ))
        g_keys = frozenset(f"{roi_id}_G{n}" for n in g_nums)
        if g_keys:
            result[roi_id] = g_keys
    return result


# ─────────────────────────────────────────────────────────────────
# COVERAGE METRICS
# ─────────────────────────────────────────────────────────────────

def compute_unique_coverage(selected: list,
                             tail_g_map: Dict[str, FrozenSet[str]]
                             ) -> Tuple[float, int, Dict[str, float]]:
    """Avg unique G-coverage % across all ROIs (same G counted once per ROI)."""
    roi_pct = {}
    for roi_id, g_keys in tail_g_map.items():
        unique: Set[str] = set()
        for r in selected:
            unique.update(r['g_set'] & g_keys)
        roi_pct[roi_id] = len(unique) / len(g_keys) * 100 if g_keys else 0.0
    avg   = float(np.mean(list(roi_pct.values())))
    n_cov = sum(1 for p in roi_pct.values() if p > 0)
    return avg, n_cov, roi_pct


# ─────────────────────────────────────────────────────────────────
# STEP 1 — PURE GREEDY (no anchor, all ES considered)
# ─────────────────────────────────────────────────────────────────

def pure_greedy(all_rows: list,
                universe: FrozenSet[str],
                tail_g_map: Dict[str, FrozenSet[str]]
                ) -> Tuple[list, list]:
    """
    Single Greedy run over ALL ES sequences.

    Speed trick — Inverted index:
      uncovered_g_to_es[g_key] = set of ES indices still covering that G-key
      When we need the best next ES: only look at ES-ims touching uncovered G-values.
      After selecting an ES: remove its G-keys from uncovered_g_to_es.

    At each step: pick the ES that covers the most still-uncovered G-values.

    Returns:
      snapshots : list of dicts (one per selected ES, in selection order)
      selected  : list of row dicts in selection order
    """
    print(f"  Universe: {len(universe):,} G-values  |  Pool: {len(all_rows):,} ES")
    print("  Building inverted index ...")

    # index: g_key → set of row indices
    inverted: Dict[str, Set[int]] = defaultdict(set)
    for idx, r in enumerate(all_rows):
        for gk in r['g_set']:
            inverted[gk].add(idx)

    # count[idx] = how many uncovered G-values ES[idx] currently covers
    count: Dict[int, int] = {idx: len(r['g_set']) for idx, r in enumerate(all_rows)}

    # Use a max-count structure: bucket_sort by count
    # max_count[n] = set of ES indices with exactly n uncovered G-values
    max_val = max(count.values()) if count else 0
    buckets: Dict[int, Set[int]] = defaultdict(set)
    for idx, n in count.items():
        buckets[n].add(idx)

    selected_indices: Set[int] = set()
    covered: Set[str] = set()
    selected: list = []
    snapshots: list = []

    print(f"  Starting greedy loop ...")
    step = 0

    while True:
        step += 1

        # Find the ES with max uncovered G-values
        best_idx = None
        while max_val > 0:
            if buckets[max_val]:
                best_idx = next(iter(buckets[max_val]))
                break
            max_val -= 1

        if best_idx is None or max_val == 0:
            print(f"  Stopped at step {step}: no more uncovered G-values.")
            break

        # Select this ES
        r = all_rows[best_idx]
        new_g = r['g_set'] - covered

        if not new_g:
            # This shouldn't happen with correct bucket management, but safety check
            buckets[max_val].discard(best_idx)
            continue

        # Update state
        covered.update(new_g)
        selected_indices.add(best_idx)
        buckets[count[best_idx]].discard(best_idx)
        count[best_idx] = 0

        # For every newly covered G-key, decrement count of all ES that also cover it
        for gk in new_g:
            for other_idx in inverted[gk]:
                if other_idx in selected_indices:
                    continue
                old_n = count[other_idx]
                if old_n > 0:
                    buckets[old_n].discard(other_idx)
                    count[other_idx] = old_n - 1
                    buckets[old_n - 1].add(other_idx)

        selected.append(r)

        pct_global  = len(covered) / len(universe) * 100
        n_rois_cov  = sum(
            1 for roi_id, g_keys in tail_g_map.items()
            if covered & g_keys
        )

        snapshots.append({
            'Step'                 : step,
            'ES_Label'             : r['label'],
            'Sequence_Length'      : r['seq_len'],
            'New_G_Covered'        : len(new_g),
            'Total_G_Covered'      : len(covered),
            'Global_Coverage_%'    : round(pct_global, 4),
            'ROIs_Covered'         : n_rois_cov,
            'ROIs_Total'           : len(tail_g_map),
            'Exact_Sign_Sequence'  : r['sequence'],
        })

        print(f"  [{step:>5}]  {r['label']:>12}  "
              f"Len={r['seq_len']:>2}  "
              f"+{len(new_g):>4} G  →  "
              f"{len(covered):>5}/{len(universe)}  "
              f"({pct_global:.2f}%)  "
              f"ROIs={n_rois_cov}/{len(tail_g_map)}")

        if covered >= universe:
            print(f"  ✓ Full coverage at step {step}!")
            break

    return snapshots, selected


# ─────────────────────────────────────────────────────────────────
# STEP 2 — BACKWARD PRUNING
# ─────────────────────────────────────────────────────────────────

def backward_prune(selected: list,
                   tail_g_map: Dict[str, FrozenSet[str]]) -> list:
    """Remove ES whose removal doesn't reduce avg_unique_coverage."""
    pruned   = list(selected)
    base_avg, _, _ = compute_unique_coverage(pruned, tail_g_map)
    removed_count  = 0
    changed = True

    print(f"  Pruning {len(pruned)} ES (base avg={base_avg:.4f}%) ...")
    while changed:
        changed = False
        for i in range(len(pruned) - 1, -1, -1):
            test = pruned[:i] + pruned[i+1:]
            if not test:
                continue
            avg, _, _ = compute_unique_coverage(test, tail_g_map)
            if avg >= base_avg - 1e-9:
                removed = pruned.pop(i)
                base_avg = avg
                removed_count += 1
                changed  = True
                print(f"    Removed {removed['label']} "
                      f"(0 unique contribution)  → {len(pruned)} ES left")
                break   # restart after each removal

    print(f"  Pruning done: removed {removed_count}, kept {len(pruned)} ES")
    return pruned


# ─────────────────────────────────────────────────────────────────
# STEP 3 — CONFLICT-FREE G ASSIGNMENT
# ─────────────────────────────────────────────────────────────────

def assign_conflict_free(roi_id: str,
                          es_list: list,
                          tail_g_set: FrozenSet[str]) -> Dict[str, str]:
    """One unique G per ES column for this ROI, no repeats."""
    assigned: Set[str] = set()
    assignments: Dict[str, str] = {}

    es_cands = []
    for r in es_list:
        cands = [(gk, r['g_sign_map'].get(gk, gk.split('_G')[-1]))
                 for gk in (r['g_set'] & tail_g_set)]
        if cands:
            es_cands.append((r['label'], cands))

    es_cands.sort(key=lambda x: len(x[1]))   # fewest candidates first

    for es_label, cands in es_cands:
        chosen = None
        for full_gk, g_val in cands:
            if full_gk not in assigned:
                chosen = g_val
                assigned.add(full_gk)
                break
        assignments[es_label] = chosen or ''

    return assignments


# ─────────────────────────────────────────────────────────────────
# OUTPUT
# ─────────────────────────────────────────────────────────────────

def build_and_export(snapshots, optimal, tail_g_map, universe):
    avg_cov, n_rois_cov, roi_pct = compute_unique_coverage(optimal, tail_g_map)

    # ── Summary ──────────────────────────────────────────────
    global_cov = set()
    for r in optimal:
        global_cov.update(r['g_set'])

    summary_data = [
        ('Total ES in dataset',                  'see source CSV'),
        ('Total ROIs',                           len(tail_g_map)),
        ('Total G-values in universe',           len(universe)),
        ('─── AFTER GREEDY + PRUNING ──────────', ''),
        ('Number of ES in optimal combination',  len(optimal)),
        ('ROIs covered (>0%)',                   n_rois_cov),
        ('Average unique G-coverage % per ROI',  f"{avg_cov:.4f}%"),
        ('Global G-values covered',              len(global_cov)),
        ('Global coverage %',                    f"{len(global_cov)/len(universe)*100:.4f}%"),
        ('ES labels in combination',
         ', '.join(r['label'] for r in optimal)),
    ]
    df_summary = pd.DataFrame(summary_data, columns=['Metric', 'Value'])

    # ── Greedy Cover detail ───────────────────────────────────
    df_greedy = pd.DataFrame(snapshots)

    # ── Per-ROI Coverage ──────────────────────────────────────
    roi_rows = []
    for roi_id in sorted(tail_g_map):
        total_g = len(tail_g_map[roi_id])
        pct     = roi_pct[roi_id]
        unique  = set()
        for r in optimal:
            unique.update(r['g_set'] & tail_g_map[roi_id])
        contrib = [r['label'] for r in optimal
                   if r['g_set'] & tail_g_map[roi_id]]
        roi_rows.append({
            'ROI'                    : roi_id,
            'Total_G_Values'         : total_g,
            'Unique_G_Covered'       : len(unique),
            'Coverage_%'             : round(pct, 2),
            'Contributing_ES_Count'  : len(contrib),
            'Contributing_ES_Labels' : ', '.join(contrib),
        })
    df_roi = pd.DataFrame(roi_rows).sort_values('Coverage_%', ascending=False)

    # ── ROI × ES Matrix ───────────────────────────────────────
    print("  Building conflict-free ROI × ES matrix ...")
    matrix_rows = []
    for roi_id, g_keys in sorted(tail_g_map.items()):
        asgn = assign_conflict_free(roi_id, optimal, g_keys)
        row  = {'ROI': roi_id,
                'Total_G': len(g_keys),
                'Coverage_%': round(roi_pct[roi_id], 2)}
        for r in optimal:
            row[r['label']] = asgn.get(r['label'], '')
        matrix_rows.append(row)
    df_matrix = pd.DataFrame(matrix_rows)
    meta = ['ROI', 'Total_G', 'Coverage_%']
    df_matrix = df_matrix[meta + [r['label'] for r in optimal]].fillna('')

    # ── Export ────────────────────────────────────────────────
    os.makedirs('Docs', exist_ok=True)
    ts  = time.strftime("%Y%m%d_%H%M%S")
    out = f'Docs/PureGreedy_Cover_{ts}.xlsx'

    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df_summary.to_excel(writer, sheet_name='Summary',          index=False)
        df_greedy.to_excel(writer,  sheet_name='Greedy_Steps',     index=False)
        df_roi.to_excel(writer,     sheet_name='Per_ROI_Coverage', index=False)
        df_matrix.to_excel(writer,  sheet_name='ROI_ES_Matrix',    index=False)

    print(f"\n  Saved: {out}")
    return out


# ─────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────

def main():
    t0 = time.time()
    print("=" * 65)
    print("PURE GREEDY SET COVER  (all ES, no anchor)")
    print("=" * 65)

    # Load
    src = find_source_file()
    if not src:
        print("ERROR: No ES CSV found in Docs/")
        return
    print(f"\n[1/4] Loading: {src}")
    df = pd.read_csv(src)
    for col in ('ES_Label', 'ES_label'):
        if col in df.columns:
            df['_label'] = df[col].astype(str).str.strip()
            break
    else:
        df['_label'] = [f"ES{i+1}" for i in range(len(df))]
    print(f"      {len(df):,} ES loaded")

    print(f"\n[2/4] Loading tail tables from: {TAILS_FOLDER}")
    tail_g_map = load_tail_g_values(TAILS_FOLDER)
    universe   = frozenset().union(*tail_g_map.values())
    print(f"      {len(tail_g_map)} ROIs  |  {len(universe):,} G-values")

    # Parse rows
    print(f"\n[3/4] Parsing ES data ...")
    target_col = ('Participating ROIs' if 'Participating ROIs' in df.columns
                  else 'Tails_Detail')
    all_rows = []
    for _, row in df.iterrows():
        roi_str = row[target_col]
        g_set   = parse_g_keys(roi_str) & universe
        if not g_set:
            continue
        all_rows.append({
            'label'      : row['_label'],
            'seq_len'    : int(row['Sequence Length']),
            'patients'   : row['Unique Patients Count'],
            'g_set'      : g_set,
            'g_sign_map' : parse_g_sign_map(roi_str),
            'sequence'   : row['Exact Sign Sequence'],
            'part_rois'  : roi_str,
        })
    print(f"      {len(all_rows):,} ES with tail G-values")

    # Step 1: Pure Greedy
    print(f"\n[4/4] Running Pure Greedy over ALL {len(all_rows):,} ES ...")
    snapshots, greedy_selected = pure_greedy(all_rows, universe, tail_g_map)

    # Step 2: Backward Pruning
    print(f"\n  Running Backward Pruning ...")
    optimal = backward_prune(greedy_selected, tail_g_map)

    # Evaluate
    avg_cov, n_rois, _ = compute_unique_coverage(optimal, tail_g_map)
    print(f"\n  RESULT: {len(optimal)} ES | ROIs={n_rois}/{len(tail_g_map)} | "
          f"avg={avg_cov:.4f}%")

    # Export
    out = build_and_export(snapshots, optimal, tail_g_map, universe)

    elapsed = time.time() - t0
    print(f"\n  Total time: {elapsed:.0f}s")
    print("=" * 65)


if __name__ == '__main__':
    main()
