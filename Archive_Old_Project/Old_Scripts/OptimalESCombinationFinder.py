"""
OptimalESCombinationFinder.py
------------------------------
מוצא קומבינציות אופטימליות של ES sequences.

הגדרה מדויקת:
  - כל ES מכיל G-values ממספר ROIs/חולים
  - G מסוים נספר פעם אחת בלבד לכל ROI (ייחודי)
  - כיסוי ROI = G ייחודיים / סה"כ G של ה-ROI
  - מטרה: מקסימום avg_unique_coverage% עם מינימום ES

Part 1 — Coverage Curve (single greedy run):
  מריץ greedy אחד (בכל צעד מוסיף את ה-ES עם הכי הרבה G ייחודיים חדשים).
  שומר snapshot לאחר כל הוספה: n_ES → avg_coverage%.
  מייצר "עקומת כיסוי" — רואים את ה-elbow (היכן הכיסוי מתייצב).

Part 2 — 100 Best Combinations (בכל anchor שונה):
  100 ריצות greedy, כל אחת עם anchor שונה (ה-100 ES הכי דומיננטיים).
  בכל ריצה: greedy עד שאין שיפור, ואז Backward Pruning.
  שומר כל קומבינציה עם: n_ES, avg_coverage%, efficiency score.
  ממיין לפי: avg_coverage% DESC → n_ES ASC.
  הקומבינציה הראשונה = טובה ביותר.

פלט Excel:
  Sheet 1: Coverage_Curve        -- n_ES vs avg_coverage% (Part 1)
  Sheet 2: Summary_100_Combos    -- 100 קומבינציות ממוינות
  Sheet 3: ES_COM_1_Details      -- מטריצה ROI x ES ללא חפיפות
  Sheet 4: Per_ROI_Coverage      -- כיסוי ייחודי לכל ROI
"""

import pandas as pd
import numpy as np
import re
import os
import glob
import time
from typing import List, Dict, Tuple, FrozenSet, Set


# ─────────────────────────────────────────────────────────────────
TAILS_FOLDER    = 'All_Tail_Proteins_Tables'
N_COMBINATIONS  = 10


# ─────────────────────────────────────────────────────────────────
# FILE LOADING
# ─────────────────────────────────────────────────────────────────

def find_source_report():
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
        part = part.strip()
        m = re.match(r'(.+?)_G(\d+)\s*\(([-+])\)', part)
        if m:
            roi_prefix, g_num, sign = m.groups()
            full_key = f"{roi_prefix.strip()}_G{g_num}"
            result[full_key] = f"G{g_num}({sign})"
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
# UNIQUE COVERAGE  (core metric)
# ─────────────────────────────────────────────────────────────────

def compute_unique_coverage(selected: list,
                             tail_g_map: Dict[str, FrozenSet[str]]
                             ) -> Tuple[float, int, Dict[str, float]]:
    """
    Computes per-ROI unique G-coverage.
    Same G appearing in multiple ES is counted ONCE per ROI.

    Returns: (avg_pct, n_rois_covered, {roi_id: coverage%})
    """
    roi_pct = {}
    for roi_id, g_keys in tail_g_map.items():
        unique: Set[str] = set()
        for r in selected:
            unique.update(r['g_set'] & g_keys)
        pct = len(unique) / len(g_keys) * 100 if g_keys else 0.0
        roi_pct[roi_id] = pct
    avg   = float(np.mean(list(roi_pct.values())))
    n_cov = sum(1 for p in roi_pct.values() if p > 0)
    return avg, n_cov, roi_pct


# ─────────────────────────────────────────────────────────────────
# FAST INCREMENTAL: delta coverage when adding one ES
# ─────────────────────────────────────────────────────────────────

def delta_avg_coverage(candidate: dict,
                        current_covered: Dict[str, Set[str]],
                        tail_g_map: Dict[str, FrozenSet[str]]) -> float:
    """
    How much avg_coverage% increases if we add candidate to current combo.
    Uses the already-computed current_covered dict for speed.
    """
    total_rois = len(tail_g_map)
    total_delta = 0.0
    for roi_id, g_keys in tail_g_map.items():
        new_g = (candidate['g_set'] & g_keys) - current_covered[roi_id]
        total_delta += len(new_g) / len(g_keys) * 100
    return total_delta / total_rois


# ─────────────────────────────────────────────────────────────────
# PART 1 — COVERAGE CURVE (one greedy run, all snapshots)
# ─────────────────────────────────────────────────────────────────

def coverage_curve_greedy(all_rows: list,
                           tail_g_map: Dict[str, FrozenSet[str]]
                           ) -> Tuple[list, list]:
    """
    Single greedy run: at each step adds ES with maximum delta avg_coverage.
    Returns:
        snapshots : list of {n_es, avg_cov, roi_covered, es_label, delta}
        selected  : list of row dicts in selection order
    """
    print("  Running coverage curve greedy ...")

    # current_covered[roi_id] = set of G-keys already covered for that ROI
    current_covered = {roi_id: set() for roi_id in tail_g_map}
    cur_avg   = 0.0
    selected  = []
    snapshots = []
    remaining = list(all_rows)

    while remaining:
        best_pos   = None
        best_delta = 0.0

        for pos, r in enumerate(remaining):
            delta = delta_avg_coverage(r, current_covered, tail_g_map)
            if delta > best_delta:
                best_delta = delta
                best_pos   = pos

        if best_delta < 1e-9:
            break   # no ES adds any new coverage

        r = remaining.pop(best_pos)
        selected.append(r)

        # update current_covered
        for roi_id, g_keys in tail_g_map.items():
            current_covered[roi_id].update(r['g_set'] & g_keys)

        cur_avg += best_delta
        n_rois = sum(1 for v in current_covered.values() if v)

        snapshots.append({
            'n_ES'           : len(selected),
            'ES_Label'       : r['label'],
            'Delta_Avg_%'    : round(best_delta, 4),
            'Avg_Coverage_%' : round(cur_avg, 4),
            'ROIs_Covered'   : n_rois,
            'ROIs_Total'     : len(tail_g_map),
            'Efficiency'     : round(cur_avg / len(selected), 4),
        })

        print(f"    [{len(selected):>4}]  {r['label']:>12}  "
              f"+{best_delta:.3f}%  avg={cur_avg:.2f}%  "
              f"ROIs={n_rois}/{len(tail_g_map)}")

    return snapshots, selected


# ─────────────────────────────────────────────────────────────────
# PART 2 — MULTI-START GREEDY (100 anchors)
# ─────────────────────────────────────────────────────────────────

def greedy_from_anchor(anchor: dict,
                        all_rows: list,
                        tail_g_map: Dict[str, FrozenSet[str]]
                        ) -> list:
    """Fast greedy using delta_avg_coverage."""
    current_covered = {roi_id: set(anchor['g_set'] & g_keys)
                       for roi_id, g_keys in tail_g_map.items()}
    selected  = [anchor]
    remaining = [r for r in all_rows if r['label'] != anchor['label']]

    while remaining:
        best_pos   = None
        best_delta = 0.0
        for pos, r in enumerate(remaining):
            delta = delta_avg_coverage(r, current_covered, tail_g_map)
            if delta > best_delta:
                best_delta = delta
                best_pos   = pos

        if best_delta < 1e-9:
            break

        r = remaining.pop(best_pos)
        selected.append(r)
        for roi_id, g_keys in tail_g_map.items():
            current_covered[roi_id].update(r['g_set'] & g_keys)

    return selected


def backward_prune(selected: list,
                   tail_g_map: Dict[str, FrozenSet[str]]
                   ) -> list:
    """Remove ES whose removal doesn't decrease avg_unique_coverage."""
    pruned   = list(selected)
    base_avg, _, _ = compute_unique_coverage(pruned, tail_g_map)
    changed  = True
    while changed:
        changed = False
        for i in range(len(pruned) - 1, -1, -1):
            test = pruned[:i] + pruned[i+1:]
            if not test:
                continue
            avg, _, _ = compute_unique_coverage(test, tail_g_map)
            if avg >= base_avg - 1e-9:
                pruned.pop(i)
                base_avg = avg
                changed  = True
                break
    return pruned


# ─────────────────────────────────────────────────────────────────
# CONFLICT-FREE G ASSIGNMENT
# ─────────────────────────────────────────────────────────────────

def assign_conflict_free(roi_id: str,
                          es_list: list,
                          tail_g_set: FrozenSet[str]) -> Dict[str, str]:
    """One unique G per ES column, no repeats across columns for same ROI."""
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
# MAIN
# ─────────────────────────────────────────────────────────────────

def main():
    t0 = time.time()
    print("=" * 70)
    print("OPTIMAL ES COMBINATION FINDER")
    print("=" * 70)

    # Load
    report_path = find_source_report()
    if not report_path:
        print("ERROR: No ES sequence CSV found in Docs/")
        return
    print(f"\n[1/5] Loading from:\n      {report_path}")
    df = pd.read_csv(report_path)
    for col in ('ES_Label', 'ES_label'):
        if col in df.columns:
            df['_label'] = df[col].astype(str).str.strip()
            break
    else:
        df['_label'] = [f"ES{i+1}" for i in range(len(df))]
    print(f"      {len(df):,} ES loaded")

    print(f"\n[2/5] Loading tail tables from: {TAILS_FOLDER}")
    tail_g_map = load_tail_g_values(TAILS_FOLDER)
    universe   = frozenset().union(*tail_g_map.values())
    print(f"      {len(tail_g_map)} ROIs | {len(universe):,} G-values")

    print(f"\n[3/5] Parsing ES data ...")
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
            'part_rois'  : row[target_col],
        })
    all_rows.sort(key=lambda r: -len(r['g_set']))
    print(f"      {len(all_rows):,} ES with tail G-values")

    # ── Part 1: Coverage Curve ────────────────────────────────
    print(f"\n[4/5] Part 1 — Coverage curve (single greedy) ...")
    snapshots, _ = coverage_curve_greedy(all_rows, tail_g_map)
    df_curve = pd.DataFrame(snapshots)
    print(f"      Coverage curve: {len(snapshots)} points "
          f"(max avg={snapshots[-1]['Avg_Coverage_%']:.2f}% "
          f"with {snapshots[-1]['n_ES']} ES)")

    # ── Part 2: 100 combinations ──────────────────────────────
    print(f"\n[5/5] Part 2 — {N_COMBINATIONS} anchored combos + pruning ...")
    anchor_pool = all_rows[:N_COMBINATIONS]
    combinations = []
    seen          = set()

    for i, anchor in enumerate(anchor_pool):
        t_i = time.time()
        print(f"  [{i+1:>3}/{N_COMBINATIONS}]  {anchor['label']} "
              f"({len(anchor['g_set'])}G) ...", end=' ', flush=True)

        sel    = greedy_from_anchor(anchor, all_rows, tail_g_map)
        pruned = backward_prune(sel, tail_g_map)

        key = frozenset(r['label'] for r in pruned)
        if key in seen:
            print("duplicate")
            continue
        seen.add(key)

        avg, n_rois, roi_pct = compute_unique_coverage(pruned, tail_g_map)
        n_es = len(pruned)
        eff  = round(avg / n_es, 4) if n_es else 0

        print(f"{n_es} ES | ROIs={n_rois}/{len(tail_g_map)} | "
              f"avg={avg:.2f}% | eff={eff}  ({time.time()-t_i:.1f}s)")

        combinations.append({
            'com_label' : '',
            'anchor'    : anchor['label'],
            'selected'  : pruned,
            'avg_cov'   : avg,
            'n_rois'    : n_rois,
            'n_es'      : n_es,
            'eff'       : eff,
            'roi_pct'   : roi_pct,
        })

    if not combinations:
        print("No valid combinations found.")
        return

    # Sort: avg_cov DESC → n_es ASC
    combinations.sort(key=lambda c: (-c['avg_cov'], c['n_es']))
    for rank, c in enumerate(combinations, 1):
        c['com_label'] = f"ES_COM_{rank}"

    # Summary
    df_summary = pd.DataFrame([{
        'ES_COM'             : c['com_label'],
        'Anchor_ES'          : c['anchor'],
        'ES_Count'           : c['n_es'],
        'ROIs_Covered'       : c['n_rois'],
        'ROIs_Total'         : len(tail_g_map),
        'Avg_Unique_G_Cov_%' : round(c['avg_cov'], 2),
        'Efficiency_avg/n'   : c['eff'],
        'ES_Labels'          : ', '.join(r['label'] for r in c['selected']),
    } for c in combinations])

    best = combinations[0]
    best_selected = best['selected']

    # Per-ROI coverage
    roi_rows = []
    for roi_id in sorted(tail_g_map):
        total_g = len(tail_g_map[roi_id])
        pct     = best['roi_pct'][roi_id]
        unique  = set()
        for r in best_selected:
            unique.update(r['g_set'] & tail_g_map[roi_id])
        contrib = [r['label'] for r in best_selected
                   if r['g_set'] & tail_g_map[roi_id]]
        roi_rows.append({
            'ROI'                    : roi_id,
            'Total_G_Values'         : total_g,
            'Unique_G_Covered'       : len(unique),
            'Coverage_%'             : round(pct, 2),
            'Contributing_ES_Count'  : len(contrib),
            'Contributing_ES_Labels' : ', '.join(contrib),
        })
    df_roi = (pd.DataFrame(roi_rows)
                .sort_values('Coverage_%', ascending=False))

    # ROI x ES matrix (conflict-free)
    print("  Building ROI x ES matrix ...")
    matrix_rows = []
    for roi_id, g_keys in sorted(tail_g_map.items()):
        asgn = assign_conflict_free(roi_id, best_selected, g_keys)
        row  = {'ROI': roi_id,
                'Total_G': len(g_keys),
                'Coverage_%': round(best['roi_pct'][roi_id], 2)}
        for r in best_selected:
            row[r['label']] = asgn.get(r['label'], '')
        matrix_rows.append(row)
    df_matrix = pd.DataFrame(matrix_rows)
    meta = ['ROI', 'Total_G', 'Coverage_%']
    df_matrix = df_matrix[meta + [r['label'] for r in best_selected]].fillna('')

    # Export
    os.makedirs('Docs', exist_ok=True)
    ts  = time.strftime("%Y%m%d_%H%M%S")
    out = f'Docs/Optimal_ES_Combination_{ts}.xlsx'

    with pd.ExcelWriter(out, engine='openpyxl') as writer:
        df_curve.to_excel(writer,   sheet_name='Coverage_Curve',      index=False)
        df_summary.to_excel(writer, sheet_name='Summary_100_Combos',  index=False)
        df_matrix.to_excel(writer,  sheet_name=f'{best["com_label"]}_Details', index=False)
        df_roi.to_excel(writer,     sheet_name='Per_ROI_Coverage',     index=False)

    elapsed = time.time() - t0
    print(f"\n  Saved: {out}  ({elapsed:.0f}s)")

    print("\n" + "=" * 70)
    print("TOP 10 COMBINATIONS  (sorted by avg_coverage DESC, n_ES ASC)")
    print("=" * 70)
    print(f"  {'COM':<14} {'ES':>4} {'ROIs':>6} {'Avg%':>7} {'Eff':>8}")
    print("  " + "-" * 48)
    for c in combinations[:10]:
        print(f"  {c['com_label']:<14} {c['n_es']:>4} "
              f"{c['n_rois']:>4}/{len(tail_g_map)} "
              f"{c['avg_cov']:>6.2f}% {c['eff']:>8.4f}")
    print("=" * 70)
    print(f"\nCoverage curve elbow: check 'Coverage_Curve' sheet")
    print(f"Look for the row where 'Delta_Avg_%' drops below 1%")


if __name__ == '__main__':
    main()
