"""
OptimalESCombination.py
-----------------------
מוצא קומבינציה אופטימלית של ES sequences ש:
1. מכסה את כל 120 הסקשנים/ROIs
2. משתמש במינימום ES שונים
3. ממקסם את הממוצע של כיסוי וקטורי G עבור כל סקשן

אלגוריתם:
- Greedy עם עדיפות לרצפים ארוכים
- חישוב כיסוי ממוצע עבור כל קומבינציה
- הקצאה ייחודית של G-values (ללא כפילויות בין ES עבור אותו ROI)

פלט:
- Excel עם סיכום קומבינציות + פירוט הקצאות (תא אחד = G אחד בלבד)
"""

import pandas as pd
import numpy as np
import re
import os
import glob
import time
from collections import defaultdict
from typing import Dict, Set, List, Tuple, FrozenSet

# ═════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ═════════════════════════════════════════════════════════════════════

TAILS_FOLDER = 'All_Tail_Proteins_Tables'
SEQUENCES_FILE = 'Docs/Exact_Sequences_Miner_20260630_090923.csv'
MAX_COMBINATIONS = 100  # מספר קומבינציות לבדוק
OUTPUT_FOLDER = 'Docs'

# ═════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ═════════════════════════════════════════════════════════════════════

def load_tail_tables(tails_folder: str) -> Dict[str, Set[str]]:
    """
    טוען את כל טבלאות ה-Tail ומחזיר מיפוי:
    roi_id → set of G-keys (e.g., "P1_ROI_03_G4")
    """
    print(f"טוען טבלאות Tail מ-{tails_folder}...")
    files = sorted(
        glob.glob(os.path.join(tails_folder, '*.xlsx')) +
        glob.glob(os.path.join(tails_folder, '*.csv'))
    )
    
    tail_g_map = {}
    for f in files:
        if os.path.basename(f).startswith('~$'):
            continue
            
        roi_id = (os.path.basename(f)
                  .replace('_tail_proteins.xlsx', '')
                  .replace('_tail_proteins.csv', ''))
        
        try:
            df = pd.read_excel(f) if f.endswith('.xlsx') else pd.read_csv(f)
        except Exception as e:
            print(f"  אזהרה: לא ניתן לקרוא {f}: {e}")
            continue
        
        # מצא עמודות G
        g_nums = sorted(set(
            int(re.search(r'G(\d+)', c).group(1))
            for c in df.columns
            if re.match(r'G\d+[+-]$', c)
        ))
        
        g_keys = set(f"{roi_id}_G{n}" for n in g_nums)
        if g_keys:
            tail_g_map[roi_id] = g_keys
    
    print(f"  נטענו {len(tail_g_map)} ROIs עם סה\"כ {sum(len(v) for v in tail_g_map.values())} G-values")
    return tail_g_map


def parse_participating_rois(roi_string: str) -> Set[str]:
    """
    מנתח 'Participating ROIs' לסט של G-keys
    דוגמה: 'P1_ROI_03_G4 (+) | P9_ROI_51_G13 (-)' → {'P1_ROI_03_G4', 'P9_ROI_51_G13'}
    """
    if pd.isna(roi_string):
        return set()
    
    g_keys = set()
    for part in str(roi_string).split(' | '):
        # הסר סימנים (+)/(-) ורווחים
        key = re.sub(r'\s*\([+-]\)\s*$', '', part).strip()
        if key:
            g_keys.add(key)
    
    return g_keys


def extract_roi_from_g_key(g_key: str) -> str:
    """
    'P1_ROI_03_G4' → 'P1_ROI_03'
    """
    match = re.match(r'(.+?)_G\d+', g_key)
    return match.group(1) if match else None


def load_es_sequences(sequences_file: str) -> pd.DataFrame:
    """
    טוען את קובץ ה-ES sequences ומוסיף עמודות עזר
    """
    print(f"טוען ES sequences מ-{sequences_file}...")
    
    try:
        df = pd.read_csv(sequences_file)
        if len(df.columns) == 1 and '\t' in df.columns[0]:
            df = pd.read_csv(sequences_file, sep='\t')
    except Exception:
        df = pd.read_excel(sequences_file)
    
    # הוסף תוויות ES
    df['ES_Label'] = [f"ES{i+1}" for i in range(len(df))]
    
    # נתח Participating ROIs לסטים
    df['G_Keys'] = df['Participating ROIs'].apply(parse_participating_rois)
    df['G_Count'] = df['G_Keys'].apply(len)
    
    # סנן רק ES עם G-values
    df = df[df['G_Count'] > 0].copy()
    df.reset_index(drop=True, inplace=True)
    
    print(f"  נטענו {len(df):,} ES sequences (עם G-values)")
    return df


def build_es_roi_mapping(df_es: pd.DataFrame, tail_g_map: Dict[str, Set[str]]) -> Tuple[
    Dict[Tuple[str, str], Set[str]],  # (roi, es_label) → set of G-keys
    Dict[str, Set[str]]  # es_label → set of all G-keys
]:
    """
    בונה מיפויים:
    1. (roi_id, es_label) → G-keys שה-ES מכסה ב-ROI זה
    2. es_label → כל ה-G-keys שה-ES מכסה בכל ה-ROIs
    """
    print("בונה מיפוי ES ← ROI...")
    
    es_roi_g_map = defaultdict(set)  # (roi, es) → G-keys
    es_all_g_map = {}  # es → all G-keys
    
    for idx, row in df_es.iterrows():
        es_label = row['ES_Label']
        g_keys = row['G_Keys']
        es_all_g_map[es_label] = g_keys
        
        # חלק לפי ROI
        for g_key in g_keys:
            roi_id = extract_roi_from_g_key(g_key)
            if roi_id and roi_id in tail_g_map:
                # ודא שזה G-value תקף מטבלת ה-tail
                if g_key in tail_g_map[roi_id]:
                    es_roi_g_map[(roi_id, es_label)].add(g_key)
    
    print(f"  מיפוי הושלם: {len(es_all_g_map)} ES sequences")
    return dict(es_roi_g_map), es_all_g_map


# ═════════════════════════════════════════════════════════════════════
# GREEDY COMBINATION GENERATION
# ═════════════════════════════════════════════════════════════════════

def calculate_combination_coverage(
    combo: List[str],
    tail_g_map: Dict[str, Set[str]],
    es_roi_g_map: Dict[Tuple[str, str], Set[str]]
) -> Tuple[float, int, Dict[str, float]]:
    """
    מחשב את הכיסוי עבור קומבינציה נתונה של ES:
    
    Returns:
        avg_coverage: ממוצע % כיסוי על כל ה-ROIs
        rois_covered: מספר ROIs עם כיסוי > 0
        roi_coverage_pct: dict של roi_id → % כיסוי
    """
    roi_coverage_pct = {}
    rois_covered = 0
    
    for roi_id, total_g_keys in tail_g_map.items():
        # מצא את כל ה-G-keys שהקומבינציה מכסה ב-ROI זה
        covered_g = set()
        for es_label in combo:
            covered_g.update(es_roi_g_map.get((roi_id, es_label), set()))
        
        # חשב % כיסוי
        n_covered = len(covered_g)
        n_total = len(total_g_keys)
        pct = (n_covered / n_total * 100) if n_total > 0 else 0.0
        roi_coverage_pct[roi_id] = pct
        
        if pct > 0:
            rois_covered += 1
    
    # ממוצע כיסוי
    avg_coverage = np.mean(list(roi_coverage_pct.values()))
    
    return avg_coverage, rois_covered, roi_coverage_pct


def greedy_es_combination(
    df_es: pd.DataFrame,
    tail_g_map: Dict[str, Set[str]],
    es_roi_g_map: Dict[Tuple[str, str], Set[str]],
    es_all_g_map: Dict[str, Set[str]],
    max_combinations: int
) -> List[Dict]:
    """
    יוצר קומבינציות ES באמצעות אלגוריתם Greedy:
    
    1. בוחר ES התחלתי (לפי אורך רצף)
    2. מוסיף ES נוספים שמשפרים את הכיסוי הממוצע
    3. ממשיך עד כיסוי מקסימלי של ROIs
    
    Returns:
        רשימה של dictionaries עם פרטי כל קומבינציה
    """
    print(f"\nיוצר {max_combinations} קומבינציות אופטימליות...")
    
    # מיין ES לפי אורך רצף (DESC) ואז כמות G-values (DESC)
    df_sorted = df_es.sort_values(
        by=['Sequence Length', 'G_Count'],
        ascending=[False, False]
    ).reset_index(drop=True)
    
    all_rois = set(tail_g_map.keys())
    total_rois = len(all_rois)
    
    # בנה מיפוי ES → ROIs מכוסים (מראש)
    print("  בונה מיפוי ES → ROIs...")
    es_to_rois = {}
    for es_label in es_all_g_map.keys():
        rois_covered = set()
        for roi_id in all_rois:
            if es_roi_g_map.get((roi_id, es_label)):
                rois_covered.add(roi_id)
        if rois_covered:
            es_to_rois[es_label] = rois_covered
    
    combinations = []
    
    for combo_idx in range(min(max_combinations, len(df_sorted))):
        # התחל עם ES התחלתי
        anchor_es = df_sorted.iloc[combo_idx]['ES_Label']
        combo = [anchor_es]
        
        # ROIs שכבר נכללו
        covered_rois = es_to_rois.get(anchor_es, set()).copy()
        
        # Greedy: הוסף ES שמשפר את הכיסוי הממוצע
        iteration = 0
        max_iterations = 200  # הגבל איטרציות
        
        while len(covered_rois) < total_rois and iteration < max_iterations:
            iteration += 1
            
            best_es = None
            best_improvement = 0
            best_new_rois = set()
            
            # נסה כל ES שעדיין לא נבחר (רק מתוך הממוינים)
            for idx, row in df_sorted.iterrows():
                es_label = row['ES_Label']
                if es_label in combo:
                    continue
                
                # ROIs חדשים שה-ES הזה יוסיף
                es_rois = es_to_rois.get(es_label, set())
                new_rois = es_rois - covered_rois
                
                if not new_rois:
                    continue
                
                # העדף ES עם: 1) ROIs חדשים 2) רצף ארוך
                improvement = len(new_rois) * 1000 + row['Sequence Length']
                
                if improvement > best_improvement:
                    best_improvement = improvement
                    best_es = es_label
                    best_new_rois = new_rois
            
            if best_es is None or len(best_new_rois) == 0:
                # אין שיפור נוסף
                break
            
            combo.append(best_es)
            covered_rois.update(best_new_rois)
        
        # חשב מדדים לקומבינציה זו
        avg_cov, rois_cov, roi_cov_pct = calculate_combination_coverage(
            combo, tail_g_map, es_roi_g_map
        )
        
        combinations.append({
            'ES_COM': f"ES_COM_{combo_idx + 1}",
            'ES_Count': len(combo),
            'ROIs_Covered': rois_cov,
            'ROIs_Total': total_rois,
            'Avg_Section_G_Coverage_Pct': round(avg_cov, 2),
            'combo_list': combo,
            'roi_coverage_pct': roi_cov_pct
        })
        
        if (combo_idx + 1) % 10 == 0:
            print(f"  הושלמו {combo_idx + 1}/{max_combinations} קומבינציות...")
    
    return combinations


# ═════════════════════════════════════════════════════════════════════
# CONFLICT-FREE G ASSIGNMENT
# ═════════════════════════════════════════════════════════════════════

def assign_conflict_free_gs(
    roi_id: str,
    es_list: List[str],
    es_roi_g_map: Dict[Tuple[str, str], Set[str]],
    df_es: pd.DataFrame
) -> Dict[str, str]:
    """
    מקצה בדיוק G-value אחד לכל ES עבור ROI נתון.
    מבטיח שאין כפילויות של G-values בין ES שונים עבור אותו ROI.
    
    אסטרטגיה:
    1. מיין ES לפי מספר אופציות (מועדף למי שיש לו פחות אופציות)
    2. בחר G-value ראשון שעדיין לא הוקצה
    
    Returns:
        dict: es_label → "G<num>(+/-)"
    """
    assigned_gs = set()  # G-keys שכבר הוקצו
    assignments = {}
    
    # איסוף מועמדים לכל ES
    es_candidates = []
    for es in es_list:
        g_keys = es_roi_g_map.get((roi_id, es), set())
        if g_keys:
            es_candidates.append((es, list(g_keys)))
    
    # מיין לפי מספר מועמדים (פחות מועמדים = עדיפות גבוהה)
    es_candidates.sort(key=lambda x: len(x[1]))
    
    # הקצה G-value לכל ES
    for es, g_keys in es_candidates:
        chosen_g = None
        for g_key in g_keys:
            if g_key not in assigned_gs:
                chosen_g = g_key
                break
        
        if chosen_g:
            assigned_gs.add(chosen_g)
            # המר ל-format "G<num>(+/-)"
            # צריך למצוא את הסימן מה-Participating ROIs המקורי
            g_num = re.search(r'_G(\d+)$', chosen_g).group(1)
            
            # מצא את הסימן מתוך ה-dataframe המקורי
            es_row = df_es[df_es['ES_Label'] == es].iloc[0]
            roi_str = es_row['Participating ROIs']
            sign_match = re.search(rf'{re.escape(chosen_g)}\s*\(([+-])\)', roi_str)
            sign = sign_match.group(1) if sign_match else '?'
            
            assignments[es] = f"G{g_num}({sign})"
    
    return assignments


# ═════════════════════════════════════════════════════════════════════
# MAIN PIPELINE
# ═════════════════════════════════════════════════════════════════════

def main():
    start_time = time.time()
    
    print("=" * 80)
    print("אלגוריתם אופטימיזציה לקומבינציות ES")
    print("=" * 80)
    
    # ─── שלב 1: טעינת נתונים ───
    tail_g_map = load_tail_tables(TAILS_FOLDER)
    df_es = load_es_sequences(SEQUENCES_FILE)
    es_roi_g_map, es_all_g_map = build_es_roi_mapping(df_es, tail_g_map)
    
    total_rois = len(tail_g_map)
    total_g_values = sum(len(v) for v in tail_g_map.values())
    print(f"\nסיכום נתונים:")
    print(f"  סה\"כ ROIs: {total_rois}")
    print(f"  סה\"כ G-values: {total_g_values:,}")
    print(f"  סה\"כ ES sequences: {len(df_es):,}")
    
    # ─── שלב 2: יצירת קומבינציות ───
    combinations = greedy_es_combination(
        df_es, tail_g_map, es_roi_g_map, es_all_g_map, MAX_COMBINATIONS
    )
    
    # ─── שלב 3: מיון קומבינציות ───
    print("\nממיין קומבינציות לפי:")
    print("  1. מספר ROIs מכוסים (יותר = טוב)")
    print("  2. ממוצע כיסוי G (גבוה = טוב)")
    print("  3. מספר ES (פחות = טוב)")
    
    df_combos = pd.DataFrame(combinations).sort_values(
        by=['ROIs_Covered', 'Avg_Section_G_Coverage_Pct', 'ES_Count'],
        ascending=[False, False, True]
    ).reset_index(drop=True)
    
    # עדכן מספור ES_COM
    for idx in range(len(df_combos)):
        df_combos.at[idx, 'ES_COM'] = f"ES_COM_{idx + 1}"
    
    # ─── שלב 4: בניית טבלת פירוט ל-ES_COM_1 (הטוב ביותר) ───
    print("\nבונה טבלת הקצאות לקומבינציה הטובה ביותר...")
    
    best_combo = df_combos.iloc[0]
    best_combo_list = best_combo['combo_list']
    best_combo_name = best_combo['ES_COM']
    
    print(f"  {best_combo_name}:")
    print(f"    ES sequences: {len(best_combo_list)}")
    print(f"    ROIs covered: {best_combo['ROIs_Covered']}/{total_rois}")
    print(f"    ממוצע כיסוי G: {best_combo['Avg_Section_G_Coverage_Pct']:.2f}%")

    # ─── בניית טבלת Optimal_Combination עם ממוצע כיסוי G לכל ES ───
    print("\nמחשב ממוצע כיסוי G לכל ES בקומבינציה...")
    
    optimal_combo_rows = []
    for es_label in best_combo_list:
        es_row = df_es[df_es['ES_Label'] == es_label].iloc[0]
        
        # מצא כל ה-ROIs שה-ES הזה מופיע בהם
        rois_with_es = [
            roi_id for roi_id in tail_g_map
            if es_roi_g_map.get((roi_id, es_label))
        ]
        
        # חשב % כיסוי G לכל ROI שה-ES מופיע בו
        per_roi_coverage = []
        for roi_id in rois_with_es:
            covered_g = es_roi_g_map.get((roi_id, es_label), set())
            total_g   = len(tail_g_map[roi_id])
            pct = len(covered_g) / total_g * 100 if total_g > 0 else 0.0
            per_roi_coverage.append(pct)
        
        avg_g_cov_per_section = round(np.mean(per_roi_coverage), 2) if per_roi_coverage else 0.0
        
        optimal_combo_rows.append({
            'ES_Label'                       : es_label,
            'Sequence_Length'                : int(es_row['Sequence Length']),
            'Unique_Patients_Count'          : es_row['Unique Patients Count'],
            'Total_G_Values_in_ES'           : int(es_row['G_Count']),
            'ROIs_ES_Appears_In'             : len(rois_with_es),
            'Avg_G_Coverage_%_Per_Section'   : avg_g_cov_per_section,
            'Exact_Sign_Sequence'            : es_row['Exact Sign Sequence'],
        })
    
    df_optimal_combo = pd.DataFrame(optimal_combo_rows)
    
    # בנה מטריצה עם הקצאות ייחודיות
    matrix_data = {}
    
    for roi_id in sorted(tail_g_map.keys()):
        row_assignments = assign_conflict_free_gs(
            roi_id, best_combo_list, es_roi_g_map, df_es
        )
        
        matrix_data[roi_id] = {es: row_assignments.get(es, "") for es in best_combo_list}
    
    df_matrix = pd.DataFrame.from_dict(matrix_data, orient='index')
    df_matrix = df_matrix.reindex(columns=best_combo_list).fillna("")
    df_matrix.index.name = 'ROI'
    df_matrix.reset_index(inplace=True)
    
    # ─── שלב 5: שמירה ───
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    timestamp = time.strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(OUTPUT_FOLDER, f'Optimal_ES_Combination_{timestamp}.xlsx')
    
    print(f"\nשומר תוצאות ל-{output_file}...")
    
    # הכן טבלת סיכום (ללא combo_list ו-roi_coverage_pct)
    table_summary = df_combos.drop(columns=['combo_list', 'roi_coverage_pct']).copy()
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        table_summary.to_excel(writer, sheet_name='Combinations_Summary', index=False)
        df_optimal_combo.to_excel(writer, sheet_name='Optimal_Combination', index=False)
        df_matrix.to_excel(writer, sheet_name=f'{best_combo_name}_Details', index=False)
    
    elapsed = time.time() - start_time
    print(f"\nהושלם בהצלחה!")
    print(f"  זמן ריצה: {elapsed:.2f} שניות")
    print(f"  קובץ פלט: {output_file}")
    print("=" * 80)


if __name__ == "__main__":
    main()
