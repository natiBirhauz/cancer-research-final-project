# פרויקט גמר: ניתוח Surprisal Analysis, רצפי סימנים מדויקים (ES) ושחזור R²

מחקר בסביבת המיקרו-סביבה הגידולית (TME) בשיטת Imaging Mass Cytometry (IMC).
האלגוריתם מוצא קבוצה מינימלית ואופטימלית של **רצפי סימנים מדויקים (Exact Sign Sequences - ES)** המשותפים ל-120 סקשנים רקמתיים (ROIs) מ-19 חולים, המכסים באופן מקסימלי את כלל וקטורי ה-$G$, תוך שימור סימנים מוחלט (`+`/`-`), אפס כפילויות, ושחזור נאמן של המידע המקורי לפי מדד **$R^2$**.

---

## מבנה התיקייה לאחר ארגון

```text
פרויקט גמר/
├── 01_extract_tail_tables.py        # שלב 1: חילוץ וריכוז טבלאות ה-TAILS מתיקיות החולים
├── 02_cache_r2_data.py              # שלב 2: חישוב מקבילי של מטריצות SVD ושונות (Docs/r2_cache.pkl)
├── 03_mine_all_es_combinations.py   # שלב 3: כריית כל 120,791 הצירופים מאורך 3 עד 8
├── 04_compare_anchor_strategies.py  # שלב 4: השוואת 6 אסטרטגיות עוגן ובחירת המודל המנצח (Hybrid)
├── 05_run_winning_pipeline.py       # שלב 5: הפקת הדוחות הסופיים (Optimal_ES_Result_Signs_R2.xlsx)
├── run_pipeline.py                  # סקריפט ראשי להרצת כל 5 השלבים מקצה לקצה
│
├── All_Tail_Proteins_Tables/        # 120 קבצי TAILS עם סימני חלבונים מדויקים (+/-)
├── all patients (without isotope)/   # תיקיות המקור הגולמיות של כל חולה (X_raw, G, L)
│
├── Docs/                            # תיקיית התוצרים והדוחות הפעילים
│   ├── Algorithm_Full_Summary.html  # דוח אינטראקטיבי מלא ומעוצב
│   ├── Algorithm_Full_Summary.md    # תיעוד אלגוריתמי מלא ב-Markdown
│   ├── Optimal_ES_Result_Signs_R2.xlsx # קובץ האקסל המרכזי (8 גיליונות מלאים)
│   ├── Anchor_Strategies_Comparison.xlsx # השוואת 6 אסטרטגיות עוגן + פירוט 32 המודולים
│   ├── All_Possible_ES_Combinations_From_Length_3.xlsx # כל 120,791 הצירופים (אקסל)
│   ├── All_Possible_ES_Combinations_From_Length_3.csv  # כל 120,791 הצירופים (CSV)
│   ├── R2 comparison per ROI.xlsx   # דוח R² נפרד לפי המדריך
│   ├── winning_hybrid_32_es_detailed.csv # טבלת 32 מודולי ה-ES המנצחים
│   ├── r2_cache.pkl                 # קובץ זיכרון מטמון למטריצות SVD
│   └── anchor_comparison_results.pkl# תוצאות השוואת העוגנים
│
├── Archive_Old_Project/             # ארכיון מסודר של כל הקבצים הישנים שאינם בשימוש
│   ├── Old_Scripts/                 # 40 סקריפטים ישנים מגרסאות קודמות
│   ├── Old_Docs/                    # 109 דוחות והרצות היסטוריות ישנות
│   ├── Old_Directories/             # All_G_Tables, Tails, FinalProject, scratch
│   └── Old_Root_Files/              # קבצי עזר וטיוטות ישנות
│
└── R2 Analysis Guide.html           # מדריך מתודולוגי לחישובי R²
```

---

## שלבי הצינור החישובי (Pipeline)

ניתן להריץ את הצינור כולו בפקודה אחת:
```bash
python run_pipeline.py
```
או להריץ כל שלב בנפרד:

### שלב 1: חילוץ טבלאות ה-TAILS
```bash
python 01_extract_tail_tables.py
```
מרכז את כל 120 קבצי `tail_proteins.xlsx` מתיקיות החולים לתוך `All_Tail_Proteins_Tables/` בפורמט שמי קנוני (`P{X}_ROI_{Y}_tail_proteins.xlsx`).

### שלב 2: חישוב מטריצות SVD ומשקולות שונות (Cache)
```bash
python 02_cache_r2_data.py
```
מבצע חישוב מקבילי של מטריצות הביטוי, מקדמי התאים $L$, וקטורי $G$, ושחזור $R^2_{full}$ עבור כל 120 הסקשנים. שומר את `Docs/r2_cache.pkl`.

### שלב 3: כריית כלל הצירופים האפשריים החל מאורך 3
```bash
python 03_mine_all_es_combinations.py
```
מחלץ מתוך 2,778 וקטורי ה-$G$ את כל תתי-הצירופים הקנוניים מאורך 3 עד 8 (120,791 צירופים ייחודיים), ומסווג צירופים שכיחים (המופיעים ב-2 וקטורים ומעלה). מייצא ל-CSV ולאקסל.

### שלב 4: השוואת 6 אסטרטגיות עוגן
```bash
python 04_compare_anchor_strategies.py
```
מריץ אלגוריתם Greedy Set Cover עם 6 נקודות עוגן שונות (שונות SVD, שכיחות גלובלית, ספציפיות עמוקה 7-8, חדירה חיסונית, סטרומה, ומודל היברידי). שומר את `Docs/Anchor_Strategies_Comparison.xlsx`.

### שלב 5: הרצת המודל המנצח והפקת הדוחות המלאים
```bash
python 05_run_winning_pipeline.py
```
מריץ את האסטרטגיה ההיברידית (32 מודולים), בונה מטריצת הקצאה נקייה ללא קונפליקטים (0 כפילויות), מחשב $R^2$ לכל סקשן ומייצא את `Docs/Optimal_ES_Result_Signs_R2.xlsx`.

---

## תוצאות עיקריות (האסטרטגיה ההיברידית המנצחת)

* **32 מודולי ES בלבד** (צמצום משמעותי מ-88 מודולים בדוחות ישנים).
* **כיסוי וקטורי $G$:** **2,776 מתוך 2,778 (99.93%)**.
* **הסבר שונות ביולוגית:** **87.61%**.
* **איכות שחזור $R^2$:** ממוצע $R^2_{optimal} = \mathbf{0.8838}$, מחיר קומפרסיה $\Delta R^2 = \mathbf{0.0784}$ בלבד.
* **92.5% מהסקשנים** ברמת שחזור גבוהה (מצוין או סביר).
* **אפס כפילויות ואפס סתירות סימנים** בכל 120 הסקשנים.
