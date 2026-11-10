# מדריך אלגוריתמי ומתמטי מקיף: פלטפורמת Surprisal Analysis (SA) לזיהוי סט מינימלי של חתימות (ES)
# Comprehensive Algorithmic & Mathematical Guide: Surprisal Analysis (SA) – Minimal Exact Sign Signatures (ES) Platform

מסמך זה מציג הסבר מלא ומפורט **בעברית ובאנגלית** על המתודולוגיה המתמטית, שלבי האלגוריתם, ההתמודדות עם דאטה-סטים בכל סדר גודל (מפאנלים ממוקדים של עשרות חלבונים ועד דאטה-סטים של 20,000+ גנים), ומבנה קבצי הפלט של הפלטפורמה.

---

## חלק א': מדריך מלא בעברית (Hebrew Guide)

### 1. מטרת הפלטפורמה והאתגר החישובי
הפלטפורמה מקבלת **כל אוסף של דאטה-סטים שעברו אנליזת הפתעה (Surprisal Analysis - SA)** וחולקים ביניהם חלבונים או גנים משותפים (למשל פרוטאומיקה, טרנסקריפטומיקה / RNA-Seq, דאטה-סטים קליניים כגון TCGA, או מדידות ברמת התא הבודד).

בכל דאטה-סט שעבר Surprisal Analysis מתקבלים וקטורי תהליכים ביולוגיים בלתי-תלויים ($G_1, G_2, \dots, G_K$). בכל וקטור $G_k$, הגנים או החלבונים בעלי המשקל המשמעותי ביותר ("הזנבות" – Tails) נושאים כיוון פעילות מוגדר:
* **סימן חיובי (`+` / $G+$):** גנים או חלבונים בביטוי-יתר (Upregulated) בתהליך.
* **סימן שלילי (`-` / $G-$):** גנים או חלבונים בדיכוי (Downregulated) בתהליך.

**המטרה המרכזית:**  
למצוא **קבוצה מינימלית ונקייה מסתירות של חתימות סימנים מדויקות (Exact Sign Signatures - ES)** שחוזרת על עצמה על פני הדאטה-סטים השונים, כך שבאמצעות מספר קטן של חתימות משותפות נוכל לתאר, לכסות ולשחזר (לפי מדד $R^2$) את כלל וקטורי ה-$G$ של כל הדאטה-סטים.

---

### 2. פירוק Surprisal Analysis ומשקלי השונות $w(g)$
עבור מטריצת ביטוי $X$ (בגודל $N_{\text{features}} \times N_{\text{samples/cells}}$) בכל דאטה-סט, הפירוק הלוגריתמי לפי Surprisal Analysis מוגדר כ:
$$\ln(X_{i,j}) = \sum_{\alpha=0}^K G_{i,\alpha} \cdot L_{j,\alpha} = \left( G \times L^T \right)_{i,j}$$

* **מצב שיווי המשקל ($\alpha = 0$, וקטורי $G_0, L_0$):** המצב הבסיסי היציב המשותף לדגימות/תאים.
* **תהליכי הסטייה ($\alpha \ge 1$, וקטורי $G_\alpha, L_\alpha$):** אילוצים ותהליכים ביולוגיים פעילים המטים את המערכת משיווי משקל.
* **משקל השונות של כל וקטור $w(g)$:** תרומתו היחסית של וקטור $G_k$ לסך אנרגיית הסטייה בדאטה-סט מחושבת מתוך אמפליטודות ה-$L$:
  $$w(g_k) = \frac{\|L_k\|^2}{\sum_{\alpha \ge 1} \|L_\alpha\|^2}, \quad \|L_k\|^2 = \sum_j L_{j,k}^2$$

---

### 3. חילוץ זנבות ($G+ / G-$) ואינווריאנטיות להיפוך סימן ($ES \equiv -ES$)
1. **סף חיתוך (`Cutoff Loading |G| >= threshold`):** בכל וקטור $G_k$, רק גנים/חלבונים שערכם המוחלט עולה על הסף (למשל $|G_{i,k}| \ge 0.05$) נכללים בזנב הפעיל ומקבלים סימן `+` או `-`.
2. **אינווריאנטיות להיפוך סימן ($ES \equiv -ES$):** בפירוק וקטורים עצמיים (SVD), מתקיים השוויון המתמטי:
   $$G_k \times L_k^T = (-G_k) \times (-L_k)^T$$
   כלומר, חתימה שבה `GeneA(+) | GeneB(-)` שקולה לחלוטין לחתימה ההפוכה `GeneA(-) | GeneB(+)` (ההבדל מתבטא רק בסימן האמפליטודה $L_k$). לכן, האלגוריתם מנרמל כל חתימה ל**צורה קנונית אחידה (Canonical Form)**: הגנים ממוינים אלפביתית, ואם הגן הראשון נושא סימן `-`, כל הסימנים בחתימה מתהפכים. פעולה זו מונעת כפילויות מיותרות.

---

### 4. כלל אפס הסתירות (Zero-Contradiction Matching Rule)
כאשר בודקים האם חתימה מועמדת $\text{ES}$ מכסה וקטור $G$:
1. **חפיפה מינימלית:** חייבת להיות חפיפה של לפחות `min_overlap` גנים/חלבונים משותפים בין ה-$\text{ES}$ לבין וקטור ה-$G$.
2. **התאמה ישרה (`+` Direct Match):** לכל הגנים המשותפים יש **בדיוק אותו סימן** ב-$\text{ES}$ ובווקטור ה-$G$.
3. **התאמה הפוכה (`-` Inverse Match):** לכל הגנים המשותפים יש **סימן הפוך בדיוק** ב-$\text{ES}$ לעומת וקטור ה-$G$.
4. **פסילת סתירות:** אם אפילו גן משותף אחד נושא סימן זהה וגן משותף אחר נושא סימן הפוך – זוהי **סתירה רגולטורית**, והחתימה נפסלת מיד עבור אותו וקטור $G$.

---

### 5. מנוע כרייה דו-מצבי: התמודדות עם דאטה-סטים של 20,000+ גנים ואורך חתימה 1,000+
אחת השאלות המרכזיות בניתוח נתונים אומיים רחבי-היקף היא: **כיצד האלגוריתם מתמודד עם דאטה-סטים בעלי 20,000 גנים שבהם `Min Sequence Length = 1000`?**

* **בפאנלים קטנים (עד ~14 חלבונים בזנב ואורך חתימה עד 10):** האלגוריתם מפעיל **מנייה קומבינטורית מדויקת** של כל תתי-הקבוצות $\binom{|G|}{k}$.
* **בדאטה-סטים גדולים (מאות עד 20,000+ גנים, ואורכי חתימה של עשרות, מאות או 1,000+):** מנייה נאיבית של $\binom{2000}{1000}$ אינה אפשרית חישובית. לכן הפלטפורמה מפעילה אוטומטית את **מנוע החיתוכים עקבי-הסימן (Scalable Sign-Consistent Intersection Mining)**:
  * מתמטית, כל חתימת $\text{ES}$ החוזרת בלפחות שני וקטורי $G$ ($G_a$ ו-$G_b$) חייבת להיות מוכלת ב**חיתוך עקבי-הסימן** ביניהם:
    $$I_+(G_a, G_b) = \{p \in G_a \cap G_b \mid \text{sign}_a(p) = \text{sign}_b(p)\}$$
    $$I_-(G_a, G_b) = \{p \in G_a \cap G_b \mid \text{sign}_a(p) = -\text{sign}_b(p)\}$$
  * מכיוון שמספר וקטורי ה-$G$ הוא מסדר גודל של עשרות עד אלפים בודדים, חישוב החיתוכים הישרים וההפוכים בין זוגות וקטורים מתבצע בסיבוכיות פולינומית מהירה ($O(N_G^2 \cdot |G|)$) ומאתר תוך שניות בודדות חתימות משותפות באורך 100, 500 או 1,000+ גנים!

---

### 6. אלגוריתם ה-Greedy Set Cover ופונקציית המטרה הרב-מטרתית
כדי לבחור את הסט המינימלי והאיכותי ביותר של חתימות $\text{ES}$, האלגוריתם פותר את בעיית הכיסוי (Set Cover) באופן חמדן איטרטיבי. בכל צעד נבחרת החתימה הממקסמת את הציון הרב-מטרתי (Hybrid Multi-Objective):
$$\text{Score}(\text{ES}) = \left[ \sum_{g \in \text{New}(\text{ES})} w(g) \right] \times \sqrt{|\text{New}(\text{ES})|}$$
* האיבר $\sum w(g)$ מבטיח תעדוף של וקטורי $G$ הנושאים שונות ואנרגיה גבוהה בדאטה-סטים.
* האיבר $\sqrt{|\text{New}(\text{ES})|}$ מבטיח בחירה של חתימות אוניברסליות המכסות מספר רב של וקטורים בדאטה-סטים שונים.
* ברגע שווקטור $G$ מכוסה על ידי חתימת $\text{ES}$, הוא ננעל ומוקצה לה באופן חד-חד-ערכי (0 כפילויות).

---

### 7. בקרת איכות שחזור ($R^2$ ו-$\Delta R^2$)
כדי לוודא שהסט המינימלי של חתימות ה-$\text{ES}$ אינו מאבד מידע ביולוגי חשוב, הפלטפורמה מחשבת לכל דאטה-סט:
1. **$R^2_{\text{full}}$:** איכות השחזור של מטריצת הנתונים $\ln(X)$ תוך שימוש בכל וקטורי ה-$G$ של הדאטה-סט.
2. **$R^2_{\text{optimal}}$:** איכות השחזור תוך שימוש **אך ורק בווקטורי ה-$G$ שכוסו על ידי חתימות ה-$\text{ES}$ שנבחרו**.
3. **אובדן הדחיסה ($\Delta R^2 = R^2_{\text{full}} - R^2_{\text{optimal}}$):** ערך של $\Delta R^2 < 0.05$ מעיד על שחזור באיכות גבוהה מאוד (High Fidelity).

---

### 8. מבנה קבצי הפלט הממוקדים (Streamlined Deliverables)
הפלטפורמה מייצרת מערכת קבצים נקייה, ללא כפילויות וללא מידע עודף:
1. **`Optimal_ES_Result.xlsx` (קובץ המאסטר המרכזי – 5 גיליונות ממוקדים):**
   * `Summary`: סיכום פרמטרי ההרצה, נתוני הבדיקה המקדימה וכל מדדי הכיסוי וה-$R^2$.
   * `Optimal_ES_SetCover`: הטבלה המרכזית של חתימות ה-ES שנבחרו, אורך החתימה, מספר הווקטורים שכוסו (ישר `+` והפוך `-`), והכיסוי והשונות המצטברים.
   * `ES_Conflict_Free_Matrix`: מטריצת הקצאה נקייה מסתירות (דאטה-סטים × חתימות ES) המראה איזה וקטור $G$ (ובאיזה כיוון) מוקצה לכל ES.
   * `Per_Dataset_Coverage_and_R2`: גיליון מאוחד המרכז לכל דאטה-סט את אחוזי הכיסוי, השונות, ציוני ה-$R^2$ המלא והאופטימלי, ורשימת ה-ES שהוקצו לו.
   * `Per_G_Coverage_Detail`: פירוט מלא ברמת כל וקטור $G$ בודד (ה-ES שהוקצה לו, כיוון ההתאמה, אחוז החפיפה והגנים/חלבונים המשותפים).
2. **`Analysis_Summary_Report.html` (דוח מסכם חזותי):** קובץ HTML עצמאי הכולל את פרמטרי ההרצה, כרטיסי המדדים, גרף עמודות של ה-ES, מפת חום (Heatmap) של מטריצת ההקצאה, והטבלאות המרכזיות.
3. **קבצים משלימים (אופציונליים):**
   * `Anchor_Strategies_Comparison.xlsx`: גיליון השוואה מרוכז בין 5 אסטרטגיות עוגן שונות.
   * `ES_Combinations.csv`: קטלוג החתימות שנכרו אשר חוזרות בלפחות 2 וקטורי $G$ (ללא צירופים בודדים שמופיעים בווקטור אחד בלבד).

---

## Part B: Comprehensive Guide in English

### 1. Platform Scope & Scientific Objective
This platform accepts **any collection of datasets analyzed via Surprisal Analysis (SA)** that share common proteins or genes (proteomics, transcriptomics/RNA-seq, single-cell or bulk cohorts). It identifies a **minimal, conflict-free set of Exact Sign Signatures (ES)** that describes and reconstructs their unbalanced process vectors ($G_1, G_2, \dots, G_K$) while preserving regulatory polarity (`+` vs `-`).

### 2. Surprisal Analysis & Variance Weights $w(g)$
For an expression matrix $X$, log-expression is decomposed into the baseline state ($G_0, L_0$) and active biological constraints ($G_\alpha, L_\alpha$ for $\alpha \ge 1$):
$$\ln(X_{i,j}) = \sum_{\alpha=0}^K G_{i,\alpha} \cdot L_{j,\alpha}$$
Each process vector $G_k$ carries a normalized variance weight $w(g_k) = \|L_k\|^2 / \sum_{\alpha \ge 1} \|L_\alpha\|^2$.

### 3. Dual-Mode Scalable ES Mining (From 10 Proteins to 20,000+ Genes)
- **Compact Panels:** Performs exact subset enumeration $\binom{|G|}{k}$ across lengths $[L_{\min}, L_{\max}]$.
- **High-Dimensional Panels (up to 20,000+ genes & $L_{\min} = 1000+$):** Automatically engages **Sign-Consistent Pairwise & Cluster Intersection Mining**, computing direct ($+$) and inverse ($-$) concordant intersections across G-vectors in $O(N_G^2 \cdot |G|)$ time without combinatorial explosion.

### 4. Multi-Objective Greedy Set Cover & R² Verification
At each step, the greedy solver selects the candidate ES maximizing $\text{Score}(\text{ES}) = \left[\sum_{g \in \text{New}(\text{ES})} w(g)\right] \times \sqrt{|\text{New}(\text{ES})|}$, locking covered G-vectors into a conflict-free 1-to-1 allocation matrix and verifying reconstruction accuracy via $R^2_{\text{optimal}}$ and $\Delta R^2 = R^2_{\text{full}} - R^2_{\text{optimal}}$.
