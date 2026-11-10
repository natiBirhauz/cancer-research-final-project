# מפרט טכני: קטלוג החתימות החוזרות (`ES_Combinations.csv`)
# Technical Specification: Recurring ES Signatures Catalog (`ES_Combinations.csv`)

מסמך זה מתאר את המבנה והמשמעות של הקובץ **`ES_Combinations.csv`** המופק על ידי פלטפורמת ה-Surprisal Analysis (SA).

---

## 1. מהי חתימת Exact Sign Signature (ES) וכיצד מסונן הקובץ?
* **חתימת ES:** צירוף של גנים או חלבונים בעלי סימן כיווני מוגדר (`+` לביטוי-יתר או `-` לדיכוי) מתוך וקטורי ה-$G$ של הדאטה-סטים, מנורמל לצורה קנונית אחידה ($\text{ES} \equiv -\text{ES}$).
* **סינון צירופים בודדים (Singletons):** בדאטה-סטים רחבי-היקף נכרים לעיתים עשרות אלפי צירופים שמופיעים בווקטור $G$ אחד בלבד (רעש נקודתי). כדי למנוע עומס ומידע מיותר, הקובץ `ES_Combinations.csv` מסנן החוצה צירופים מונו-ווקטוריים ומכיל **אך ורק חתימות חוזרות** המופיעות בלפחות `Min G Recurrence` וקטורי $G$ (ברירת מחדל: $\ge 2$ וקטורים).

---

## 2. מבנה העמודות בקובץ (`ES_Combinations.csv`)

| # | שם העמודה (Column Name) | סוג נתון | הסבר בעברית | Description in English |
|---|:---|:---:|:---|:---|
| 1 | `Length` | Integer | מספר החלבונים או הגנים בחתימת ה-ES. | Number of genes/proteins in the signature. |
| 2 | `Frequency_Count` | Integer | מספר וקטורי ה-$G$ השונים שבהם חתימת ה-ES מופיעה ללא סתירות. | Number of distinct G-vectors matched by this signature. |
| 3 | `Exact_Sign_Sequence` | String | רצף הגנים/חלבונים והסימנים שלהם בצורה קנונית ממוינת אלפביתית. | Alphabetically ordered canonical gene/protein sign sequence. |
| 4 | `Covered_G_Vectors` | String | רשימת מזהי וקטורי ה-$G$ המכוסים על ידי החתימה (מופרדים ב-`;`). | Semicolon-delimited list of matching G-vectors across datasets. |
