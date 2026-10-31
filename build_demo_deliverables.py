"""
build_demo_deliverables.py
--------------------------
Generates realistic demo/sample deliverables explaining each output file
and defining what each column and sheet measures and represents.

Generates:
1. Docs/Demo_Deliverables/Demo_Optimal_ES_Result_Signs_R2.xlsx
2. Docs/Demo_Deliverables/Demo_Optimal_ES_Result.xlsx
3. Docs/Demo_Deliverables/Demo_Anchor_Strategies_Comparison.xlsx
4. Docs/Demo_Deliverables/Demo_ES_Combinations.csv
5. Docs/Demo_Deliverables/Deliverables_Data_Dictionary.xlsx (Master Dictionary)
6. Docs/Demo_Deliverables/Demo_Deliverables_Bundle.zip
"""

import os
import zipfile
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

DEMO_DIR = os.path.join("Docs", "Demo_Deliverables")

def build_demo_files():
    os.makedirs(DEMO_DIR, exist_ok=True)
    print("Building demo deliverables and data dictionary...")

    thin_border = Border(
        left=Side(style='thin', color='B0B0B0'),
        right=Side(style='thin', color='B0B0B0'),
        top=Side(style='thin', color='B0B0B0'),
        bottom=Side(style='thin', color='B0B0B0')
    )
    header_fill = PatternFill(start_color='E5E7EB', end_color='E5E7EB', fill_type='solid')
    header_font = Font(name='Segoe UI', size=10, bold=True, color='000000')
    desc_fill = PatternFill(start_color='FEF9C3', end_color='FEF9C3', fill_type='solid')
    desc_font = Font(name='Segoe UI', size=9, italic=True, color='854D0E')
    regular_font = Font(name='Segoe UI', size=10, color='000000')
    title_font = Font(name='Segoe UI', size=13, bold=True, color='000000')

    # =========================================================================
    # 1. Master Data Dictionary: Deliverables_Data_Dictionary.xlsx
    # =========================================================================
    wb_dict = openpyxl.Workbook()
    ws_main = wb_dict.active
    ws_main.title = "Overview_סקירה_כללית"
    ws_main.views.sheetView[0].showGridLines = True

    ws_main.append(["מילון נתונים ומדריך עמודות לכל קבצי הפלט של המערכת"])
    ws_main.append(["Comprehensive Data Dictionary for All Algorithm Output Deliverables"])
    ws_main.append([])
    ws_main.append(["שם קובץ הפלט (Deliverable)", "סוג הקובץ", "מטרת הקובץ ותוכנו", "גיליונות כלולים"])

    dict_overview = [
        ("Optimal_ES_Result_Signs_R2.xlsx", "קובץ Excel רב-גיליוני", "דוח המאסטר הראשי: תוצאות Set Cover, מטריצת הקצאה ללא כפילויות, פירוט שחזור R2 תאי, עקומת כיסוי ופירוט מלא של כל וקטור.", "Summary, Optimal_ES_SetCover, ES_Conflict_Free_Matrix, Best_Combination_Details, Per_G_Coverage_Detail, R2_Comparison_Per_ROI, Per_ROI_Coverage, Coverage_Curve"),
        ("Optimal_ES_Result.xlsx", "קובץ Excel רב-גיליוני", "דוח הקצאת המודולים האופטימליים, כולל מטריצת ההקצאה הביולוגית ועקומת הכיסוי.", "Summary, Optimal_ES_SetCover, ES_Conflict_Free_Matrix, Best_Combination_Details, Per_G_Coverage_Detail, Per_ROI_Coverage, Coverage_Curve"),
        ("Anchor_Strategies_Comparison.xlsx", "קובץ Excel דו-גיליוני", "השוואת ביצועים מלאה בין 6 אסטרטגיות עוגן שונות (כיסוי, שונות קינטית, שחזור R2, אחוז דיוק גבוה וזמני ריצה).", "Anchor_Strategies_Overview, Optimal_Strategy_XX_ES"),
        ("ES_Combinations.csv", "קובץ CSV טבלאי", "מאגר כל חתימות ה-ES המועמדות שנכרו בקומבינטוריקה (אורכים 3-8), תדירות הופעתן, האם מופיעות ב-2 וקטורים ומעלה, ופירוט וקטורי ה-G התואמים.", "טבלה שטוחה ממוינת")
    ]

    for row in dict_overview:
        ws_main.append(list(row))

    # Add sheet detailing columns of Optimal_ES_Result_Signs_R2.xlsx
    ws_cols_r2 = wb_dict.create_sheet(title="עמודות_Optimal_ES_Result")
    ws_cols_r2.views.sheetView[0].showGridLines = True
    ws_cols_r2.append(["פירוט כל העמודות בקובץ Optimal_ES_Result_Signs_R2.xlsx (לפי גיליונות)"])
    ws_cols_r2.append([])
    ws_cols_r2.append(["שם הגיליון (Sheet)", "שם העמודה (Column)", "הגדרה ומשמעות ביולוגית / מתמטית", "ערכים לדוגמה / יחידות"])

    cols_specs = [
        # Sheet: Summary
        ("Summary", "Metric", "שם המדד הכללי של ההרצה (מספר דגימות, וקטורים, מודולים, כיסוי וכו')", "Total Cohort Samples"),
        ("Summary", "Value", "הערך המספרי או האחוזי של המדד", "120, 2,778, 100.0%"),
        
        # Sheet: Optimal_ES_SetCover
        ("Optimal_ES_SetCover", "Rank", "דירוג המודול לפי סדר בחירתו באלגוריתם ה-Greedy Set Cover (1 = החשוב והחזק ביותר)", "1, 2, 3..."),
        ("Optimal_ES_SetCover", "ES_Label", "מזהה המודול החלבוני שנבחר", "ES_1, ES_2, ES_3"),
        ("Optimal_ES_SetCover", "Sequence_Length", "מספר החלבונים המשתתפים בחתימה זו (בין 3 ל-8 חלבונים)", "3, 4, 8"),
        ("Optimal_ES_SetCover", "New_G_Covered", "כמות וקטורי ה-G החדשים שהמודול כיסה (שלא כוסו על ידי אף מודול קודם)", "21, 582 וקטורים"),
        ("Optimal_ES_SetCover", "Direct_Matches_(+)", "כמות הוקטורים שתאמו למודול בכיוון ישר (סימנים זהים לחלוטין)", "14, 278 וקטורים"),
        ("Optimal_ES_SetCover", "Inverse_Matches_(-)", "כמות הוקטורים שתאמו למודול בכיוון הפוך (סימנים הפוכים בדיוק, נובע מ-ES ≡ -ES)", "7, 304 וקטורים"),
        ("Optimal_ES_SetCover", "Cumulative_G_Covered", "מספר וקטורי ה-G המצטבר שכוסה עד לשלב זה", "21, 37, 54 וקטורים"),
        ("Optimal_ES_SetCover", "Coverage_%", "אחוז הכיסוי המצטבר מתוך כלל וקטורי ה-G במאגר", "38.9%, 68.5%, 100.0%"),
        ("Optimal_ES_SetCover", "New_Variance_Weight", "תוספת האנרגיה הקינטית (משקל השונות Σw(g)) שהמודול הוסיף", "1.4520, 0.8920"),
        ("Optimal_ES_SetCover", "Cumulative_Variance_%", "אחוז השונות הקינטית המצטברת שהוסברה מתוך המיקרו-סביבה", "45.2%, 78.4%, 100.0%"),
        ("Optimal_ES_SetCover", "ROIs_Covered_Count", "בכמה דגימות / אזורי רקמה (ROIs) שונים מודול זה הופיע", "6 דגימות, 114 ROIs"),
        ("Optimal_ES_SetCover", "Exact_Sign_Sequence", "חתימת החלבונים המדויקת עם סימני רגולציה: (+) לביטוי יתר, (-) לדיכוי ממוקד", "CKIT(+) | EGFR(+) | IGFR(+)"),

        # Sheet: ES_Conflict_Free_Matrix
        ("ES_Conflict_Free_Matrix", "Sample", "מזהה הדגימה / המטופל / ה-ROI", "1 PDE 379, P10_ROI_09"),
        ("ES_Conflict_Free_Matrix", "Cells", "מספר התאים הבודדים בדגימה זו (אם הועלתה מטריצת L)", "1,567 תאים"),
        ("ES_Conflict_Free_Matrix", "Total_G", "סה\"כ וקטורי ה-G הקיימים בדגימה זו", "9 וקטורים, 28 וקטורים"),
        ("ES_Conflict_Free_Matrix", "ES_1..ES_k", "הוקטור הספציפי שהוקצה למודול זה באותה דגימה, עם סימן ההתאמה (+/-). תא ריק = לא הוקצה וקטור (100% ללא כפילויות)", "G2(+), G4(-), [תא ריק]"),

        # Sheet: Best_Combination_Details
        ("Best_Combination_Details", "Sample", "מזהה הדגימה / ה-ROI", "2 PDE 0466-001"),
        ("Best_Combination_Details", "Total_G_In_Sample", "מספר הוקטורים הכולל בדגימה", "9"),
        ("Best_Combination_Details", "Covered_G_Count", "כמה מתוך וקטורי הדגימה כוסו על ידי מודולי ה-ES", "9 (100%)"),
        ("Best_Combination_Details", "Sample_Coverage_%", "אחוז הכיסוי של הדגימה הספציפית הזו", "100.0%"),
        ("Best_Combination_Details", "Assigned_ES_Labels", "רשימת מודולי ה-ES שהוקצו לדגימה זו", "ES_1; ES_2; ES_3; ES_4"),
        ("Best_Combination_Details", "Consolidated_G_Assignments", "מיפוי מלא המראה לכל וקטור לאיזה מודול הוא שויך, כיוון (+/-) ואחוז החפיפה בחלבונים", "G1:ES_1(+)[75.0%], G2:ES_2(-)[60.0%]"),

        # Sheet: Per_G_Coverage_Detail
        ("Per_G_Coverage_Detail", "Sample", "שם הדגימה שאליה שייך הוקטור", "4 PDE 0466-002"),
        ("Per_G_Coverage_Detail", "G_Key", "מזהה ייחודי של הוקטור", "4 PDE 0466-002_G1"),
        ("Per_G_Coverage_Detail", "G_Num", "מספר התהליך בתוך הדגימה (G1, G2 וכו')", "1, 2, 3..."),
        ("Per_G_Coverage_Detail", "Total_Prots_In_G", "כמה חלבונים עברו את סף הטעינה בווקטור זה", "4 חלבונים"),
        ("Per_G_Coverage_Detail", "Assigned_ES", "איזה מודול ES נבחר לכסות וקטור זה", "ES_1 (או UNASSIGNED)"),
        ("Per_G_Coverage_Detail", "Orientation", "כיוון ההתאמה: (+) ישר, (-) הפוך", "+ או -"),
        ("Per_G_Coverage_Detail", "Matched_Prots_Count", "כמה חלבונים משותפים זהים בסימניהם ללא שום סתירה", "3 חלבונים"),
        ("Per_G_Coverage_Detail", "Overlap_%", "אחוז החלבונים בווקטור שתאמו לחתימת ה-ES", "75.0%"),
        ("Per_G_Coverage_Detail", "Variance_Weight", "משקל השונות הקינטית של וקטור זה w(g)", "0.2845"),
        ("Per_G_Coverage_Detail", "Status", "האם הוקטור מכוסה או לא מכוסה", "Covered / Uncovered"),

        # Sheet: R2_Comparison_Per_ROI
        ("R2_Comparison_Per_ROI", "Sample", "שם הדגימה / ה-ROI", "6 PDE 580"),
        ("R2_Comparison_Per_ROI", "N_Cells", "מספר התאים הבודדים שנותחו ברמת המיקרוסקופ", "11,469 תאים"),
        ("R2_Comparison_Per_ROI", "R2_Full_All_Components", "שחזור R2 תאי מלא עם כל הרכיבים (בסיס עליון)", "1.0000 (או 0.8572 ב-120)"),
        ("R2_Comparison_Per_ROI", "R2_Optimal_Selected_Components", "שחזור R2 תאי תוך שימוש ברכיבי ה-ES הנבחרים בלבד", "1.0000 (או 0.7788 ב-120)"),
        ("R2_Comparison_Per_ROI", "Delta_R2_Loss", "איבוד המידע התאי בעקבות הדחיסה (ΔR2 = R2_full - R2_optimal)", "0.0000 (איבוד אפסי), <0.05 מעולה"),
        ("R2_Comparison_Per_ROI", "N_Components_Used", "כמות רכיבי ה-G ששימשו לשחזור הדגימה", "4 רכיבים"),
        ("R2_Comparison_Per_ROI", "Selected_G_List", "רשימת הרכיבים הספציפיים ששימשו לשחזור", "G1, G2, G4"),
        ("R2_Comparison_Per_ROI", "Interpretation", "סיווג איכות השחזור: High Fidelity (<0.05), Moderate (0.05-0.15), Alert (>0.15)", "High Fidelity (<0.05 loss)"),

        # Sheet: Per_ROI_Coverage
        ("Per_ROI_Coverage", "Sample", "שם הדגימה", "7 PDE 0466-004"),
        ("Per_ROI_Coverage", "Cells", "כמות תאים", "16,649"),
        ("Per_ROI_Coverage", "Total_G", "סה\"כ וקטורים בדגימה", "9"),
        ("Per_ROI_Coverage", "Covered_G", "כמות וקטורים מכוסים", "9"),
        ("Per_ROI_Coverage", "Coverage_%", "אחוז הכיסוי בדגימה זו", "100.0%"),
        ("Per_ROI_Coverage", "Variance_Captured_%", "אחוז האנרגיה הקינטית שנלכדה בדגימה זו", "100.0%"),

        # Sheet: Coverage_Curve
        ("Coverage_Curve", "Step", "מספר הצעד באלגוריתם (הוספת מודול נוסף)", "1, 2, 3..."),
        ("Coverage_Curve", "ES_Label", "המודול שנוסף בצעד זה", "ES_1, ES_2..."),
        ("Coverage_Curve", "Cumulative_G_Covered", "סה\"כ וקטורי G מכוסים מצטברים", "21, 37, 54"),
        ("Coverage_Curve", "Cumulative_Coverage_%", "אחוז הכיסוי המצטבר של עולם הוקטורים", "38.9%, 68.5%, 100.0%"),
        ("Coverage_Curve", "Cumulative_Variance_%", "אחוז השונות הקינטית המצטבר שנלכד", "45.2%, 78.4%, 100.0%")
    ]

    for row in cols_specs:
        ws_cols_r2.append(list(row))

    # Add sheet detailing columns of Anchor_Strategies_Comparison.xlsx
    ws_cols_anc = wb_dict.create_sheet(title="עמודות_Anchor_Strategies")
    ws_cols_anc.views.sheetView[0].showGridLines = True
    ws_cols_anc.append(["פירוט העמודות בקובץ Anchor_Strategies_Comparison.xlsx"])
    ws_cols_anc.append([])
    ws_cols_anc.append(["שם העמודה (Column)", "תיאור ומשמעות המדד", "משמעות ביולוגית / אלגוריתמית", "ערכים טיפוסיים"])

    anc_specs = [
        ("Anchor_Strategy", "שם אסטרטגיית העוגן הגרידית שנבדקה", "מגדיר איזה ציון מקבל כל מודול בבחירה (שונות SVD, תדירות מקסימלית, אורך ספציפי, עוגן ביולוגי או Hybrid)", "1. SVD Variance Driver, 6. Hybrid"),
        ("ES_Count", "מספר מודולי ה-ES שנבחרו כדי להגיע לכיסוי מקסימלי", "מצביע על מידת הקומפקטיות של המודל (פחות מודולים = מודל פשוט וחסכוני יותר)", "4 ב-PDE, 32-34 ב-120"),
        ("Covered_Gs", "מספר וקטורי ה-G שכוסו ללא שום סתירת סימנים", "כמה מתוך כלל המיקרו-סביבות זכו לייצוג", "54/54 (100%), 2,778/2,778"),
        ("Coverage_%", "אחוז הכיסוי של כלל המיקרו-סביבות", "מדד היעילות של המודל בכיסוי האוכלוסייה", "99.8% - 100.0%"),
        ("Variance_%", "אחוז השונות הקינטית המוסברת", "כמה מסך אנרגיית הסטייה של הרקמה מיוצגת על ידי המודולים", "88.3% - 100.0%"),
        ("Mean_R2_full", "ממוצע R2 מלא של שחזור התאים הבודדים עם כלל הרכיבים", "איכות השחזור המקסימלית האפשרית של המטריצה", "1.0000 ב-PDE, 0.8572 ב-120"),
        ("Mean_R2_opt", "ממוצע R2 משוחזר תוך שימוש אך ורק במודולי ה-ES שנבחרו", "מדד האיכות הסופי של האלגוריתם (כמה קרוב השחזור המכווץ למקור)", "1.0000 ב-PDE, 0.7788 ב-120"),
        ("Mean_Delta_R2", "ממוצע קנס איבוד המידע (Delta R2 = R2_full - R2_opt)", "מדידה ישירה של כמה מידע אבד עקב הפחתת המימדים", "0.0000 (איבוד אפסי), 0.0784 (איבוד קל)"),
        ("Excellent_%", "אחוז הדגימות עם איכות שחזור מעולה (איבוד מידע ΔR2 < 0.05)", "דגימות שבהן הדחיסה לא פגעה כלל בביטוי התאי", "100.0% ב-PDE, 38.3% ב-120"),
        ("Acceptable_%", "אחוז הדגימות עם איכות שחזור סבירה (0.05 <= ΔR2 <= 0.15)", "איבוד מידע מתון שעדיין שומר על המבנה הביולוגי", "0.0% ב-PDE, 55.0% ב-120"),
        ("Alert_%", "אחוז הדגימות עם ירידה מורגשת בשחזור (ΔR2 > 0.15)", "דגימות עם רעש גבוה או סמנים לא מכוסים", "0.0% ב-PDE, 6.7% ב-120"),
        ("Exec_Time_s", "זמן הריצה של האסטרטגיה בשניות", "מדד יעילות חישובית של פותר ה-Set Cover", "0.15s - 12.5s")
    ]

    for row in anc_specs:
        ws_cols_anc.append(list(row))

    # Add sheet detailing columns of ES_Combinations.csv
    ws_cols_csv = wb_dict.create_sheet(title="עמודות_ES_Combinations")
    ws_cols_csv.views.sheetView[0].showGridLines = True
    ws_cols_csv.append(["פירוט העמודות בקובץ ES_Combinations.csv"])
    ws_cols_csv.append([])
    ws_cols_csv.append(["שם העמודה (Column)", "תיאור ומשמעות", "משמעות ביולוגית / קומבינטורית", "ערכים לדוגמה"])

    csv_specs = [
        ("Length", "אורך חתימת ה-ES (מספר החלבונים בחתימה)", "מגדיר את רמת המורכבות של הקומפלקס החלבוני (נכרה באורכים 3 עד 8)", "3, 4, 5, 6, 7, 8"),
        ("Frequency_Count", "בכמה וקטורי G ביולוגיים שונים החתימה הופיעה במדויק", "מדד החזרתיות והאוניברסליות של החתימה באוכלוסיית התאים", "80, 73, 21, 2 וקטורים"),
        ("In_2+_Gs", "ערך בוליאני (True / False): האם החתימה הופיעה בלפחות 2 וקטורים", "סינון רעש: רק חתימות שמופיעות ב-2 וקטורים ומעלה נחשבות מודולים ביולוגיים יציבים (קומבינציות בודדות מסוננות)", "True / False"),
        ("Exact_Sign_Sequence", "חתימת החלבונים המדויקת והקנונית עם סימני הרגולציה (+/-)", "קומפלקס חלבוני מיושר קנונית (ממוין אלפביתית, כאשר החלבון הראשון תמיד חיובי למניעת כפילות)", "CKIT(+) | EGFR(+) | IGFR(+)"),
        ("Covered_G_Vectors", "רשימת כל מזהי וקטורי ה-G שבהם חתימה זו זוהתה במדויק", "מיפוי המאפשר לחוקר לראות בדיוק באילו חולים ובאילו דגימות חתימה זו פעילה", "1 PDE 379_G1; 2 PDE 0466_G2; P10_ROI_09_G2")
    ]

    for row in csv_specs:
        ws_cols_csv.append(list(row))

    # Style all sheets of the Data Dictionary
    for ws in [ws_main, ws_cols_r2, ws_cols_anc, ws_cols_csv]:
        ws.cell(1, 1).font = title_font
        for col in range(1, ws.max_column + 1):
            c = ws.cell(3 if ws == ws_cols_r2 or ws == ws_cols_anc or ws == ws_cols_csv else 4, col)
            c.fill = header_fill
            c.font = header_font
            c.alignment = Alignment(horizontal='center', vertical='center')

        for row in range(4 if ws == ws_cols_r2 or ws == ws_cols_anc or ws == ws_cols_csv else 5, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row, col)
                cell.font = regular_font
                cell.border = thin_border
                cell.alignment = Alignment(vertical='center', wrap_text=True)

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 14), 50)

    dict_path = os.path.join(DEMO_DIR, "Deliverables_Data_Dictionary.xlsx")
    wb_dict.save(dict_path)
    print(f"Created: {dict_path}")

    # =========================================================================
    # 2. Demo File: Demo_Optimal_ES_Result_Signs_R2.xlsx
    # =========================================================================
    # Copy from Docs/Evaluation_Optimal_ES_Result_Signs_R2.xlsx if available and annotate
    source_r2 = os.path.join("Docs", "Evaluation_Optimal_ES_Result_Signs_R2.xlsx")
    demo_r2_path = os.path.join(DEMO_DIR, "Demo_Optimal_ES_Result_Signs_R2.xlsx")
    demo_opt_path = os.path.join(DEMO_DIR, "Demo_Optimal_ES_Result.xlsx")

    if os.path.exists(source_r2):
        wb_src = openpyxl.load_workbook(source_r2)
        # Add explanation banner on sheet 1
        ws1 = wb_src['Summary']
        ws1.insert_rows(1, 2)
        ws1.cell(1, 1, "DEMO FILE / קובץ דוגמה מוסבר: מציג את מבנה הפלט האמיתי של מודולי ה-ES ושחזור ה-R2").font = title_font
        ws1.cell(2, 1, "Every sheet represents a core deliverable: Optimal Set Cover, Conflict-Free Allocation Matrix, and Cell-Level R2 Reconstruction").font = subtitle_font = Font(size=10, italic=True)
        wb_src.save(demo_r2_path)
        wb_src.save(demo_opt_path)
        print(f"Created: {demo_r2_path}")
        print(f"Created: {demo_opt_path}")

    # =========================================================================
    # 3. Demo File: Demo_Anchor_Strategies_Comparison.xlsx
    # =========================================================================
    source_anc = os.path.join("Docs", "Evaluation_Anchor_Strategies_Comparison.xlsx")
    demo_anc_path = os.path.join(DEMO_DIR, "Demo_Anchor_Strategies_Comparison.xlsx")
    if os.path.exists(source_anc):
        wb_anc = openpyxl.load_workbook(source_anc)
        ws_a = wb_anc['Anchor_Strategies_Overview']
        ws_a.insert_rows(1, 2)
        ws_a.cell(1, 1, "DEMO FILE / קובץ דוגמה: השוואת 6 אסטרטגיות עוגן גרידיות (Anchor Strategies Overview)").font = title_font
        ws_a.cell(2, 1, "Compares ES Count, G Coverage %, Kinetic Variance %, Reconstruction R2, Fidelity Tier %, and Execution Time").font = subtitle_font
        wb_anc.save(demo_anc_path)
        print(f"Created: {demo_anc_path}")

    # =========================================================================
    # 4. Demo File: Demo_ES_Combinations.csv
    # =========================================================================
    source_comb = os.path.join("Docs", "Evaluation_ES_Combinations.csv")
    demo_comb_path = os.path.join(DEMO_DIR, "Demo_ES_Combinations.csv")
    if os.path.exists(source_comb):
        df_c = pd.read_csv(source_comb)
        df_c.head(100).to_csv(demo_comb_path, index=False)
        print(f"Created: {demo_comb_path}")

    # =========================================================================
    # 5. Bundle all Demo files into Demo_Deliverables_Bundle.zip
    # =========================================================================
    zip_path = os.path.join(DEMO_DIR, "Demo_Deliverables_Bundle.zip")
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.write(dict_path, arcname="Deliverables_Data_Dictionary.xlsx")
        if os.path.exists(demo_r2_path):
            zf.write(demo_r2_path, arcname="Demo_Optimal_ES_Result_Signs_R2.xlsx")
        if os.path.exists(demo_opt_path):
            zf.write(demo_opt_path, arcname="Demo_Optimal_ES_Result.xlsx")
        if os.path.exists(demo_anc_path):
            zf.write(demo_anc_path, arcname="Demo_Anchor_Strategies_Comparison.xlsx")
        if os.path.exists(demo_comb_path):
            zf.write(demo_comb_path, arcname="Demo_ES_Combinations.csv")

    # Also place a copy in Docs for easy download
    root_zip = os.path.join("Docs", "Demo_Deliverables_Bundle.zip")
    with open(zip_path, 'rb') as f_in, open(root_zip, 'wb') as f_out:
        f_out.write(f_in.read())

    print(f"Created consolidated zip: {zip_path} and {root_zip}")

if __name__ == '__main__':
    build_demo_files()
