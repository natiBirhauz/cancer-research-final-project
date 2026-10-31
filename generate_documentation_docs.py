"""
generate_documentation_docs.py
------------------------------
Generates high-fidelity HTML and Microsoft Word (.docx) documentation 
for the Cancer Research Surprisal Analysis & ES Algorithm Guide.
"""

import os
import re
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

MD_PATH = os.path.join('Docs', 'Algorithm_Explanation_Guide.md')
HTML_PATH = os.path.join('Docs', 'Algorithm_Explanation_Guide.html')
DOCX_PATH = os.path.join('Docs', 'Algorithm_Explanation_Guide.docx')
SVG_PATH = os.path.join('Docs', 'algorithm_workflow_schema.svg')

def generate_html():
    svg_embed = ""
    if os.path.exists(SVG_PATH):
        with open(SVG_PATH, 'r', encoding='utf-8') as fp:
            svg_embed = fp.read()

    template = """<!DOCTYPE html>
<html lang="he" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>מדריך אלגוריתמי מקיף: ניתוח Surprisal Analysis, רצפי ES ושחזור R²</title>
    <!-- KaTeX for math rendering -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.css">
    <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/katex.min.js"></script>
    <script defer src="https://cdn.jsdelivr.net/npm/katex@0.16.8/dist/contrib/auto-render.min.js" 
            onload="renderMathInElement(document.body);"></script>
    <style>
        :root {
            --primary: #1e40af;
            --primary-light: #3b82f6;
            --bg-body: #f8fafc;
            --bg-card: #ffffff;
            --text-dark: #0f172a;
            --text-muted: #64748b;
            --border: #e2e8f0;
            --success: #15803d;
            --warning: #b45309;
            --danger: #b91c1c;
        }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            background-color: var(--bg-body);
            color: var(--text-dark);
            margin: 0;
            padding: 40px 20px;
            line-height: 1.8;
            font-size: 16px;
        }
        .container {
            max-width: 1100px;
            margin: 0 auto;
            background: var(--bg-card);
            border-radius: 16px;
            padding: 50px 60px;
            box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05), 0 8px 10px -6px rgba(0, 0, 0, 0.01);
            border: 1px solid var(--border);
        }
        h1, h2, h3, h4 {
            color: var(--primary);
            font-weight: 700;
        }
        h1 {
            font-size: 2.2em;
            border-bottom: 3px solid var(--primary-light);
            padding-bottom: 16px;
            margin-bottom: 24px;
        }
        h2 {
            font-size: 1.6em;
            margin-top: 40px;
            border-bottom: 1px solid var(--border);
            padding-bottom: 8px;
        }
        h3 {
            font-size: 1.25em;
            margin-top: 25px;
        }
        p, li {
            font-size: 1.05em;
            color: #334155;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 24px 0;
            font-size: 0.95em;
            border-radius: 8px;
            overflow: hidden;
            border: 1px solid var(--border);
        }
        th, td {
            padding: 12px 16px;
            text-align: right;
            border-bottom: 1px solid var(--border);
        }
        th {
            background: #1e3a8a;
            color: #ffffff;
            font-weight: 600;
        }
        tr:nth-child(even) {
            background: #f8fafc;
        }
        tr:hover {
            background: #f1f5f9;
        }
        .badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 0.85em;
            font-weight: 600;
        }
        .badge-winner {
            background: #dcfce7;
            color: #15803d;
            border: 1px solid #86efac;
        }
        .callout {
            background: #eff6ff;
            border-right: 5px solid var(--primary-light);
            padding: 20px;
            border-radius: 8px 0 0 8px;
            margin: 25px 0;
        }
        .schema-container {
            margin: 40px 0;
            text-align: center;
            border: 1px solid var(--border);
            border-radius: 12px;
            padding: 20px;
            background: #ffffff;
            box-shadow: 0 4px 12px rgba(0,0,0,0.03);
        }
        code {
            background: #f1f5f9;
            padding: 2px 6px;
            border-radius: 4px;
            font-family: Consolas, Monaco, monospace;
            font-size: 0.9em;
            color: #0f172a;
        }
        pre {
            background: #0f172a;
            color: #f8fafc;
            padding: 18px 24px;
            border-radius: 8px;
            overflow-x: auto;
            direction: ltr;
            text-align: left;
            font-family: Consolas, Monaco, monospace;
            font-size: 0.9em;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>מדריך אלגוריתמי מקיף: ניתוח Surprisal Analysis, כריית רצפי ES ושחזור R²</h1>
        <p><strong>פרויקט גמר: חקר המיקרו-סביבה הגידולית (TME) בטכנולוגיית IMC</strong></p>

        <div class="callout">
            <strong>תקציר מנהלים:</strong> מסמך זה מפרט בצורה פשוטה, ברורה ומעמיקה את מלוא השלבים של האלגוריתם שפותח — החל מפירוק נתוני ביטוי החלבונים בשיטת Surprisal Analysis, דרך חילוץ זנבות הפעילות, כריית 120,791 רצפי סימנים מדויקים (ES), ועד לאופטימיזציה החמדנית ההיברידית (32 מודולים בלבד) המשחזרת 99.93% מהווקטורים ו-87.6% מהשונות הביולוגית ברמת התא הבודד.
        </div>

        <h2>סכמת זרימה חזותית של האלגוריתם (Workflow Schema)</h2>
        <div class="schema-container">
            __SVG_EMBED__
        </div>

        <h2>1. מבוא והאתגר המדעי</h2>
        <p>במחקרי סרטן מודרניים, סביבת המיקרו-סביבה הגידולית (TME) מהווה מערכת מורכבת הכוללת תאי חיסון, תאי סטרומה, כלי דם ותאי סרטן. בטכנולוגיית <strong>Imaging Mass Cytometry (IMC)</strong> נמדדו עשרות סמנים חלבוניים ב-120 סקשנים רקמתיים (ROIs) מ-19 חולים.</p>
        <p>האתגר המרכזי הינו מציאת קבוצה מינימלית של תבניות פעילות חלבוניות משותפות, תוך שמירה מלאה על כיוון הפעילות (חיובי או שלילי) ואפס כפילויות, ושחזור נאמן של המידע המקורי ברמת התא הבודד.</p>

        <h2>2. אנליזת הפתעה (Surprisal Analysis)</h2>
        <p>שיטה תרמודינמית מתורת האינפורמציה המפרקת את מטריצת ביטוי החלבונים X לפי:</p>
        <p style="text-align: center; font-size: 1.2em; font-family: monospace;">ln(X) = G &times; L<sup>T</sup></p>
        <ul>
            <li><strong>וקטור G_k:</strong> מתאר את טביעת האצבע החלבונית של התהליך הביולוגי k.</li>
            <li><strong>וקטור L_k:</strong> מתאר את עוצמת הפעילות של התהליך בכל תא ותא.</li>
            <li><strong>רכיב הבסיס (G0, L0):</strong> שיווי המשקל הסטטיסטי ברקמה.</li>
            <li><strong>משקולת השונות הפיזיקלית w(g):</strong> מוגדרת כיחס האנרגיה w(g) = ||L_k||² / &Sigma;||L||².</li>
        </ul>

        <h2>3. מושג ה-"זנבות" (Tails: G+ ו-G-)</h2>
        <p>בהתפלגות ערכי G, רוב החלבונים מתרכזים סביב האפס (רעש). המניעים האמיתיים מרוכזים בקצוות:</p>
        <ul>
            <li><strong>G+ (זנב חיובי):</strong> חלבונים המבוטאים ביתר כאשר התהליך פעיל.</li>
            <li><strong>G- (זנב שלילי):</strong> חלבונים המדוכאים כאשר התהליך פעיל.</li>
        </ul>
        <p>נאספו סך הכל 2,778 וקטורי G בעלי אורכי זנב של 2 עד 8 חלבונים.</p>

        <h2>4. רצפי סימנים מדויקים (Exact Sign Sequences - ES)</h2>
        <p>רצף ES מגדיר תת-קבוצת חלבונים והסימן המדויק שלהם (`+` או `-`).</p>
        <p><strong>נורמליזציה קנונית:</strong> מכיוון שכיוון הווקטור העצמי שרירותי (G &times; L<sup>T</sup> = -G &times; -L<sup>T</sup>), מתקיים ES &equiv; -ES. כל רצף מנורמל אלפביתית ומיוצג באופן חד-חד-ערכי.</p>

        <h2>5. כריית 120,791 הצירופים: Singletons מול 2+ Gs</h2>
        <table>
            <thead>
                <tr>
                    <th>אורך הרצף</th>
                    <th>סך כל הצירופים</th>
                    <th>צירופים שכיחים (2+ Gs)</th>
                    <th>צירוף מוביל</th>
                </tr>
            </thead>
            <tbody>
                <tr><td>3</td><td>18,081</td><td>6,589</td><td>CD20(+) | CD3(+) | CD8(+) (ב-80 וקטורים)</td></tr>
                <tr><td>4</td><td>43,866</td><td>9,878</td><td>CD20(+) | CD3(+) | CD31(-) | HLADR(-) (ב-38 וקטורים)</td></tr>
                <tr><td>5</td><td>37,085</td><td>5,302</td><td>CD20(+) | CD3(+) | CD31(-) | CD8(+) | HLADR(-) (ב-24 וקטורים)</td></tr>
                <tr><td>6</td><td>16,922</td><td>1,446</td><td>C11(+) | CD20(-) | CD3(-) | CD31(+) | CD8(-) | HLADR(+) (ב-13 וקטורים)</td></tr>
                <tr><td>7</td><td>4,344</td><td>277</td><td>C11(+) | CD20(-) | CD3(-) | CD31(+) | CD8(-) | HLADR(+) | Ki67(+) (ב-7 וקטורים)</td></tr>
                <tr><td>8</td><td>493</td><td>19</td><td>CD11c(+) | CD163(-) | CD20(+) | CD3(+) | CD31(-) | HLADR(-) | LAG3(-) | PDPN(+)</td></tr>
                <tr style="font-weight: bold; background: #e2e8f0;">
                    <td>סה"כ</td><td>120,791</td><td>23,511 (19.5%)</td><td>כל מרחב הדאטא מאורך 3</td>
                </tr>
            </tbody>
        </table>

        <h2>6. מודל ההתאמה וכלל אפס הסתירות (Zero Contradictions)</h2>
        <p>התאמה ישירה (`+`) דורשת שכל החלבונים המשותפים יהיו זהי סימנים. התאמה הפוכה (`-`) דורשת סימנים הפוכים לחלוטין. <strong>סתירה בסימן של חלבון בודד פוסלת מיידית את ההתאמה!</strong></p>

        <h2>7. האלגוריתם החמדן ופונקציית המטרה ההיברידית</h2>
        <p>האלגוריתם בוחר בכל שלב את המודול הממקסם את הציון המשולב:</p>
        <p style="text-align: center; font-size: 1.2em; font-family: monospace;">Score(ES) = [&Sigma; w(g)] &times; &radic;(New_G_Count)</p>
        <p>וקטורים שנקלטו מוסרים ממאגר הבחירה — הקצאה זרה לחלוטין (0 כפילויות וקטורים).</p>

        <h2>8. השוואת 6 אסטרטגיות עוגן</h2>
        <table>
            <thead>
                <tr>
                    <th>אסטרטגיה</th>
                    <th>מודולי ES</th>
                    <th>וקטורי G מכוסים</th>
                    <th>כיסוי %</th>
                    <th>שונות מוסברת %</th>
                    <th>R² אופטימלי</th>
                    <th>מחיר קומפרסיה ΔR²</th>
                </tr>
            </thead>
            <tbody>
                <tr><td>1. SVD Variance Driver</td><td>31</td><td>2,775</td><td>99.89%</td><td>87.60%</td><td>0.8837</td><td>0.0785</td></tr>
                <tr><td>2. Max Frequency (Universality)</td><td>31</td><td>2,774</td><td>99.86%</td><td>87.57%</td><td>0.8834</td><td>0.0788</td></tr>
                <tr><td>3. Deep Specificity (Lengths 7-8)</td><td>32</td><td>2,775</td><td>99.89%</td><td>87.59%</td><td>0.8837</td><td>0.0785</td></tr>
                <tr><td>4. Immune Seed (CD3+/CD8+)</td><td>33</td><td>2,775</td><td>99.89%</td><td>87.59%</td><td>0.8836</td><td>0.0787</td></tr>
                <tr><td>5. Stroma Seed (FAP+/SMA+)</td><td>34</td><td>2,774</td><td>99.86%</td><td>87.55%</td><td>0.8833</td><td>0.0790</td></tr>
                <tr style="background: #dcfce7; font-weight: bold;">
                    <td>6. Hybrid Model <span class="badge badge-winner">המנצח</span></td>
                    <td>32</td><td>2,776</td><td>99.93%</td><td>87.61%</td><td>0.8838</td><td>0.0784</td>
                </tr>
            </tbody>
        </table>

        <h2>9. שחזור R² ואיכות המידע</h2>
        <ul>
            <li><strong>ממוצע R² אופטימלי = 0.8838</strong> עם עלות קומפרסיה זניחה של &Delta;R² = 0.0784.</li>
            <li><strong>92.5% מהסקשנים</strong> דורגו באיכות שחזור גבוהה (מצוין או סביר).</li>
            <li>מטריצת ההקצאה בת 120 השורות נקייה לחלוטין מכל כפילות שורה (0 row duplicates).</li>
        </ul>

        <h2>10. בחינת הדאטאסט הנוסף (Evaluation Datasets)</h2>
        <p>האלגוריתם הופעל על 6 דאטאסטים נוספים (47,932 תאים, 10 חלבונים, 54 וקטורי G). תוך 4 מודולי ES בלבד הושג <strong>כיסוי מלא של 100%</strong> לכלל הווקטורים ללא שום סתירה או כפילות.</p>

        <h2>11. מדריך הרצת ממשק המשתמש (GUI)</h2>
        <p>להרצת ממשק הרשת (Gradio):</p>
        <pre>python run_gui.py</pre>
        <p>להרצת ממשק חלונות טבעי (Tkinter Desktop):</p>
        <pre>python run_gui.py --desktop</pre>
    </div>
</body>
</html>
"""
    html_content = template.replace("__SVG_EMBED__", svg_embed)
    with open(HTML_PATH, 'w', encoding='utf-8') as fp:
        fp.write(html_content)
    print(f"Generated HTML Guide: {HTML_PATH}")

def generate_docx():
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    title_p = doc.add_paragraph()
    title_run = title_p.add_run("מדריך אלגוריתמי מקיף: ניתוח Surprisal Analysis, רצפי ES ושחזור R²")
    title_run.font.name = "Calibri"
    title_run.font.size = Pt(22)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    title_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    sub_p = doc.add_paragraph()
    sub_run = sub_p.add_run("פרויקט גמר: חקר המיקרו-סביבה הגידולית (TME) בשיטת IMC | תיעוד מקיף וברור")
    sub_run.font.name = "Calibri"
    sub_run.font.size = Pt(12)
    sub_run.font.italic = True
    sub_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    sub_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_paragraph("―" * 45).alignment = WD_ALIGN_PARAGRAPH.RIGHT

    sections_text = [
        ("1. מבוא והאתגר המדעי", 
         "במחקרי סרטן מודרניים, סביבת המיקרו-סביבה הגידולית (TME) מהווה מערכת אקולוגית מורכבת הכוללת תאי חיסון, תאי סטרומה, כלי דם ותאי סרטן. בטכנולוגיית Imaging Mass Cytometry (IMC) נמדדו עשרות סמנים חלבוניים ב-120 סקשנים רקמתיים (ROIs) מ-19 חולים.\n\n"
         "האתגר המרכזי הינו מציאת קבוצה מינימלית של תבניות פעילות חלבוניות משותפות (Exact Sign Sequences - ES), תוך שמירה מלאה על כיוון הפעילות (חיובי או שלילי) ואפס כפילויות, ושחזור נאמן של המידע המקורי ברמת התא הבודד לפי מדד R²."),

        ("2. אנליזת הפתעה (Surprisal Analysis)",
         "שיטת ה-Surprisal Analysis מפרקת את מטריצת ביטוי החלבונים X לפי נוסחת הפירוק:\n"
         "ln(X) = G × L^T\n\n"
         "• וקטור G_k: מתאר את טביעת האצבע החלבונית של התהליך הביולוגי k.\n"
         "• וקטור L_k: מתאר את עוצמת הפעילות של התהליך בכל תא ותא ברקמה.\n"
         "• רכיב הבסיס (G0, L0): מייצג את שיווי המשקל הסטטיסטי ברקמה.\n"
         "• משקולת השונות הפיזיקלית w(g): מוגדרת כיחס האנרגיה w(g) = ||L_k||² / Σ||L||²."),

        ("3. מושג ה-'זנבות' (Tails: G+ ו-G-)",
         "בהתפלגות ערכי G, רוב החלבונים מתרכזים סביב האפס (רעש שאינו פעיל בתהליך). המניעים האמיתיים מרוכזים בקצוות (זנבות ההתפלגות):\n"
         "• G+ (זנב חיובי): חלבונים המבוטאים ביתר כאשר התהליך פעיל.\n"
         "• G- (זנב שלילי): חלבונים המדוכאים כאשר התהליך פעיל.\n"
         "נאספו סך הכל 2,778 וקטורי G בעלי אורכי זנב של 2 עד 8 חלבונים."),

        ("4. רצפי סימנים מדויקים (Exact Sign Sequences - ES) ונורמליזציה קנונית",
         "רצף ES מגדיר תת-קבוצת חלבונים והסימן המדויק שלהם (+ או -).\n\n"
         "נורמליזציה קנונית:\n"
         "מכיוון שכיוון הווקטור העצמי נקבע שרירותית בפירוק (G × L^T = -G × -L^T), רצף שכל סימניו הפוכים (-ES) מייצג בדיוק את אותו תהליך ביולוגי. כדי למנוע כפילויות, כל רצף מנורמל אלפביתית ואם החלבון הראשון שלילי הופכים את כל הסימנים. כך כל תהליך מקבל מפתח ייחודי וקנוני."),

        ("5. כריית 120,791 הצירופים: Singletons מול 2+ Gs",
         "מתוך 2,778 וקטורי ה-G נכרו כל תתי-הצירופים האפשריים החל מאורך 3 חלבונים ועד 8 חלבונים (120,791 צירופים ייחודיים סה\"כ).\n\n"
         "ההבחנה המכרעת:\n"
         "• צירופים בודדים (Singletons - 80.5%): הופיעו בווקטור G יחיד בכל המחקר — מהווים רעש או שונות נקודתית.\n"
         "• צירופים שכיחים (2+ Gs - 19.5%): הופיעו בלפחות 2 וקטורים שונים. חזרתיות זו היא ההוכחה האמפירית לתבנית ביולוגית יציבה ואמיתית!"),

        ("6. מודל ההתאמה הביולוגית וכלל אפס הסתירות",
         "כאשר בוחנים האם מודול ES מכסה וקטור G מסוים:\n"
         "1. נדרשת חפיפה של לפחות 2 חלבונים.\n"
         "2. התאמה ישירה (+): כל החלבונים המשותפים מופיעים עם אותו סימן בדיוק.\n"
         "3. התאמה הפוכה (-): כל החלבונים המשותפים מופיעים עם סימן הפוך בדיוק.\n"
         "4. כלל אפס הסתירות (Zero Contradictions): סתירה בסימן של חלבון בודד פוסלת מיידית ומוחלטת את ההתאמה!"),

        ("7. האלגוריתם החמדן ופונקציית המטרה ההיברידית",
         "פונקציית המטרה ההיברידית המנצחת:\n"
         "Score(ES) = [Σ w(g)] × √(New_G_Count)\n\n"
         "היא משלבת שונות פיזיקלית מרבית עם תפוצה רחבה באוכלוסייה, ומפעילה הקצאה זרה (Disjoint G-Allocation) כך שאף וקטור G אינו משוכפל."),

        ("8. תוצאות השוואת 6 אסטרטגיות עוגן",
         "הניסוי הוכיח יציבות פנומנלית: כל 6 האסטרטגיות התכנסו ל-31-34 מודולים עם כיסוי של 99.9% ו-87.6% שונות מוסברת. המודל ההיברידי ניצח עם 32 מודולי ES, כיסוי של 2,776 מתוך 2,778 וקטורים (99.93%), ושחזור ממוצע של R²_optimal = 0.8838 (עם מחיר קומפרסיה מזערי של ΔR² = 0.0784 בלבד)."),

        ("9. שחזור R² ואיכות המידע ברמת התא הבודד",
         "מדד R² מחושב עבור כל תא ותא בנפרד מול ביטוי החלבונים המקורי:\n"
         "92.5% מכלל 120 הסקשנים דורגו ברמת שחזור גבוהה (מצוין או סביר), תוך 0 כפילויות שורה במטריצת ההקצאה."),

        ("10. בחינת הדאטאסט הנוסף (Data for algorithm evaluation)",
         "האלגוריתם נבחן בהצלחה מלאה על 6 דאטאסטים נוספים (47,932 תאים, 10 חלבונים, 54 וקטורי G).\n"
         "4 מודולי ES בלבד השיגו 100% כיסוי של כלל הווקטורים, עם אפס סתירות ואפס כפילויות."),

        ("11. ממשק המשתמש (GUI)",
         "פותחו שני ממשקים נוחים ונגישים:\n"
         "1. ממשק רשת מודרני בדפדפן (Gradio): הרצה באמצעות python run_gui.py\n"
         "2. ממשק שולחן עבודה מקומי לחלונות (Tkinter): הרצה באמצעות python run_gui.py --desktop")
    ]

    for title, content in sections_text:
        h = doc.add_heading(level=2)
        h_run = h.add_run(title)
        h_run.font.name = "Calibri"
        h_run.font.size = Pt(14)
        h_run.font.bold = True
        h_run.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
        h.alignment = WD_ALIGN_PARAGRAPH.RIGHT

        p = doc.add_paragraph()
        p_run = p.add_run(content)
        p_run.font.name = "Calibri"
        p_run.font.size = Pt(11)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.save(DOCX_PATH)
    print(f"Generated Word Document: {DOCX_PATH}")

if __name__ == '__main__':
    generate_html()
    generate_docx()
