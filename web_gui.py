"""
web_gui.py
----------
Simplified, Clean High-Contrast Web User Interface for the
Cancer Research TME Surprisal Analysis & Exact Sign Sequence (ES) Platform.

Key Features:
- 100% Bilingual: High-contrast English layout + Complete Hebrew User Guide Section.
- Forces Light Mode (__theme=light) to prevent OS dark mode conflicts.
- Downloadable input template (Sample_Input_Template.xlsx).
- Downloadable demo deliverables & column data dictionary (Deliverables_Data_Dictionary.xlsx).
- User-supplied dataset ingestion (Excel files or .zip archive).
- Automated execution generating the 4 master deliverables:
  1. Optimal_ES_Result.xlsx
  2. Optimal_ES_Result_Signs_R2.xlsx
  3. ES_Combinations.csv
  4. Anchor_Strategies_Comparison.xlsx
  Plus Deliverables_Bundle.zip (consolidated 1-click download).
- Ultra-streamlined interface without clutter or heavy preview tables.
"""

import os
import sys
import time
import pandas as pd
import gradio as gr
from pipeline_engine import run_user_pipeline

TEMPLATE_FILE = "Sample_Input_Template.xlsx"
if not os.path.exists(TEMPLATE_FILE):
    TEMPLATE_FILE = os.path.join("Docs", "Sample_Input_Template.xlsx")

DICT_FILE = os.path.join("Docs", "Deliverables_Data_Dictionary.xlsx")
if not os.path.exists(DICT_FILE):
    DICT_FILE = os.path.join("Docs", "Demo_Deliverables", "Deliverables_Data_Dictionary.xlsx")

HEAD_SCRIPT = """
<script>
    if (window.location.search.indexOf('__theme=light') === -1) {
        const url = new URL(window.location);
        url.searchParams.set('__theme', 'light');
        window.location.replace(url.href);
    }
    document.documentElement.classList.remove('dark');
    document.documentElement.classList.add('light');
    if (document.body) {
        document.body.classList.remove('dark');
        document.body.classList.add('light');
    }
</script>
"""

CUSTOM_CSS = """
/* Global force: pure white background and pure black text across all Gradio elements */
:root, html, body, .gradio-container, gradio-app, .dark, .light {
    background-color: #FFFFFF !important;
    color: #000000 !important;
    --body-text-color: #000000 !important;
    --block-label-text-color: #000000 !important;
    --block-title-text-color: #000000 !important;
    --body-text-color-subdued: #222222 !important;
    --background-fill-primary: #FFFFFF !important;
    --background-fill-secondary: #FFFFFF !important;
    --block-background-fill: #FFFFFF !important;
    --border-color-primary: #000000 !important;
    --input-background-fill: #FFFFFF !important;
    --input-text-color: #000000 !important;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif !important;
}

/* Ensure all headers, labels, text, and spans are pure black */
h1, h2, h3, h4, h5, h6, p, span, label, div, strong, em, b, a {
    color: #000000 !important;
}

/* Input textboxes and dropdowns */
input, textarea, select {
    background-color: #FFFFFF !important;
    color: #000000 !important;
    border: 1px solid #000000 !important;
    border-radius: 4px !important;
}

/* Blocks and containers */
.block, fieldset, .form {
    background-color: #FFFFFF !important;
    border: 1px solid #000000 !important;
    border-radius: 6px !important;
}

/* File dropboxes and download buttons */
.file-preview, .upload-container, .drop-zone, .file-item, .file-item * {
    background-color: #FFFFFF !important;
    color: #000000 !important;
    border-color: #000000 !important;
}

/* Header styling */
.header-box {
    background-color: #FFFFFF !important;
    border: 2px solid #000000 !important;
    padding: 22px !important;
    border-radius: 8px !important;
    margin-bottom: 20px !important;
}
.header-title {
    font-size: 1.8rem !important;
    font-weight: 800 !important;
    margin: 0 0 6px 0 !important;
    color: #000000 !important;
}
.header-subtitle {
    font-size: 1rem !important;
    color: #222222 !important;
    margin: 0 !important;
}

/* Hebrew Section Specific RTL Styling */
.hebrew-box {
    direction: rtl !important;
    text-align: right !important;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Tahoma, Arial, sans-serif !important;
    line-height: 1.6 !important;
}
.hebrew-box h3, .hebrew-box h4, .hebrew-box h5, .hebrew-box p, .hebrew-box li, .hebrew-box strong {
    color: #000000 !important;
}

/* High contrast Run button */
.btn-run {
    background-color: #000000 !important;
    color: #FFFFFF !important;
    font-size: 1.15rem !important;
    font-weight: 700 !important;
    border: 2px solid #000000 !important;
    border-radius: 6px !important;
    padding: 14px 28px !important;
    cursor: pointer !important;
}
.btn-run * {
    color: #FFFFFF !important;
}
.btn-run:hover {
    background-color: #222222 !important;
    color: #FFFFFF !important;
}
"""

def execute_pipeline_job(
    uploaded_files,
    strategy,
    min_len,
    max_len,
    threshold,
    min_rec,
    progress=gr.Progress(track_tqdm=False)
):
    """Executes the pipeline and produces the 4 master deliverables."""
    # 1. Parse uploaded files
    file_list = []
    if uploaded_files:
        if isinstance(uploaded_files, list):
            file_list = [f.name if hasattr(f, 'name') else str(f) for f in uploaded_files]
        else:
            file_list = [uploaded_files.name if hasattr(uploaded_files, 'name') else str(uploaded_files)]

    if not file_list:
        error_html = """
        <div style="background: #FFFFFF; border: 2px solid #000000; border-radius: 6px; padding: 16px; margin: 15px 0;">
            <p style="margin: 0; color: #000000 !important; font-weight: bold; font-size: 1rem;">
                [ERROR] No dataset provided. Please upload Excel spreadsheet(s) (.xlsx) or a .zip archive containing the dataset files.
            </p>
        </div>
        """
        return error_html, None, None, None, None, None

    def progress_callback(pct, msg):
        if progress:
            progress(pct, desc=msg)

    timestamp = time.strftime('%Y%m%d_%H%M%S')
    output_run_dir = os.path.join('Docs', f"Run_{timestamp}")

    try:
        summary, _, _, _, _ = run_user_pipeline(
            file_paths=file_list,
            folder_path=None,
            min_es_length=int(min_len),
            max_es_length=int(max_len),
            threshold=float(threshold),
            min_overlap=2,
            min_g_recurrence=int(min_rec),
            strategy=strategy,
            output_dir=output_run_dir,
            progress_callback=progress_callback
        )

        # Build clean Black & White Summary HTML
        summary_html = f"""
        <div style="background: #FFFFFF; border: 2px solid #000000; border-radius: 6px; padding: 18px; margin: 15px 0;">
            <div style="border-bottom: 2px solid #000000; padding-bottom: 8px; margin-bottom: 14px;">
                <h3 style="margin: 0; color: #000000 !important; font-size: 1.25rem; font-weight: bold;">
                    Execution Complete ({summary['duration_sec']} seconds)
                </h3>
                <p style="margin: 4px 0 0 0; color: #111111 !important; font-size: 0.9rem;">
                    Strategy: <strong>{strategy}</strong> | Cutoff: <strong>|G| &ge; {threshold}</strong>
                </p>
            </div>
            
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; margin-bottom: 14px;">
                <div style="border: 1px solid #000000; padding: 10px; border-radius: 4px; background: #FFFFFF;">
                    <div style="color: #444444 !important; font-size: 0.75rem; text-transform: uppercase; font-weight: bold;">Cohort Samples</div>
                    <div style="font-size: 1.4rem; font-weight: bold; color: #000000 !important;">{summary['total_samples']}</div>
                </div>
                <div style="border: 1px solid #000000; padding: 10px; border-radius: 4px; background: #FFFFFF;">
                    <div style="color: #444444 !important; font-size: 0.75rem; text-transform: uppercase; font-weight: bold;">Total G-Vectors</div>
                    <div style="font-size: 1.4rem; font-weight: bold; color: #000000 !important;">{summary['total_g']:,}</div>
                </div>
                <div style="border: 1px solid #000000; padding: 10px; border-radius: 4px; background: #FFFFFF;">
                    <div style="color: #444444 !important; font-size: 0.75rem; text-transform: uppercase; font-weight: bold;">Optimal ES Modules</div>
                    <div style="font-size: 1.4rem; font-weight: bold; color: #000000 !important;">{summary['es_count']}</div>
                </div>
                <div style="border: 1px solid #000000; padding: 10px; border-radius: 4px; background: #FFFFFF;">
                    <div style="color: #444444 !important; font-size: 0.75rem; text-transform: uppercase; font-weight: bold;">Universe Coverage</div>
                    <div style="font-size: 1.4rem; font-weight: bold; color: #000000 !important;">{summary['coverage_pct']}%</div>
                </div>
                <div style="border: 1px solid #000000; padding: 10px; border-radius: 4px; background: #FFFFFF;">
                    <div style="color: #444444 !important; font-size: 0.75rem; text-transform: uppercase; font-weight: bold;">Kinetic Variance</div>
                    <div style="font-size: 1.4rem; font-weight: bold; color: #000000 !important;">{summary['variance_explained_pct']}%</div>
                </div>
                <div style="border: 1px solid #000000; padding: 10px; border-radius: 4px; background: #FFFFFF;">
                    <div style="color: #444444 !important; font-size: 0.75rem; text-transform: uppercase; font-weight: bold;">Mean Optimal R²</div>
                    <div style="font-size: 1.4rem; font-weight: bold; color: #000000 !important;">{summary['avg_r2_optimal']:.4f}</div>
                </div>
                <div style="border: 1px solid #000000; padding: 10px; border-radius: 4px; background: #FFFFFF;">
                    <div style="color: #444444 !important; font-size: 0.75rem; text-transform: uppercase; font-weight: bold;">Compression Loss (ΔR²)</div>
                    <div style="font-size: 1.4rem; font-weight: bold; color: #000000 !important;">{summary['avg_delta_r2']:.4f}</div>
                </div>
                <div style="border: 1px solid #000000; padding: 10px; border-radius: 4px; background: #FFFFFF;">
                    <div style="color: #444444 !important; font-size: 0.75rem; text-transform: uppercase; font-weight: bold;">High Fidelity Rate</div>
                    <div style="font-size: 1.4rem; font-weight: bold; color: #000000 !important;">{summary['high_fidelity_pct']}%</div>
                </div>
            </div>

            <p style="margin: 0; color: #000000 !important; font-size: 0.95rem; font-weight: bold;">
                All 4 master deliverable files and the consolidated ZIP bundle are generated and available for download below.
            </p>
        </div>
        """

        return (
            summary_html,
            summary['opt_es_result_path'],
            summary['opt_signs_r2_path'],
            summary['es_combinations_csv_path'],
            summary['anchor_xlsx_path'],
            summary['zip_bundle_path']
        )

    except Exception as e:
        import traceback
        err_msg = traceback.format_exc()
        error_html = f"""
        <div style="background: #FFFFFF; border: 2px solid #000000; border-radius: 6px; padding: 16px; margin: 15px 0;">
            <p style="margin: 0 0 8px 0; color: #000000 !important; font-weight: bold; font-size: 1rem;">
                [ERROR] An error occurred during execution:
            </p>
            <pre style="margin: 0; background: #F3F4F6; color: #000000 !important; padding: 10px; border-radius: 4px; font-size: 0.85rem; overflow-x: auto;">{str(e)}\n\n{err_msg}</pre>
        </div>
        """
        return error_html, None, None, None, None, None

def build_app():
    with gr.Blocks(title="Cancer Research TME Algorithm Workbench") as demo:
        
        # Header Box
        gr.HTML("""
        <div class="header-box">
            <h1 class="header-title">Cancer Research: Tumor Microenvironment (TME) Analysis Platform</h1>
            <p class="header-subtitle">
                Surprisal Analysis Matrix Ingestion &bull; Combinatorial Exact Sign Sequence (ES) Mining &bull; Greedy Multi-Anchor Set Cover &bull; Single-Cell R² Fidelity
            </p>
        </div>
        """)

        # -------------------------------------------------------------
        # Section 1: מדריך למשתמש בעברית – הנחיות, מדדים וקבצי פלט
        # -------------------------------------------------------------
        with gr.Accordion("🇮🇱 מדריך מקיף בעברית: הנחיות שימוש, הסבר המדדים ומילון עמודות לקבצי הפלט (לחץ לפתיחה)", open=True):
            gr.HTML(r"""
            <div class="hebrew-box" style="background: #FFFFFF; border: 1px solid #000000; border-radius: 6px; padding: 18px; margin-bottom: 16px;">
                <h3 style="margin-top: 0; font-size: 1.25rem; font-weight: bold; border-bottom: 2px solid #000000; padding-bottom: 8px;">
                    1. מדריך הפעלה שלב-אחר-שלב (How to Run Without Errors)
                </h3>
                <ol style="padding-right: 20px; font-size: 0.95rem; line-height: 1.8;">
                    <li><strong>שלב 1 – הכנת הנתונים:</strong> ודא שהנתונים שלך ערוכים באחד משני הפורמטים הנתמכים:
                        <ul style="padding-right: 20px; margin: 6px 0;">
                            <li><strong>פורמט א' (מטריצות Raw - Gvalues & Lambavalues):</strong> קובץ אקסל המכיל שני גיליונות: גיליון <code>Gvalues</code> (שורות חלבונים, עמודות G0..Gk) וגיליון <code>Lambavalues</code> (שורות תאים בודדים, עמודות L0..Lk).</li>
                            <li><strong>פורמט ב' (טבלאות Tail):</strong> קובץ אקסל המכיל עמודות זנב כגון <code>G1+, G1-, G2+, G2-...</code> עם שמות החלבונים המשתתפים.</li>
                        </ul>
                    </li>
                    <li><strong>שלב 2 – טעינת הקבצים:</strong> גרור את קבצי האקסל (או קובץ ZIP המכיל את כל התיקיות) לתיבת ההעלאה מטה. המערכת קוראת את כל הקבצים וממזגת אותם אוטומטית.</li>
                    <li><strong>שלב 3 – בחירת פרמטרים:</strong> בחר באסטרטגיית העוגן להפקת קבצי המאסטר והגדר את טווח אורכי החתימות הרצוי (אורכי רצף 3 עד 8).</li>
                    <li><strong>שלב 4 – הרצה:</strong> לחץ על כפתור <code>▶ Run Algorithm & Generate Deliverables</code>.</li>
                    <li><strong>שלב 5 – הורדת התוצאות:</strong> מיד בסיום יופיעו כפתורי הורדה ל-4 קבצי המאסטר וקובץ ZIP מאוחד.</li>
                </ol>

                <h3 style="margin-top: 20px; font-size: 1.25rem; font-weight: bold; border-bottom: 2px solid #000000; padding-bottom: 8px;">
                    2. פירוט המדדים והפרמטרים הנשלטים בממשק (Execution Controls)
                </h3>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin: 12px 0; font-size: 0.9rem;">
                    <div style="border: 1px solid #000000; padding: 12px; border-radius: 4px; background: #FFFFFF;">
                        <strong>Primary Anchor Strategy (אסטרטגיית העוגן):</strong><br>
                        מגדירה את פונקציית המטרה המתמטית של פותר ה-Set Cover הגרידי עבור קבצי המאסטר:
                        <ul style="padding-right: 20px; margin: 4px 0;">
                            <li><strong>Hybrid Multi-Objective:</strong> פונקציית מטרה המאזנת בין שונות קינטית לחזרתיות ביולוגית: Score = Variance x sqrt(Count).</li>
                            <li><strong>SVD Variance Driver:</strong> בוחרת מודולים הממקסמים את סכום משקלי ה-SVD והשונות הקינטית: Score = Sum of Weights.</li>
                            <li><strong>Max Frequency:</strong> מתעדפת מודולים שמופיעים במספר הגדול ביותר של וקטורים ללא תלות במשקלם.</li>
                            <li><strong>Deep Specificity:</strong> מתעדפת קומפלקסים חלבוניים ארוכים וספציפיים (אורכים 5 עד 8).</li>
                        </ul>
                    </div>
                    <div style="border: 1px solid #000000; padding: 12px; border-radius: 4px; background: #FFFFFF;">
                        <strong>פרמטרי החיתוך והקומבינטוריקה:</strong>
                        <ul style="padding-right: 20px; margin: 4px 0;">
                            <li><strong>Min / Max Sequence Length (אורכי חתימה 3-8):</strong> טווח הגדלים של קומפלקסי החלבונים שנכרים בקומבינטוריקה.</li>
                            <li><strong>Cutoff Loading (|G| &ge; 0.05):</strong> סף המובהקות הביולוגי. חלבון שערכו המוחלט בווקטור נמוך מסף זה מסונן כרעש רקע ואינו נכלל בזנב.</li>
                            <li><strong>Min G Recurrence (חזרתיות מינימלית &ge; 2):</strong> מסנן קומבינציות בודדות שמופיעות רק בווקטור בודד אחד, ומבטיח שייבחרו אך ורק מודולים ביולוגיים יציבים שחוזרים בלפחות 2 מיקרו-סביבות.</li>
                        </ul>
                    </div>
                </div>

                <h3 style="margin-top: 20px; font-size: 1.25rem; font-weight: bold; border-bottom: 2px solid #000000; padding-bottom: 8px;">
                    3. הסבר על קבצי הפלט והמדדים המופקים (Output Deliverables & Metrics)
                </h3>
                <div style="font-size: 0.9rem; line-height: 1.7;">
                    <p style="margin: 6px 0;"><strong>מה זה Exact Sign Sequence (ES)?</strong> חתימת חלבונים בעלת קוטביות רגולטורית קשיחה: (+) מסמל ביטוי יתר, (-) מסמל דיכוי ממוקד. האלגוריתם אוכף התאמה מלאה בכיוון ישר (+) או הפוך (-) ושולל כל מודול שיש בו סתירת סימנים אפילו בחלבון בודד.</p>
                    <p style="margin: 6px 0;"><strong>Universe Coverage (%):</strong> אחוז וקטורי המיקרו-סביבה שכוסו על ידי מודולי ה-ES מתוך סך הווקטורים במאגר שהוזן.</p>
                    <p style="margin: 6px 0;"><strong>Kinetic Variance (%):</strong> אחוז האנרגיה הקינטית של סטיית הרקמה שהוסברה על ידי הרכיבים הנבחרים.</p>
                    <p style="margin: 6px 0;"><strong>Reconstruction Fidelity (R²) & Compression Loss (Delta R²):</strong> מדד למידת הדיוק של שחזור ביטוי החלבונים מתוך מודולי ה-ES המכווצים ביחס למקור ברמת התא הבודד. ירידה של Delta R2 נמוכה מ-0.05 מעידה על שחזור ברמת דיוק עליונה (High Fidelity).</p>
                    <p style="margin: 6px 0;"><strong>Conflict-Free Matrix:</strong> מטריצת הקצאה חד-חד-ערכית המבטיחה ששום וקטור ביולוגי לא ייספר פעמיים באף מודול (0 כפילויות).</p>
                    <p style="margin: 6px 0;"><strong>השוואת אסטרטגיות עוגן (Anchor_Strategies_Comparison.xlsx):</strong> הקובץ מכיל השוואה סימולטנית מלאה של כל 5 אסטרטגיות העוגן בגיליון Anchor_Strategies_Overview, וכן גיליונות פירוט ייעודיים לכל אסטרטגיה (Strat_1 עד Strat_5) המציגים את כל מודולי ה-ES שנבחרו עבורה. כל 5 האסטרטגיות מחושבות ומושוות אוטומטית בכל הרצה.</p>
                </div>
            </div>
            """)

            with gr.Row():
                template_download = gr.File(
                    value=TEMPLATE_FILE if os.path.exists(TEMPLATE_FILE) else None,
                    label="📄 1. הורדת קובץ Template לדוגמה של נתוני קלט (Sample_Input_Template.xlsx)",
                    interactive=False
                )
                dict_download = gr.File(
                    value=DICT_FILE if os.path.exists(DICT_FILE) else None,
                    label="📘 2. הורדת מילון נתונים מלא לכל עמודות קבצי הפלט (Deliverables_Data_Dictionary.xlsx)",
                    interactive=False
                )

        # -------------------------------------------------------------
        # Section 2: Instructions in English
        # -------------------------------------------------------------
        with gr.Accordion("📘 English Format Specifications & Guidelines (Click to View)", open=False):
            gr.HTML("""
            <div style="background: #FFFFFF; border: 1px solid #000000; border-radius: 6px; padding: 16px; margin-bottom: 16px;">
                <h3 style="margin-top: 0; color: #000000 !important; font-size: 1.15rem; font-weight: bold; border-bottom: 1px solid #000000; padding-bottom: 6px;">
                    Supported Input File Formats
                </h3>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin: 12px 0;">
                    <div style="border: 1px solid #000000; padding: 14px; border-radius: 4px; background: #FFFFFF;">
                        <h4 style="margin: 0 0 6px 0; font-weight: bold; color: #000000 !important;">Format A: Surprisal Matrix Workbooks</h4>
                        <p style="margin: 0 0 8px 0; font-size: 0.88rem; color: #111111 !important;">An Excel file (<code>.xlsx</code>) containing two sheets:</p>
                        <ul style="margin: 0; padding-left: 20px; font-size: 0.88rem; color: #000000 !important;">
                            <li><strong><code>Gvalues</code> Sheet:</strong> Protein loading matrix. Rows = proteins (marker column labeled <code>MARKER</code> or first text column), columns = components <code>G0, G1, G2, ..., Gk</code>.</li>
                            <li><strong><code>Lambavalues</code> Sheet:</strong> Single-cell amplitude matrix. Rows = segmented cells, columns = <code>L0, L1, L2, ..., Lk</code>.</li>
                        </ul>
                    </div>
                    <div style="border: 1px solid #000000; padding: 14px; border-radius: 4px; background: #FFFFFF;">
                        <h4 style="margin: 0 0 6px 0; font-weight: bold; color: #000000 !important;">Format B: Extracted Tail Protein Tables</h4>
                        <p style="margin: 0 0 8px 0; font-size: 0.88rem; color: #111111 !important;">An Excel file (<code>.xlsx</code>) containing precomputed tails:</p>
                        <ul style="margin: 0; padding-left: 20px; font-size: 0.88rem; color: #000000 !important;">
                            <li>Columns named <code>G1+, G1-, G2+, G2-, ..., Gk+, Gk-</code>.</li>
                            <li>Lists the proteins overexpressed (<code>+</code>) or repressed (<code>-</code>) in each component.</li>
                        </ul>
                    </div>
                </div>
            </div>
            """)

        # -------------------------------------------------------------
        # Section 3: Dataset Ingestion
        # -------------------------------------------------------------
        gr.Markdown("### 1. Ingest Dataset Files")
        upload_input = gr.File(
            label="Upload Excel File(s) (.xlsx / .xls) or .zip Archive",
            file_count="multiple",
            file_types=[".xlsx", ".xls", ".zip"]
        )

        # -------------------------------------------------------------
        # Section 4: Execution Parameters
        # -------------------------------------------------------------
        gr.Markdown("### 2. Execution Parameters")
        with gr.Row():
            strategy_select = gr.Dropdown(
                label="Primary Anchor Strategy",
                choices=[
                    "Hybrid Multi-Objective (Variance x Count)",
                    "SVD Variance Driver",
                    "Max Frequency / Universality",
                    "Deep Specificity",
                    "Biological Hallmark (Seed Co-expression)"
                ],
                value="Hybrid Multi-Objective (Variance x Count)",
                info="Mathematical objective for greedy Set Cover selection."
            )
            min_len_input = gr.Slider(minimum=3, maximum=6, value=3, step=1, label="Min Sequence Length")
            max_len_input = gr.Slider(minimum=4, maximum=10, value=8, step=1, label="Max Sequence Length")

        with gr.Row():
            threshold_input = gr.Slider(minimum=0.01, maximum=0.20, value=0.05, step=0.01, label="Cutoff Loading (|G|)")
            min_rec_input = gr.Slider(minimum=1, maximum=5, value=2, step=1, label="Min G Recurrence")

        # Primary Run Button
        run_btn = gr.Button("▶ Run Algorithm & Generate Deliverables", variant="primary", elem_classes=["btn-run"])

        # Summary Metrics Box
        summary_display = gr.HTML()

        # -------------------------------------------------------------
        # Section 5: Deliverables Downloads
        # -------------------------------------------------------------
        gr.Markdown("### 3. Download Generated Master Deliverable Files")
        with gr.Row():
            dl_opt_es = gr.File(label="📄 1. Optimal_ES_Result.xlsx")
            dl_opt_r2 = gr.File(label="📄 2. Optimal_ES_Result_Signs_R2.xlsx")
            dl_combos = gr.File(label="📊 3. ES_Combinations.csv")
            dl_anchors = gr.File(label="📈 4. Anchor_Strategies_Comparison.xlsx")
        
        with gr.Row():
            dl_zip = gr.File(label="📦 Download All Deliverables (Deliverables_Bundle.zip)")

        # Connect button callback
        run_btn.click(
            fn=execute_pipeline_job,
            inputs=[
                upload_input,
                strategy_select,
                min_len_input,
                max_len_input,
                threshold_input,
                min_rec_input
            ],
            outputs=[
                summary_display,
                dl_opt_es,
                dl_opt_r2,
                dl_combos,
                dl_anchors,
                dl_zip
            ]
        )

    return demo

create_web_application = build_app

def main():
    app = build_app()
    app.launch(
        server_name="127.0.0.1",
        server_port=7860,
        share=False,
        inbrowser=True,
        css=CUSTOM_CSS,
        head=HEAD_SCRIPT,
        theme=gr.themes.Default(primary_hue="neutral", secondary_hue="neutral", neutral_hue="neutral")
    )

if __name__ == '__main__':
    main()
