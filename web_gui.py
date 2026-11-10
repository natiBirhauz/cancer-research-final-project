"""
web_gui.py
----------
Browser-based Web GUI for the Surprisal Analysis (SA) – Minimal Exact Sign
Signatures (ES) Discovery Platform.

Key Features:
- General SA Platform framing (supports any collection of Surprisal Analysis
  datasets sharing common proteins or genes).
- Full Bilingual Support (English default, switchable to Hebrew / עברית).
- 2-Step Guided Workflow:
  Step 1: Platform Overview, Visual SA -> ES Illustration, User Guide & Templates,
          followed by a "Start Analysis" button at the bottom.
  Step 2: Data Upload -> Automatic Pre-Flight Dataset Check (counts of datasets,
          G-vectors, total & shared genes/proteins, and runtime estimate) ->
          Execution Parameters (supporting up to 50,000+ genes & sequence lengths 1,000+) ->
          Greedy Set Cover Analysis -> Interactive Visual Charts, Conflict-Free Heatmap,
          HTML Summary Report & Streamlined Excel Deliverables.
- Hover Information Tooltips (ⓘ) on all Execution Parameters, Summary Metrics,
  and Output Deliverables.
"""

import os
import time
import html
import gradio as gr
import pandas as pd

from pipeline_engine import (
    inspect_user_data,
    run_user_pipeline,
    generate_es_bar_chart_svg,
    generate_conflict_free_heatmap_html,
)

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
:root, html, body, .gradio-container, gradio-app, .dark, .light {
    background-color: #F8FAFC !important;
    color: #0F172A !important;
    --body-text-color: #0F172A !important;
    --block-label-text-color: #0F172A !important;
    --block-title-text-color: #0F172A !important;
    --body-text-color-subdued: #334155 !important;
    --background-fill-primary: #FFFFFF !important;
    --background-fill-secondary: #F8FAFC !important;
    --block-background-fill: #FFFFFF !important;
    --border-color-primary: #CBD5E1 !important;
    --input-background-fill: #FFFFFF !important;
    --input-text-color: #0F172A !important;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif !important;
}

.gradio-container {
    max-width: 1220px !important;
    margin: 0 auto !important;
    padding: 20px !important;
}

/* Header Banner */
.sa-header {
    background: linear-gradient(135deg, #0F172A 0%, #1E3A8A 100%) !important;
    border-radius: 10px;
    padding: 22px 26px;
    margin-bottom: 16px;
    color: #FFFFFF !important;
    box-shadow: 0 4px 12px rgba(15, 23, 42, 0.12);
}
.sa-header h1 {
    margin: 0 0 8px 0 !important;
    font-size: 1.65rem !important;
    font-weight: 800 !important;
    color: #FFFFFF !important;
}
.sa-header p {
    margin: 0 !important;
    font-size: 0.96rem !important;
    color: #E2E8F0 !important;
    line-height: 1.5 !important;
}

/* Hover Tooltip (ⓘ) Styling */
.info-badge {
    position: relative;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 18px;
    height: 18px;
    border-radius: 50%;
    background: #E0F2FE;
    color: #0369A1 !important;
    font-size: 12px;
    font-weight: bold;
    cursor: help;
    margin: 0 5px;
    border: 1px solid #7DD3FC;
    vertical-align: middle;
}
.info-badge:hover {
    background: #0284C7;
    color: #FFFFFF !important;
}
.info-badge .tooltip-box {
    visibility: hidden;
    opacity: 0;
    width: 270px;
    background-color: #0F172A !important;
    color: #FFFFFF !important;
    text-align: left;
    border-radius: 6px;
    padding: 9px 12px;
    position: absolute;
    z-index: 9999;
    bottom: 130%;
    left: 50%;
    transform: translateX(-50%);
    transition: opacity 0.18s ease-in-out;
    font-size: 0.8rem !important;
    font-weight: normal !important;
    line-height: 1.45 !important;
    box-shadow: 0 6px 16px rgba(0, 0, 0, 0.25);
    pointer-events: none;
    text-transform: none !important;
}
.rtl-box .info-badge .tooltip-box {
    text-align: right;
    direction: rtl;
}
.info-badge .tooltip-box::after {
    content: "";
    position: absolute;
    top: 100%;
    left: 50%;
    margin-left: -5px;
    border-width: 5px;
    border-style: solid;
    border-color: #0F172A transparent transparent transparent;
}
.info-badge:hover .tooltip-box {
    visibility: visible;
    opacity: 1;
}

/* Primary Action Buttons */
button.btn-primary-action {
    background: #1E3A8A !important;
    color: #FFFFFF !important;
    border: 2px solid #1E3A8A !important;
    font-weight: 700 !important;
    font-size: 1.08rem !important;
    padding: 12px 24px !important;
    border-radius: 8px !important;
    cursor: pointer !important;
    transition: background 0.15s ease !important;
}
button.btn-primary-action * {
    color: #FFFFFF !important;
}
button.btn-primary-action:hover {
    background: #1D4ED8 !important;
}

button.btn-secondary-nav {
    background: #FFFFFF !important;
    color: #0F172A !important;
    border: 1px solid #94A3B8 !important;
    font-weight: 600 !important;
    border-radius: 6px !important;
}

/* Parameter label header cards */
.param-header {
    font-size: 0.9rem;
    font-weight: 700;
    color: #0F172A;
    margin-bottom: 2px;
    display: flex;
    align-items: center;
}
"""


def build_illustration_svg(lang='en'):
    """
    Returns an inline SVG illustration demonstrating how multiple SA datasets
    with shared genes/proteins are compressed into a minimal set of Exact Sign Signatures (ES).
    """
    if lang == 'he':
        t_title = "המחשת האנליזה: מדאטה-סטים של Surprisal Analysis (SA) לסט חתימות ES מינימלי"
        b1_title = "1. אוסף דאטה-סטים לאחר SA"
        b1_sub = "גנים / חלבונים משותפים ווקטורי G"
        b2_title = "2. כריית חתימות וסינון סתירות"
        b2_sub = "אינווריאנטיות להיפוך סימן (ES ≡ -ES)"
        b3_title = "3. אלגוריתם Set Cover חמדן"
        b3_sub = "סט מינימלי של חתימות ES ושחזור R²"
        m_direct = "התאמה ישרה (+): אותו כיוון סימנים"
        m_inv = "התאמה הפוכה (-): כל הסימנים הפוכים"
        m_zero = "0 סתירות סימן בין גנים משותפים"
    else:
        t_title = "Analytical Overview: From Surprisal Analysis (SA) Datasets to Minimal ES Signatures"
        b1_title = "1. Collection of SA Datasets"
        b1_sub = "Shared Genes / Proteins & G-Vectors"
        b2_title = "2. Sign-Consistent ES Mining"
        b2_sub = "Sign-Inversion Invariance (ES ≡ -ES)"
        b3_title = "3. Greedy Set Cover & R²"
        b3_sub = "Minimal ES Signature Set & Allocation"
        m_direct = "Direct Match (+): Identical signs"
        m_inv = "Inverse Match (-): Opposite signs"
        m_zero = "Zero sign contradictions allowed"

    return f"""
    <div style="background:#FFFFFF; border:1px solid #CBD5E1; border-radius:10px; padding:14px; margin:14px 0; text-align:center;">
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 980 275" width="100%" style="max-width:980px; height:auto; font-family:-apple-system,BlinkMacSystemFont,Segoe UI,Roboto,sans-serif;">
        <defs>
          <marker id="arr" viewBox="0 0 10 10" refX="6" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse">
            <path d="M 0 1 L 10 5 L 0 9 z" fill="#2563EB"/>
          </marker>
        </defs>
        <text x="490" y="24" text-anchor="middle" font-size="14.5" font-weight="bold" fill="#0F172A">{t_title}</text>

        <!-- Box 1: Input SA Datasets -->
        <rect x="18" y="42" width="285" height="215" rx="8" fill="#F8FAFC" stroke="#94A3B8" stroke-width="1.5"/>
        <rect x="18" y="42" width="285" height="32" rx="8" fill="#1E3A8A"/>
        <text x="160" y="63" text-anchor="middle" font-size="12.5" font-weight="bold" fill="#FFFFFF">{b1_title}</text>
        <text x="160" y="90" text-anchor="middle" font-size="11" fill="#475569">{b1_sub}</text>

        <!-- Dataset cards inside Box 1 -->
        <rect x="32" y="102" width="256" height="42" rx="5" fill="#FFFFFF" stroke="#CBD5E1"/>
        <text x="42" y="120" font-size="10.5" font-weight="bold" fill="#0F172A">Dataset 1 (G1):</text>
        <text x="132" y="120" font-size="10" fill="#1D4ED8" font-family="monospace">GeneA(+) GeneB(-) GeneC(+)</text>
        <text x="42" y="136" font-size="10" fill="#64748B">Dataset 1 (G2): GeneD(+) GeneE(+) GeneF(-)</text>

        <rect x="32" y="152" width="256" height="42" rx="5" fill="#FFFFFF" stroke="#CBD5E1"/>
        <text x="42" y="170" font-size="10.5" font-weight="bold" fill="#0F172A">Dataset 2 (G1):</text>
        <text x="132" y="170" font-size="10" fill="#DC2626" font-family="monospace">GeneA(-) GeneB(+) GeneC(-)</text>
        <text x="42" y="186" font-size="10" fill="#64748B">Dataset 2 (G3): GeneD(+) GeneE(+) GeneG(+)</text>

        <rect x="32" y="202" width="256" height="42" rx="5" fill="#FFFFFF" stroke="#CBD5E1"/>
        <text x="42" y="220" font-size="10.5" font-weight="bold" fill="#0F172A">Dataset N (G2):</text>
        <text x="132" y="220" font-size="10" fill="#1D4ED8" font-family="monospace">GeneA(+) GeneB(-) GeneC(+)</text>
        <text x="42" y="236" font-size="10" fill="#64748B">ln(X) = G × Lᵀ (Weights w(g))</text>

        <!-- Arrow 1 -> 2 -->
        <line x1="305" y1="150" x2="342" y2="150" stroke="#2563EB" stroke-width="2.5" marker-end="url(#arr)"/>

        <!-- Box 2: Sign-Consistent Matching -->
        <rect x="348" y="42" width="284" height="215" rx="8" fill="#F8FAFC" stroke="#94A3B8" stroke-width="1.5"/>
        <rect x="348" y="42" width="284" height="32" rx="8" fill="#0284C7"/>
        <text x="490" y="63" text-anchor="middle" font-size="12.5" font-weight="bold" fill="#FFFFFF">{b2_title}</text>
        <text x="490" y="90" text-anchor="middle" font-size="11" fill="#475569">{b2_sub}</text>

        <rect x="364" y="104" width="252" height="36" rx="5" fill="#DBEAFE" stroke="#60A5FA"/>
        <text x="490" y="126" text-anchor="middle" font-size="10.5" font-weight="bold" fill="#1E40AF">{m_direct}</text>

        <rect x="364" y="148" width="252" height="36" rx="5" fill="#FEE2E2" stroke="#F87171"/>
        <text x="490" y="170" text-anchor="middle" font-size="10.5" font-weight="bold" fill="#991B1B">{m_inv}</text>

        <rect x="364" y="192" width="252" height="50" rx="5" fill="#F1F5F9" stroke="#CBD5E1"/>
        <text x="490" y="213" text-anchor="middle" font-size="10.5" font-weight="bold" fill="#0F172A">{m_zero}</text>
        <text x="490" y="231" text-anchor="middle" font-size="10" fill="#0369A1" font-family="monospace">Score(ES) = Σ w(g) × √|New_Gs|</text>

        <!-- Arrow 2 -> 3 -->
        <line x1="634" y1="150" x2="671" y2="150" stroke="#2563EB" stroke-width="2.5" marker-end="url(#arr)"/>

        <!-- Box 3: Minimal ES Set Cover & Heatmap -->
        <rect x="677" y="42" width="285" height="215" rx="8" fill="#F8FAFC" stroke="#94A3B8" stroke-width="1.5"/>
        <rect x="677" y="42" width="285" height="32" rx="8" fill="#059669"/>
        <text x="820" y="63" text-anchor="middle" font-size="12.5" font-weight="bold" fill="#FFFFFF">{b3_title}</text>
        <text x="820" y="90" text-anchor="middle" font-size="11" fill="#475569">{b3_sub}</text>

        <rect x="692" y="102" width="255" height="46" rx="5" fill="#ECFDF5" stroke="#6EE7B7"/>
        <text x="702" y="120" font-size="10.5" font-weight="bold" fill="#065F46">ES_1: GeneA(+) | GeneB(-) | GeneC(+)</text>
        <text x="702" y="138" font-size="10" font-weight="bold" fill="#065F46">ES_2: GeneD(+) | GeneE(+)</text>

        <!-- Mini Matrix representation -->
        <rect x="692" y="156" width="255" height="88" rx="5" fill="#FFFFFF" stroke="#CBD5E1"/>
        <text x="820" y="172" text-anchor="middle" font-size="10" font-weight="bold" fill="#0F172A">Conflict-Free Matrix &amp; R² Fidelity</text>
        <rect x="705" y="180" width="70" height="22" fill="#F1F5F9" stroke="#CBD5E1"/><text x="740" y="195" text-anchor="middle" font-size="9.5" fill="#0F172A">Dataset 1</text>
        <rect x="775" y="180" width="80" height="22" fill="#DBEAFE" stroke="#CBD5E1"/><text x="815" y="195" text-anchor="middle" font-size="9.5" font-weight="bold" fill="#1E40AF">ES_1: G1(+)</text>
        <rect x="855" y="180" width="80" height="22" fill="#DBEAFE" stroke="#CBD5E1"/><text x="895" y="195" text-anchor="middle" font-size="9.5" font-weight="bold" fill="#1E40AF">ES_2: G2(+)</text>

        <rect x="705" y="202" width="70" height="22" fill="#F1F5F9" stroke="#CBD5E1"/><text x="740" y="217" text-anchor="middle" font-size="9.5" fill="#0F172A">Dataset 2</text>
        <rect x="775" y="202" width="80" height="22" fill="#FEE2E2" stroke="#CBD5E1"/><text x="815" y="217" text-anchor="middle" font-size="9.5" font-weight="bold" fill="#991B1B">ES_1: G1(-)</text>
        <rect x="855" y="202" width="80" height="22" fill="#DBEAFE" stroke="#CBD5E1"/><text x="895" y="217" text-anchor="middle" font-size="9.5" font-weight="bold" fill="#1E40AF">ES_2: G3(+)</text>
      </svg>
    </div>
    """


def info_icon(tooltip_text):
    """Renders an ⓘ icon with a CSS hover popup."""
    safe_tip = html.escape(tooltip_text)
    return f'<span class="info-badge">i<span class="tooltip-box">{safe_tip}</span></span>'


def get_header_html(lang='en'):
    if lang == 'he':
        return """
        <div class="sa-header" style="direction:rtl; text-align:right;">
            <h1>פלטפורמת Surprisal Analysis (SA) לזיהוי סט מינימלי של חתימות (ES)</h1>
            <p>
                פלטפורמה חישובית המקבלת כל אוסף של דאטה-סטים שעברו Surprisal Analysis (SA) ויש להם חלבונים או גנים משותפים,
                ומוצאת סט מינימלי ונקי מסתירות של חתימות כיווניות (Exact Sign Signatures - ES) המתאר ומשחזר את וקטורי ה-G שלהם.
            </p>
        </div>
        """
    return """
    <div class="sa-header">
        <h1>Surprisal Analysis (SA) &ndash; Minimal Exact Sign Signatures (ES) Discovery Platform</h1>
        <p>
            A general computational platform for any collection of datasets analyzed via Surprisal Analysis (SA) that share common proteins or genes.
            Identifies a minimal, conflict-free set of Exact Sign Signatures (ES) that describes and reconstructs their G-vectors with high R² fidelity.
        </p>
    </div>
    """


def get_step1_guide_html(lang='en'):
    svg_illust = build_illustration_svg(lang=lang)
    if lang == 'he':
        return f"""
        <div class="rtl-box" style="direction:rtl; text-align:right; background:#FFFFFF; border:1px solid #CBD5E1; border-radius:10px; padding:22px; margin-bottom:14px;">
            <h2 style="margin-top:0; color:#0F172A; border-bottom:2px solid #0F172A; padding-bottom:8px;">
                שלב 1: הכרת הפלטפורמה, איור מתודולוגי ומדריך למשתמש
            </h2>
            <p style="font-size:0.96rem; line-height:1.7; color:#1E293B;">
                פלטפורמה זו נועדה לחוקרים המנתחים אוסף של דאטה-סטים (חלבונים, גנים, טרנסקריפטומיקה, פרוטאומיקה או תאים בודדים) שעברו פירוק
                <strong>Surprisal Analysis (SA)</strong>. בכל דאטה-סט קיימים וקטורי תהליכים כיווניים (<strong>G1, G2, ..., Gk</strong>) שבהם גנים או חלבונים
                בעלי משקל משמעותי מקבלים סימן פלוס (<code>+</code>) או מינוס (<code>-</code>).
                האלגוריתם מאתר באופן אוטומטי את <strong>הקבוצה המינימלית של חתימות סימנים מדויקות (Exact Sign Signatures - ES)</strong> החוזרות בין הדאטה-סטים
                ללא סתירות סימן, ומאפשר לתאר ולשחזר את כלל הדאטה-סטים בדיוק רב (מדד <strong>R²</strong>).
            </p>

            {svg_illust}

            <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-top:16px;">
                <div style="border:1px solid #CBD5E1; border-radius:8px; padding:14px; background:#F8FAFC;">
                    <h3 style="margin:0 0 8px 0; font-size:1.05rem; color:#0F172A;">📂 אילו קבצי קלט נתמכים?</h3>
                    <ul style="padding-right:18px; margin:0; font-size:0.9rem; line-height:1.65;">
                        <li><strong>פורמט א' (מטריצות SA מלאות – Gvalues &amp; Lambavalues):</strong> קובץ Excel לכל דאטה-סט הכולל גיליון <code>Gvalues</code> (שורות = חלבונים/גנים, עמודות = <code>G0..Gk</code>) וגיליון <code>Lambavalues</code> (עמודות = <code>L0..Lk</code>). מאפשר חישוב מלא של משקלי השונות ומדד השחזור R².</li>
                        <li><strong>פורמט ב' (טבלאות זנבות – Tail Tables):</strong> קובץ Excel עם עמודות <code>G1+, G1-, G2+, G2-...</code> המפרטות את הגנים/חלבונים הפעילים בכל זנב.</li>
                        <li>ניתן להעלות מספר קבצי Excel יחד או קובץ <code>.zip</code> המכיל תיקיות של דאטה-סטים.</li>
                    </ul>
                </div>
                <div style="border:1px solid #CBD5E1; border-radius:8px; padding:14px; background:#F8FAFC;">
                    <h3 style="margin:0 0 8px 0; font-size:1.05rem; color:#0F172A;">🚀 מהלך העבודה בשלב הבא (שלב 2)</h3>
                    <ol style="padding-right:18px; margin:0; font-size:0.9rem; line-height:1.65;">
                        <li><strong>העלאת הקבצים ובדיקה מקדימה אוטומטית (Pre-Check):</strong> מיד עם העלאת הקבצים תוצג סקירה של מספר הדאטה-סטים, וקטורי ה-G, מספר החלבונים/גנים הכולל והמשותף, והערכת זמן ריצה.</li>
                        <li><strong>הגדרת פרמטרים והרצה:</strong> התאמת אורך החתימה המבוקש (תומך גם בדאטה-סטים קטנים וגם בדאטה-סטים ענקיים של עד 20,000+ גנים ואורך חתימה של 1,000+) והרצת אלגוריתם ה-Set Cover החמדן.</li>
                        <li><strong>תוצאות ויזואליות ודוחות:</strong> צפייה בגרף העמודות של חתימות ה-ES, מפת חום (Heatmap) של מטריצת ההקצאה, והורדת דוח HTML מסכם וקובץ Excel מרוכז.</li>
                    </ol>
                </div>
            </div>
        </div>
        """

    return f"""
    <div style="background:#FFFFFF; border:1px solid #CBD5E1; border-radius:10px; padding:22px; margin-bottom:14px;">
        <h2 style="margin-top:0; color:#0F172A; border-bottom:2px solid #0F172A; padding-bottom:8px;">
            Step 1: Platform Overview, Visual Illustration &amp; Methodology Guide
        </h2>
        <p style="font-size:0.96rem; line-height:1.7; color:#1E293B;">
            This platform accepts any collection of datasets analyzed via <strong>Surprisal Analysis (SA)</strong> that share common proteins or genes
            (e.g., proteomics, transcriptomics/RNA-seq, single-cell or bulk cohorts). Each dataset contains unbalanced biological process vectors
            (<strong>G1, G2, ..., Gk</strong>) where active features carry positive (<code>+</code>) or negative (<code>-</code>) signs.
            The greedy Set Cover engine identifies the <strong>minimal, conflict-free set of Exact Sign Signatures (ES)</strong> that recurringly spans
            the G-vectors across datasets and evaluates reconstruction quality via the determination coefficient (<strong>R²</strong>).
        </p>

        {svg_illust}

        <div style="display:grid; grid-template-columns:1fr 1fr; gap:16px; margin-top:16px;">
            <div style="border:1px solid #CBD5E1; border-radius:8px; padding:14px; background:#F8FAFC;">
                <h3 style="margin:0 0 8px 0; font-size:1.05rem; color:#0F172A;">📂 Supported Input Formats</h3>
                <ul style="padding-left:18px; margin:0; font-size:0.9rem; line-height:1.65;">
                    <li><strong>Format A (Full SA Matrix Workbooks &ndash; Gvalues &amp; Lambavalues):</strong> Excel file(s) containing a <code>Gvalues</code> sheet (rows = genes/proteins, columns = <code>G0..Gk</code>) and a <code>Lambavalues</code> sheet (columns = <code>L0..Lk</code>). Enables exact variance weighting and R² reconstruction.</li>
                    <li><strong>Format B (Extracted Tail Tables):</strong> Excel file(s) with columns named <code>G1+, G1-, G2+, G2-...</code> listing active tail genes/proteins.</li>
                    <li>Upload multiple <code>.xlsx</code> files directly or a single <code>.zip</code> archive containing your dataset folders.</li>
                </ul>
            </div>
            <div style="border:1px solid #CBD5E1; border-radius:8px; padding:14px; background:#F8FAFC;">
                <h3 style="margin:0 0 8px 0; font-size:1.05rem; color:#0F172A;">🚀 Workflow in Step 2 (Analysis Workbench)</h3>
                <ol style="padding-left:18px; margin:0; font-size:0.9rem; line-height:1.65;">
                    <li><strong>Upload &amp; Pre-Flight Dataset Check:</strong> Automatically inspects uploaded files and reports detected datasets, G-vectors, total &amp; shared genes/proteins across datasets, and estimated runtime before execution.</li>
                    <li><strong>Configure Parameters &amp; Run:</strong> Supports both compact panels (lengths 3&ndash;10) and high-dimensional omics (e.g., 20,000 genes with Min Sequence Length = 1000).</li>
                    <li><strong>Interactive Visualizations &amp; Streamlined Deliverables:</strong> View ES coverage bar charts, Conflict-Free Allocation Heatmaps, download the standalone HTML Summary Report, and export the streamlined Master Excel workbook.</li>
                </ol>
            </div>
        </div>
    </div>
    """


def get_param_headers_html(lang='en'):
    """Returns 4 HTML labels with ⓘ hover tooltips for the Execution Parameters."""
    if lang == 'he':
        p1 = f'<div class="param-header rtl-box" style="direction:rtl;">אורך חתימה מינימלי (Min Sequence Length) {info_icon("מספר החלבונים או הגנים המינימלי שייכלל בכל חתימת ES. לדאטה-סטים קטנים מקובל 3, ולדאטה-סטים גדולים (כגון 20,000 גנים) ניתן להגדיר 100, 500 או 1,000+.")}</div>'
        p2 = f'<div class="param-header rtl-box" style="direction:rtl;">אורך חתימה מקסימלי (Max Sequence Length) {info_icon("מספר החלבונים או הגנים המקסימלי בחתימת ES בודדת. האלגוריתם יחפש חתימות בטווח שבין המינימום למקסימום.")}</div>'
        p3 = f'<div class="param-header rtl-box" style="direction:rtl;">סף מובהקות במשקל וקטור (Cutoff Loading |G|) {info_icon("סף סינון לחלבונים/גנים בווקטורי ה-G (במטריצות Gvalues). רק גנים/חלבונים שערכם המוחלט |G| גדול או שווה לסף זה ייחשבו פעילים בזנב (+ או -).")}</div>'
        p4 = f'<div class="param-header rtl-box" style="direction:rtl;">חזרתיות מינימלית בווקטורי G (Min G Recurrence) {info_icon("מספר וקטורי ה-G המינימלי שבהם חתימת ES חייבת להופיע כדי להיחשב מועמדת חוקית (ברירת מחדל: 2 ומעלה, לסינון רעש נקודתי שמופיע בווקטור יחיד).")}</div>'
    else:
        p1 = f'<div class="param-header">Min Sequence Length {info_icon("Minimum number of genes/proteins required in an Exact Sign Signature (ES). Use 3–5 for targeted panels, or 100–1000+ for high-dimensional transcriptomic/proteomic datasets (e.g., 20,000 genes).")}</div>'
        p2 = f'<div class="param-header">Max Sequence Length {info_icon("Maximum number of genes/proteins allowed in a single candidate ES signature. The engine mines recurring signatures between Min and Max length.")}</div>'
        p3 = f'<div class="param-header">Cutoff Loading (|G|) {info_icon("Significance threshold applied to raw Gvalues matrices. Only genes/proteins with absolute loading |G| >= threshold are extracted into the active G+ or G- tails.")}</div>'
        p4 = f'<div class="param-header">Min G Recurrence {info_icon("Minimum number of distinct G-vectors a signature must match (with zero contradictions) to be eligible for greedy Set Cover selection (default >= 2 filters out singletons).")}</div>'
    return p1, p2, p3, p4


def get_deliverables_info_html(lang='en'):
    """Returns explanatory header with ⓘ tooltips for each streamlined output deliverable."""
    if lang == 'he':
        return f"""
        <div class="rtl-box" style="direction:rtl; text-align:right; background:#FFFFFF; border:1px solid #CBD5E1; border-radius:8px; padding:14px; margin:12px 0;">
            <div style="font-weight:bold; font-size:1.02rem; margin-bottom:8px; color:#0F172A;">
                📥 קבצי הפלט המסכמים להורדה (לחץ או עבור עם העכבר על ⓘ להסבר על כל קובץ):
            </div>
            <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(230px, 1fr)); gap:10px; font-size:0.86rem;">
                <div style="border:1px solid #E2E8F0; padding:8px 10px; border-radius:6px; background:#F8FAFC;">
                    <strong>1. Optimal_ES_Result.xlsx</strong> {info_icon("קובץ האקסל המרכזי המאוחד (5 גיליונות ממוקדים ללא כפילויות): Summary, Optimal_ES_SetCover, ES_Conflict_Free_Matrix, Per_Dataset_Coverage_and_R2, ו-Per_G_Coverage_Detail.")}
                </div>
                <div style="border:1px solid #E2E8F0; padding:8px 10px; border-radius:6px; background:#F8FAFC;">
                    <strong>2. Analysis_Summary_Report.html</strong> {info_icon("דוח HTML חזותי עצמאי המסכם את הריצה: כולל את פרמטרי ההרצה, מדדי הכיסוי וה-R², גרף עמודות של ה-ES, מפת חום (Heatmap) ואת הטבלאות המרכזיות.")}
                </div>
                <div style="border:1px solid #E2E8F0; padding:8px 10px; border-radius:6px; background:#F8FAFC;">
                    <strong>3. Deliverables_Bundle.zip</strong> {info_icon("ארכיון ZIP מרוכז הכולל את כל קבצי הפלט והדוח המסכם להורדה בלחיצה אחת.")}
                </div>
                <div style="border:1px solid #E2E8F0; padding:8px 10px; border-radius:6px; background:#F8FAFC;">
                    <strong>4. קבצים משלימים (אופציונלי)</strong> {info_icon("Anchor_Strategies_Comparison.xlsx (גיליון השוואה בין 5 אסטרטגיות עוגן) ו-ES_Combinations.csv (רשימת הצירופים החוזרים ב-2 וקטורים ומעלה, ללא רעש של וקטור בודד).")}
                </div>
            </div>
        </div>
        """
    return f"""
    <div style="background:#FFFFFF; border:1px solid #CBD5E1; border-radius:8px; padding:14px; margin:12px 0;">
        <div style="font-weight:bold; font-size:1.02rem; margin-bottom:8px; color:#0F172A;">
            📥 Streamlined Output Deliverables (Hover over ⓘ to inspect what each file contains):
        </div>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(230px, 1fr)); gap:10px; font-size:0.86rem;">
            <div style="border:1px solid #E2E8F0; padding:8px 10px; border-radius:6px; background:#F8FAFC;">
                <strong>1. Optimal_ES_Result.xlsx</strong> {info_icon("Unified Master Excel Workbook with 5 non-redundant sheets: Summary, Optimal_ES_SetCover, ES_Conflict_Free_Matrix, Per_Dataset_Coverage_and_R2 (merged), and Per_G_Coverage_Detail.")}
            </div>
            <div style="border:1px solid #E2E8F0; padding:8px 10px; border-radius:6px; background:#F8FAFC;">
                <strong>2. Analysis_Summary_Report.html</strong> {info_icon("Standalone visual HTML report containing execution parameters, KPI summary cards, interactive ES bar chart, Conflict-Free Heatmap, and core tables.")}
            </div>
            <div style="border:1px solid #E2E8F0; padding:8px 10px; border-radius:6px; background:#F8FAFC;">
                <strong>3. Deliverables_Bundle.zip</strong> {info_icon("Consolidated ZIP archive containing the Master Excel workbook, HTML summary report, and supplementary tables.")}
            </div>
            <div style="border:1px solid #E2E8F0; padding:8px 10px; border-radius:6px; background:#F8FAFC;">
                <strong>4. Advanced / Optional Files</strong> {info_icon("Anchor_Strategies_Comparison.xlsx (benchmark table comparing 5 greedy anchor objectives) and ES_Combinations.csv (filtered to recurring signatures in 2+ G-vectors).")}
            </div>
        </div>
    </div>
    """


def run_precheck_job(uploaded_files, threshold, lang='en'):
    """
    Runs immediately after file upload (before running the main greedy algorithm)
    to show:
    1. Number of datasets, G-vectors, and unique genes/proteins detected.
    2. Number of genes/proteins shared across all datasets (and across 2+ datasets).
    3. Estimated runtime and tail size distribution.
    """
    if not uploaded_files:
        if lang == 'he':
            return '<div class="rtl-box" style="direction:rtl; text-align:right; padding:12px; border:1px dashed #94A3B8; border-radius:8px; background:#F8FAFC; color:#475569;">העלה קבצי Excel (.xlsx/.xls) או ארכיון .zip כדי לבצע בדיקה מקדימה אוטומטית של הדאטה-סטים לפני ההרצה.</div>'
        return '<div style="padding:12px; border:1px dashed #94A3B8; border-radius:8px; background:#F8FAFC; color:#475569;">Upload Excel file(s) (.xlsx/.xls) or a .zip archive above to run the automatic Pre-Flight Dataset Check prior to analysis.</div>'

    file_list = [f.name if hasattr(f, 'name') else str(f) for f in (uploaded_files if isinstance(uploaded_files, list) else [uploaded_files])]

    try:
        info = inspect_user_data(file_paths=file_list, threshold=float(threshold))
        if lang == 'he':
            return f"""
            <div class="rtl-box" style="direction:rtl; text-align:right; background:#F0FDF4; border:2px solid #16A34A; border-radius:10px; padding:16px; margin:10px 0;">
                <div style="font-weight:800; font-size:1.08rem; color:#14532D; margin-bottom:10px;">
                    ✅ בדיקה מקדימה הושלמה בהצלחה (Pre-Flight Dataset Check)
                </div>
                <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(185px, 1fr)); gap:12px;">
                    <div style="background:#FFFFFF; border:1px solid #BBF7D0; border-radius:8px; padding:10px;">
                        <div style="font-size:0.78rem; color:#475569; font-weight:bold;">1. דאטה-סטים ווקטורים שזוהו {info_icon("מספר הדאטה-סטים התקינים שנקראו ומספר וקטורי ה-G הפעילים הכולל בהם.")}</div>
                        <div style="font-size:1.25rem; font-weight:800; color:#0F172A; margin-top:4px;">{info['num_datasets']} דאטה-סטים | {info['total_g_vectors']:,} וקטורי G</div>
                    </div>
                    <div style="background:#FFFFFF; border:1px solid #BBF7D0; border-radius:8px; padding:10px;">
                        <div style="font-size:0.78rem; color:#475569; font-weight:bold;">סה"כ חלבונים / גנים ייחודיים {info_icon("מספר החלבונים או הגנים הייחודיים הכולל שזוהו על פני כל הדאטה-סטים שהועלו.")}</div>
                        <div style="font-size:1.25rem; font-weight:800; color:#0F172A; margin-top:4px;">{info['total_unique_features']:,} (פעילים בזנבות: {info['active_tail_features']:,})</div>
                    </div>
                    <div style="background:#FFFFFF; border:1px solid #BBF7D0; border-radius:8px; padding:10px;">
                        <div style="font-size:0.78rem; color:#475569; font-weight:bold;">2. חלבונים / גנים משותפים {info_icon("כמה חלבונים/גנים משותפים לכל הדאטה-סטים (100%), וכמה משותפים ללפחות 2 דאטה-סטים.")}</div>
                        <div style="font-size:1.25rem; font-weight:800; color:#0369A1; margin-top:4px;">{info['shared_all_datasets']:,} בכל הדאטה-סטים ({info['shared_2plus_datasets']:,} ב-2+)</div>
                    </div>
                    <div style="background:#FFFFFF; border:1px solid #BBF7D0; border-radius:8px; padding:10px;">
                        <div style="font-size:0.78rem; color:#475569; font-weight:bold;">גודל זנב וקטור G (מינ' / חציון / מקס') {info_icon("מספר החלבונים/גנים הפעילים בכל וקטור G. מסייע בבחירת טווח אורכי החתימות (Min/Max Sequence Length).")}</div>
                        <div style="font-size:1.2rem; font-weight:800; color:#0F172A; margin-top:4px;">{info['min_tail_len']} / {info['median_tail_len']} / {info['max_tail_len']}</div>
                    </div>
                    <div style="background:#FFFFFF; border:1px solid #BBF7D0; border-radius:8px; padding:10px;">
                        <div style="font-size:0.78rem; color:#475569; font-weight:bold;">3. הערכת זמן ריצה צפוי {info_icon("הערכת זמן הריצה של אלגוריתם הכרייה וה-Greedy Set Cover על בסיס כמות הדאטה-סטים ווקטורי ה-G.")}</div>
                        <div style="font-size:1.25rem; font-weight:800; color:#15803D; margin-top:4px;">~{info['estimated_time_sec']} שניות</div>
                    </div>
                </div>
            </div>
            """
        return f"""
        <div style="background:#F0FDF4; border:2px solid #16A34A; border-radius:10px; padding:16px; margin:10px 0;">
            <div style="font-weight:800; font-size:1.08rem; color:#14532D; margin-bottom:10px;">
                ✅ Pre-Flight Dataset Inspection Complete (Ready for Analysis)
            </div>
            <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(185px, 1fr)); gap:12px;">
                <div style="background:#FFFFFF; border:1px solid #BBF7D0; border-radius:8px; padding:10px;">
                    <div style="font-size:0.78rem; color:#475569; font-weight:bold;">1. Detected Datasets &amp; G-Vectors {info_icon("Total valid SA datasets parsed and the total number of active G-vectors extracted across them.")}</div>
                    <div style="font-size:1.25rem; font-weight:800; color:#0F172A; margin-top:4px;">{info['num_datasets']} Datasets | {info['total_g_vectors']:,} Gs</div>
                </div>
                <div style="background:#FFFFFF; border:1px solid #BBF7D0; border-radius:8px; padding:10px;">
                    <div style="font-size:0.78rem; color:#475569; font-weight:bold;">Total Unique Genes / Proteins {info_icon("Total distinct genes/proteins detected across the uploaded datasets, and how many participate in active tails (|G| >= cutoff).")}</div>
                    <div style="font-size:1.25rem; font-weight:800; color:#0F172A; margin-top:4px;">{info['total_unique_features']:,} ({info['active_tail_features']:,} active)</div>
                </div>
                <div style="background:#FFFFFF; border:1px solid #BBF7D0; border-radius:8px; padding:10px;">
                    <div style="font-size:0.78rem; color:#475569; font-weight:bold;">2. Shared Genes / Proteins {info_icon("Number of genes/proteins present across ALL uploaded datasets (100% intersection), and across at least 2 datasets.")}</div>
                    <div style="font-size:1.25rem; font-weight:800; color:#0369A1; margin-top:4px;">{info['shared_all_datasets']:,} in All ({info['shared_2plus_datasets']:,} in 2+)</div>
                </div>
                <div style="background:#FFFFFF; border:1px solid #BBF7D0; border-radius:8px; padding:10px;">
                    <div style="font-size:0.78rem; color:#475569; font-weight:bold;">G-Tail Size (Min / Median / Max) {info_icon("Distribution of active genes/proteins per G-vector. Useful for choosing Min & Max Sequence Length.")}</div>
                    <div style="font-size:1.2rem; font-weight:800; color:#0F172A; margin-top:4px;">{info['min_tail_len']} / {info['median_tail_len']} / {info['max_tail_len']}</div>
                </div>
                <div style="background:#FFFFFF; border:1px solid #BBF7D0; border-radius:8px; padding:10px;">
                    <div style="font-size:0.78rem; color:#475569; font-weight:bold;">3. Estimated Runtime {info_icon("Estimated wall-clock time to complete candidate ES mining, greedy Set Cover optimization, and R² evaluation.")}</div>
                    <div style="font-size:1.25rem; font-weight:800; color:#15803D; margin-top:4px;">~{info['estimated_time_sec']} sec</div>
                </div>
            </div>
        </div>
        """
    except Exception as e:
        return f"""
        <div style="background:#FEF2F2; border:2px solid #DC2626; border-radius:8px; padding:14px; color:#991B1B; font-weight:bold;">
            [Pre-Check Warning] {html.escape(str(e))}
        </div>
        """


def format_summary_cards_html(summary, min_len, max_len, threshold, lang='en'):
    """Builds the KPI summary cards with ⓘ hover tooltips on every metric."""
    if lang == 'he':
        return f"""
        <div class="rtl-box" style="direction:rtl; text-align:right; background:#FFFFFF; border:2px solid #0F172A; border-radius:10px; padding:18px; margin:14px 0;">
            <div style="border-bottom:2px solid #CBD5E1; padding-bottom:8px; margin-bottom:14px;">
                <h3 style="margin:0; color:#0F172A; font-size:1.25rem; font-weight:800;">
                    האנליזה הושלמה בהצלחה ({summary['duration_sec']} שניות)
                </h3>
                <p style="margin:4px 0 0 0; color:#334155; font-size:0.88rem;">
                    מצב כרייה: <strong>{summary.get('mining_mode', 'Standard')}</strong> | טווח אורך חתימה: <strong>{int(min_len)}&ndash;{int(max_len)}</strong> | סף |G| &ge; <strong>{threshold}</strong>
                </p>
            </div>
            <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(165px, 1fr)); gap:12px;">
                <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                    <div style="font-size:0.76rem; font-weight:bold; color:#475569;">דאטה-סטים שנותחו {info_icon("מספר הדאטה-סטים של Surprisal Analysis שעובדו בריצה זו.")}</div>
                    <div style="font-size:1.4rem; font-weight:800; color:#0F172A;">{summary['total_samples']}</div>
                </div>
                <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                    <div style="font-size:0.76rem; font-weight:bold; color:#475569;">סה"כ וקטורי G {info_icon("סך כל וקטורי התהליכים (G1..Gk) שחולצו מכל הדאטה-סטים.")}</div>
                    <div style="font-size:1.4rem; font-weight:800; color:#0F172A;">{summary['total_g']:,}</div>
                </div>
                <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                    <div style="font-size:0.76rem; font-weight:bold; color:#475569;">חתימות ES מינימליות {info_icon("גודל הסט המינימלי של חתימות סימנים מדויקות (ES) שנבחר על ידי האלגוריתם החמדן לכיסוי הווקטורים.")}</div>
                    <div style="font-size:1.4rem; font-weight:800; color:#1D4ED8;">{summary['es_count']}</div>
                </div>
                <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                    <div style="font-size:0.76rem; font-weight:bold; color:#475569;">כיסוי וקטורי G {info_icon("אחוז וקטורי ה-G בכלל הדאטה-סטים שכוסו ללא סתירות על ידי חתימות ה-ES שנבחרו.")}</div>
                    <div style="font-size:1.4rem; font-weight:800; color:#0F172A;">{summary['coverage_pct']}%</div>
                </div>
                <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                    <div style="font-size:0.76rem; font-weight:bold; color:#475569;">שונות מוסברת (Variance) {info_icon("אחוז האנרגיה/השונות המשוקללת של וקטורי ה-G שמכוסה על ידי חתימות ה-ES שנבחרו.")}</div>
                    <div style="font-size:1.4rem; font-weight:800; color:#0F172A;">{summary['variance_explained_pct']}%</div>
                </div>
                <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                    <div style="font-size:0.76rem; font-weight:bold; color:#475569;">ממוצע R² אופטימלי {info_icon("מקדם ההסבר (R²) הממוצע בשחזור הדאטה המקורי תוך שימוש בווקטורי ה-G המכוסים על ידי ה-ES בלבד.")}</div>
                    <div style="font-size:1.4rem; font-weight:800; color:#0F172A;">{summary['avg_r2_optimal']:.4f}</div>
                </div>
                <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                    <div style="font-size:0.76rem; font-weight:bold; color:#475569;">אובדן דחיסה (ΔR²) {info_icon("ההפרש בין R² מלא (כל הרכיבים) לבין R² אופטימלי. ערך נמוך מ-0.05 מעיד על דחיסה ללא אובדן מידע מהותי.")}</div>
                    <div style="font-size:1.4rem; font-weight:800; color:#0F172A;">{summary['avg_delta_r2']:.4f}</div>
                </div>
                <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                    <div style="font-size:0.76rem; font-weight:bold; color:#475569;">שיעור שחזור באיכות גבוהה {info_icon("אחוז הדאטה-סטים שבהם אובדן השחזור (ΔR²) קטן מ-0.05 (High Fidelity).")}</div>
                    <div style="font-size:1.4rem; font-weight:800; color:#15803D;">{summary['high_fidelity_pct']}%</div>
                </div>
            </div>
        </div>
        """

    return f"""
    <div style="background:#FFFFFF; border:2px solid #0F172A; border-radius:10px; padding:18px; margin:14px 0;">
        <div style="border-bottom:2px solid #CBD5E1; padding-bottom:8px; margin-bottom:14px;">
            <h3 style="margin:0; color:#0F172A; font-size:1.25rem; font-weight:800;">
                Execution Complete ({summary['duration_sec']} seconds)
            </h3>
            <p style="margin:4px 0 0 0; color:#334155; font-size:0.88rem;">
                Mining Engine: <strong>{summary.get('mining_mode', 'Standard')}</strong> | Sequence Length: <strong>{int(min_len)}&ndash;{int(max_len)}</strong> | Cutoff: <strong>|G| &ge; {threshold}</strong>
            </p>
        </div>
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(165px, 1fr)); gap:12px;">
            <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                <div style="font-size:0.75rem; font-weight:bold; color:#475569;">DATASETS ANALYZED {info_icon("Total number of Surprisal Analysis datasets processed in this run.")}</div>
                <div style="font-size:1.4rem; font-weight:800; color:#0F172A;">{summary['total_samples']}</div>
            </div>
            <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                <div style="font-size:0.75rem; font-weight:bold; color:#475569;">TOTAL G-VECTORS {info_icon("Total unbalanced biological process vectors (G1..Gk) across all datasets.")}</div>
                <div style="font-size:1.4rem; font-weight:800; color:#0F172A;">{summary['total_g']:,}</div>
            </div>
            <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                <div style="font-size:0.75rem; font-weight:bold; color:#475569;">MINIMAL ES SIGNATURES {info_icon("Number of Exact Sign Signatures (ES) selected by the greedy Set Cover solver to span the G-vectors.")}</div>
                <div style="font-size:1.4rem; font-weight:800; color:#1D4ED8;">{summary['es_count']}</div>
            </div>
            <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                <div style="font-size:0.75rem; font-weight:bold; color:#475569;">G-VECTOR COVERAGE {info_icon("Percentage of all G-vectors across datasets covered with zero sign contradictions by the selected ES set.")}</div>
                <div style="font-size:1.4rem; font-weight:800; color:#0F172A;">{summary['coverage_pct']}%</div>
            </div>
            <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                <div style="font-size:0.75rem; font-weight:bold; color:#475569;">EXPLAINED VARIANCE {info_icon("Percentage of total deviation energy w(g) captured by the covered G-vectors.")}</div>
                <div style="font-size:1.4rem; font-weight:800; color:#0F172A;">{summary['variance_explained_pct']}%</div>
            </div>
            <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                <div style="font-size:0.75rem; font-weight:bold; color:#475569;">MEAN OPTIMAL R² {info_icon("Mean determination coefficient (R²) when reconstructing the datasets using the selected ES-covered G-vectors.")}</div>
                <div style="font-size:1.4rem; font-weight:800; color:#0F172A;">{summary['avg_r2_optimal']:.4f}</div>
            </div>
            <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                <div style="font-size:0.75rem; font-weight:bold; color:#475569;">COMPRESSION LOSS (ΔR²) {info_icon("Mean difference between full R² (all components) and optimal ES R². Values < 0.05 indicate near-lossless compression.")}</div>
                <div style="font-size:1.4rem; font-weight:800; color:#0F172A;">{summary['avg_delta_r2']:.4f}</div>
            </div>
            <div style="border:1px solid #CBD5E1; padding:10px; border-radius:8px; background:#F8FAFC;">
                <div style="font-size:0.75rem; font-weight:bold; color:#475569;">HIGH FIDELITY RATE {info_icon("Percentage of datasets achieving high reconstruction fidelity (compression loss ΔR² < 0.05).")}</div>
                <div style="font-size:1.4rem; font-weight:800; color:#15803D;">{summary['high_fidelity_pct']}%</div>
            </div>
        </div>
    </div>
    """


def execute_pipeline_job(
    uploaded_files,
    min_len,
    max_len,
    threshold,
    min_rec,
    lang='en',
    strategy="Hybrid Multi-Objective (Variance x Count)",
    progress=gr.Progress(track_tqdm=False)
):
    """
    Executes the greedy Set Cover pipeline and returns outputs.
    Note: Can be called with 5 arguments (backward-compatible test calls) or 6+ arguments from GUI.
    """
    file_list = []
    if uploaded_files:
        if isinstance(uploaded_files, list):
            file_list = [f.name if hasattr(f, 'name') else str(f) for f in uploaded_files]
        else:
            file_list = [uploaded_files.name if hasattr(uploaded_files, 'name') else str(uploaded_files)]

    if not file_list:
        msg = "[שגיאה] לא הועלו קבצים. אנא העלה קבצי Excel (.xlsx) או ארכיון .zip תחילה." if lang == 'he' else "[ERROR] No dataset provided. Please upload Excel spreadsheet(s) (.xlsx) or a .zip archive containing the SA datasets."
        error_html = f"""
        <div style="background:#FEF2F2; border:2px solid #DC2626; border-radius:8px; padding:16px; margin:15px 0; color:#991B1B; font-weight:bold;">
            {msg}
        </div>
        """
        return error_html, "", "", None, None, None, None, None, None, None, None

    def progress_callback(pct, msg):
        if progress:
            progress(pct, desc=msg)

    timestamp = time.strftime('%Y%m%d_%H%M%S')
    output_run_dir = os.path.join('Docs', f"Run_{timestamp}")

    try:
        summary, df_es, df_mat, df_dataset, df_anc = run_user_pipeline(
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

        summary_html = format_summary_cards_html(summary, min_len, max_len, threshold, lang=lang)
        bar_svg_html = generate_es_bar_chart_svg(df_es, lang=lang)
        heatmap_html = generate_conflict_free_heatmap_html(df_mat, df_es, lang=lang)

        return (
            summary_html,
            bar_svg_html,
            heatmap_html,
            df_es,
            df_dataset,
            df_mat,
            summary['opt_es_result_path'],
            summary['html_report_path'],
            summary['zip_bundle_path'],
            summary['anchor_xlsx_path'],
            summary['es_combinations_csv_path']
        )

    except Exception as e:
        import traceback
        err_msg = traceback.format_exc()
        error_html = f"""
        <div style="background:#FEF2F2; border:2px solid #DC2626; border-radius:8px; padding:16px; margin:15px 0;">
            <p style="margin:0 0 8px 0; color:#991B1B !important; font-weight:bold; font-size:1rem;">
                [ERROR] An error occurred during execution:
            </p>
            <pre style="margin:0; background:#FFFFFF; color:#0F172A !important; padding:10px; border-radius:4px; font-size:0.82rem; overflow-x:auto;">{html.escape(str(e))}\n\n{html.escape(err_msg)}</pre>
        </div>
        """
        return error_html, "", "", None, None, None, None, None, None, None, None


def build_app():
    with gr.Blocks(title="SA Minimal ES Discovery Platform") as demo:
        lang_state = gr.State("en")

        # Top Language Selector Bar (Default: English)
        with gr.Row():
            lang_radio = gr.Radio(
                choices=[("English (Default)", "en"), ("עברית (Hebrew)", "he")],
                value="en",
                label="🌐 Platform Language / שפת הפלטפורמה",
                interactive=True
            )

        # Platform Header Banner
        header_html = gr.HTML(get_header_html("en"))

        # =====================================================================
        # STEP 1 (PAGE 1): OVERVIEW, VISUAL ILLUSTRATION, GUIDE & START BUTTON
        # =====================================================================
        with gr.Column(visible=True) as step1_container:
            step1_guide_html = gr.HTML(get_step1_guide_html("en"))

            with gr.Row():
                template_download = gr.File(
                    value=TEMPLATE_FILE if os.path.exists(TEMPLATE_FILE) else None,
                    label="📄 Sample Input Template (Sample_Input_Template.xlsx)",
                    interactive=False
                )
                dict_download = gr.File(
                    value=DICT_FILE if os.path.exists(DICT_FILE) else None,
                    label="📘 Deliverables Data Dictionary (Deliverables_Data_Dictionary.xlsx)",
                    interactive=False
                )

            start_analysis_btn = gr.Button(
                "Start Analysis (Proceed to Step 2) →",
                variant="primary",
                elem_classes=["btn-primary-action"]
            )

        # =====================================================================
        # STEP 2 (PAGE 2): UPLOAD, PRE-CHECK, GREEDY ANALYSIS & VISUAL OUTPUTS
        # =====================================================================
        with gr.Column(visible=False) as step2_container:
            with gr.Row():
                back_to_guide_btn = gr.Button(
                    "← Back to Step 1 (Platform Guide & Illustration)",
                    elem_classes=["btn-secondary-nav"]
                )

            # Part 2A: Upload Datasets
            step2_upload_title = gr.Markdown("### 1. Upload Surprisal Analysis (SA) Datasets")
            upload_input = gr.File(
                label="Upload Excel File(s) (.xlsx / .xls) or .zip Archive (Format A: Gvalues/Lambavalues or Format B: Tail Tables)",
                file_count="multiple",
                file_types=[".xlsx", ".xls", ".zip"]
            )

            # Part 2B: Pre-Flight Dataset Inspection (Auto-triggered on upload)
            precheck_html = gr.HTML(run_precheck_job(None, 0.05, lang="en"))

            # Part 2C: Execution Parameters with Hover Tooltips (ⓘ) & Wide Ranges
            step2_params_title = gr.Markdown("### 2. Execution Parameters (Supports Compact Panels up to 20,000+ Gene Omics)")
            p1_init, p2_init, p3_init, p4_init = get_param_headers_html("en")

            with gr.Row():
                with gr.Column():
                    p1_hdr = gr.HTML(p1_init)
                    min_len_input = gr.Number(
                        value=3, minimum=2, maximum=50000, precision=0,
                        show_label=False
                    )
                with gr.Column():
                    p2_hdr = gr.HTML(p2_init)
                    max_len_input = gr.Number(
                        value=8, minimum=2, maximum=50000, precision=0,
                        show_label=False
                    )
                with gr.Column():
                    p3_hdr = gr.HTML(p3_init)
                    threshold_input = gr.Number(
                        value=0.05, minimum=0.0001, maximum=1.0, step=0.01,
                        show_label=False
                    )
                with gr.Column():
                    p4_hdr = gr.HTML(p4_init)
                    min_rec_input = gr.Number(
                        value=2, minimum=1, maximum=1000, precision=0,
                        show_label=False
                    )

            # Run Greedy Set Cover Button
            run_btn = gr.Button(
                "▶ Run Greedy Set Cover Analysis & Generate Deliverables",
                variant="primary",
                elem_classes=["btn-primary-action"]
            )

            # Part 2D: Summary Metrics & Visual Interactive Results
            summary_display = gr.HTML()

            with gr.Tabs():
                with gr.TabItem("📊 Visual Charts & Conflict-Free Heatmap"):
                    chart_es_html = gr.HTML()
                    heatmap_mat_html = gr.HTML()
                with gr.TabItem("📋 Interactive Output Tables"):
                    gr.Markdown("#### Optimal_ES_SetCover (Selected Minimal ES Signatures)")
                    tbl_es = gr.Dataframe(interactive=False)
                    gr.Markdown("#### Per_Dataset_Coverage_and_R2 (Unified Coverage & R² Fidelity per Dataset)")
                    tbl_dataset = gr.Dataframe(interactive=False)
                    gr.Markdown("#### ES_Conflict_Free_Matrix (Datasets × Selected ES Signatures)")
                    tbl_mat = gr.Dataframe(interactive=False)

            # Part 2E: Streamlined Deliverable Downloads with ⓘ Explanations
            deliverables_info_box = gr.HTML(get_deliverables_info_html("en"))

            with gr.Row():
                dl_opt_es = gr.File(label="📄 1. Master Workbook: Optimal_ES_Result.xlsx (5 Streamlined Sheets)")
                dl_html_report = gr.File(label="🌐 2. Visual Summary Report: Analysis_Summary_Report.html")
                dl_zip = gr.File(label="📦 3. All-in-One Bundle: Deliverables_Bundle.zip")

            with gr.Accordion("🔬 Optional / Advanced Downloads (Anchor_Strategies_Comparison.xlsx & Filtered ES_Combinations.csv)", open=False):
                with gr.Row():
                    dl_anchors = gr.File(label="📈 Anchor_Strategies_Comparison.xlsx (5-Strategy Overview)")
                    dl_combos = gr.File(label="📊 ES_Combinations.csv (Recurring 2+ Gs Signatures Only)")

        # =====================================================================
        # EVENT HANDLERS: WIZARD NAVIGATION, LANGUAGE SWITCH, PRE-CHECK & RUN
        # =====================================================================
        start_analysis_btn.click(
            fn=lambda: (gr.update(visible=False), gr.update(visible=True)),
            inputs=[],
            outputs=[step1_container, step2_container]
        )

        back_to_guide_btn.click(
            fn=lambda: (gr.update(visible=True), gr.update(visible=False)),
            inputs=[],
            outputs=[step1_container, step2_container]
        )

        # Pre-check automatically runs when files are uploaded or threshold changes
        upload_input.change(
            fn=run_precheck_job,
            inputs=[upload_input, threshold_input, lang_state],
            outputs=[precheck_html]
        )
        threshold_input.change(
            fn=run_precheck_job,
            inputs=[upload_input, threshold_input, lang_state],
            outputs=[precheck_html]
        )

        # Language Switcher Handler
        def switch_language(selected_lang, uploaded_files, threshold):
            p1, p2, p3, p4 = get_param_headers_html(selected_lang)
            pre_html = run_precheck_job(uploaded_files, threshold, lang=selected_lang)
            if selected_lang == 'he':
                return (
                    "he",
                    get_header_html("he"),
                    get_step1_guide_html("he"),
                    gr.update(label="📄 הורדת תבנית קלט לדוגמה (Sample_Input_Template.xlsx)"),
                    gr.update(label="📘 הורדת מילון עמודות לקבצי הפלט (Deliverables_Data_Dictionary.xlsx)"),
                    gr.update(value="התחל אנליזה (מעבר לשלב 2) ←"),
                    gr.update(value="→ חזרה לשלב 1 (מדריך ואיור מתודולוגי)"),
                    "### 1. העלאת דאטה-סטים של Surprisal Analysis (SA)",
                    gr.update(label="העלה קבצי Excel (.xlsx / .xls) או ארכיון .zip (פורמט א': Gvalues/Lambavalues או פורמט ב': טבלאות Tail)"),
                    pre_html,
                    "### 2. פרמטרי הרצה (תומך בדאטה-סטים קטנים ועד 20,000+ גנים)",
                    p1, p2, p3, p4,
                    gr.update(value="▶ הרץ אלגוריתם Set Cover חמדן והפק קבצי פלט"),
                    get_deliverables_info_html("he")
                )
            return (
                "en",
                get_header_html("en"),
                get_step1_guide_html("en"),
                gr.update(label="📄 Sample Input Template (Sample_Input_Template.xlsx)"),
                gr.update(label="📘 Deliverables Data Dictionary (Deliverables_Data_Dictionary.xlsx)"),
                gr.update(value="Start Analysis (Proceed to Step 2) →"),
                gr.update(value="← Back to Step 1 (Platform Guide & Illustration)"),
                "### 1. Upload Surprisal Analysis (SA) Datasets",
                gr.update(label="Upload Excel File(s) (.xlsx / .xls) or .zip Archive (Format A: Gvalues/Lambavalues or Format B: Tail Tables)"),
                pre_html,
                "### 2. Execution Parameters (Supports Compact Panels up to 20,000+ Gene Omics)",
                p1, p2, p3, p4,
                gr.update(value="▶ Run Greedy Set Cover Analysis & Generate Deliverables"),
                get_deliverables_info_html("en")
            )

        lang_radio.change(
            fn=switch_language,
            inputs=[lang_radio, upload_input, threshold_input],
            outputs=[
                lang_state,
                header_html,
                step1_guide_html,
                template_download,
                dict_download,
                start_analysis_btn,
                back_to_guide_btn,
                step2_upload_title,
                upload_input,
                precheck_html,
                step2_params_title,
                p1_hdr, p2_hdr, p3_hdr, p4_hdr,
                run_btn,
                deliverables_info_box
            ]
        )

        # Main Execution Button
        run_btn.click(
            fn=execute_pipeline_job,
            inputs=[
                upload_input,
                min_len_input,
                max_len_input,
                threshold_input,
                min_rec_input,
                lang_state
            ],
            outputs=[
                summary_display,
                chart_es_html,
                heatmap_mat_html,
                tbl_es,
                tbl_dataset,
                tbl_mat,
                dl_opt_es,
                dl_html_report,
                dl_zip,
                dl_anchors,
                dl_combos
            ]
        )

    return demo


def create_web_application():
    return build_app()


if __name__ == '__main__':
    app = build_app()
    app.launch(
        server_name="127.0.0.1",
        server_port=7860,
        inbrowser=True,
        css=CUSTOM_CSS,
        head=HEAD_SCRIPT,
        theme=gr.themes.Default(primary_hue="blue", secondary_hue="slate", neutral_hue="slate")
    )
