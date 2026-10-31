"""
web_gui.py
----------
Simplified, Clean High-Contrast Web User Interface for the
Cancer Research TME Surprisal Analysis & Exact Sign Sequence (ES) Platform.

Key Features:
- 100% English, high-contrast black text on white background (zero color clashes).
- Forces Light Mode (__theme=light) to prevent OS dark mode conflicts.
- Direct download of Sample_Input_Template.xlsx format guide.
- User-supplied dataset ingestion (Excel files, ZIP archive, or directory path).
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
    local_folder,
    strategy,
    min_len,
    max_len,
    threshold,
    min_rec,
    progress=gr.Progress(track_tqdm=False)
):
    """Executes the pipeline and produces the 4 master deliverables."""
    # 1. Parse uploaded files / directory
    file_list = []
    if uploaded_files:
        if isinstance(uploaded_files, list):
            file_list = [f.name if hasattr(f, 'name') else str(f) for f in uploaded_files]
        else:
            file_list = [uploaded_files.name if hasattr(uploaded_files, 'name') else str(uploaded_files)]

    folder_target = local_folder.strip() if (local_folder and local_folder.strip()) else None

    if not file_list and not folder_target:
        error_html = """
        <div style="background: #FFFFFF; border: 2px solid #000000; border-radius: 6px; padding: 16px; margin: 15px 0;">
            <p style="margin: 0; color: #000000 !important; font-weight: bold; font-size: 1rem;">
                [ERROR] No dataset provided. Please upload Excel spreadsheet(s), a .zip archive, or enter a valid local directory path.
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
            file_paths=file_list if file_list else None,
            folder_path=folder_target,
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

        # Instruction Guide & Sample Template Section
        with gr.Accordion("📘 Instructions & Format Guide (Click to View)", open=True):
            gr.HTML("""
            <div style="background: #FFFFFF; border: 1px solid #000000; border-radius: 6px; padding: 16px; margin-bottom: 16px;">
                <h3 style="margin-top: 0; color: #000000 !important; font-size: 1.15rem; font-weight: bold; border-bottom: 1px solid #000000; padding-bottom: 6px;">
                    How to Format Your Data Without Errors
                </h3>
                <p style="margin: 10px 0; color: #000000 !important; font-size: 0.95rem; line-height: 1.5;">
                    The algorithm processes single or multiple patient/sample datasets in two standard Excel spreadsheet formats:
                </p>
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
                <p style="margin: 8px 0 0 0; color: #000000 !important; font-size: 0.9rem;">
                    <strong>Tip:</strong> Download the pre-formatted Excel template below to inspect the exact structure and column names.
                </p>
            </div>
            """)

            with gr.Row():
                template_download = gr.File(
                    value=TEMPLATE_FILE if os.path.exists(TEMPLATE_FILE) else None,
                    label="📄 Download Sample Format Template (Sample_Input_Template.xlsx)",
                    interactive=False
                )

        # Dataset Ingestion Section
        gr.Markdown("### 1. Ingest Dataset Files")
        with gr.Row():
            upload_input = gr.File(
                label="Upload Excel File(s) (.xlsx / .xls) or .zip Archive",
                file_count="multiple",
                file_types=[".xlsx", ".xls", ".zip"]
            )
            folder_input = gr.Textbox(
                label="Or Enter Local Directory Path on Disk",
                placeholder="e.g., C:/Datasets/Cancer_Cohort",
                lines=2
            )

        # Execution Parameters Section
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

        # Deliverables Download Section
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
                folder_input,
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
