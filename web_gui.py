"""
web_gui.py
----------
Modern Web-based User Interface for the Cancer Research TME
Surprisal Analysis & Exact Sign Sequence (ES) Mining Platform.

Fully English, zero Hebrew.
Provides:
1. Step-by-step instructions on required file and folder formats.
2. Clear explanations of all inputs, controls, and scientific metrics.
3. User-driven dataset loading (single/multiple Excel files, ZIP archive, or directory path).
4. Automated execution generating the 4 master deliverables:
   - Optimal_ES_Result.xlsx
   - Optimal_ES_Result_Signs_R2.xlsx
   - ES_Combinations.csv
   - Anchor_Strategies_Comparison.xlsx
   - Plus Deliverables_Bundle.zip
5. Interactive preview tables and instant download buttons.
"""

import os
import sys
import time
import pandas as pd
import gradio as gr
from pipeline_engine import run_user_pipeline

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
    """Executes the pipeline on user-provided datasets and exports the 4 deliverables."""
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
        <div style="background-color: #FEF2F2; border: 1px solid #F87171; border-radius: 8px; padding: 16px; color: #991B1B; font-weight: 500;">
            ❌ <strong>Input Error:</strong> No dataset provided. Please upload Excel (.xlsx) file(s), a .zip archive, or specify a valid dataset folder path.
        </div>
        """
        return (
            error_html,
            pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(),
            None, None, None, None, None,
            "Error: No data inputs detected. Please upload dataset files to proceed."
        )

    log_history = []
    def progress_callback(pct, msg):
        log_history.append(f"[{time.strftime('%H:%M:%S')}] {msg}")
        if progress:
            progress(pct, desc=msg)

    timestamp = time.strftime('%Y%m%d_%H%M%S')
    output_run_dir = os.path.join('Docs', f"Run_{timestamp}")

    try:
        summary, df_es, df_mat, df_r2, df_anc = run_user_pipeline(
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

        # Build Summary Card HTML
        summary_html = f"""
        <div style="background-color: #F0FDF4; border: 1px solid #86EFAC; border-radius: 8px; padding: 20px; margin-bottom: 15px;">
            <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #BBF7D0; padding-bottom: 10px; margin-bottom: 15px;">
                <h3 style="margin: 0; color: #166534; font-size: 1.25rem;">✅ Analysis Completed Successfully in {summary['duration_sec']}s</h3>
                <span style="background-color: #DCFCE7; color: #15803D; padding: 4px 12px; border-radius: 9999px; font-weight: 600; font-size: 0.85rem;">
                    Strategy: {strategy}
                </span>
            </div>
            <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 15px;">
                <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #DCFCE7;">
                    <div style="color: #64748B; font-size: 0.8rem; text-transform: uppercase; font-weight: 600;">Cohort Samples</div>
                    <div style="font-size: 1.5rem; font-weight: 700; color: #0F172A;">{summary['total_samples']}</div>
                </div>
                <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #DCFCE7;">
                    <div style="color: #64748B; font-size: 0.8rem; text-transform: uppercase; font-weight: 600;">G-Vectors Ingested</div>
                    <div style="font-size: 1.5rem; font-weight: 700; color: #0F172A;">{summary['total_g']:,}</div>
                </div>
                <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #DCFCE7;">
                    <div style="color: #64748B; font-size: 0.8rem; text-transform: uppercase; font-weight: 600;">Optimal ES Modules</div>
                    <div style="font-size: 1.5rem; font-weight: 700; color: #2563EB;">{summary['es_count']}</div>
                </div>
                <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #DCFCE7;">
                    <div style="color: #64748B; font-size: 0.8rem; text-transform: uppercase; font-weight: 600;">Universe Coverage</div>
                    <div style="font-size: 1.5rem; font-weight: 700; color: #16A34A;">{summary['coverage_pct']}%</div>
                </div>
                <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #DCFCE7;">
                    <div style="color: #64748B; font-size: 0.8rem; text-transform: uppercase; font-weight: 600;">Kinetic Variance</div>
                    <div style="font-size: 1.5rem; font-weight: 700; color: #9333EA;">{summary['variance_explained_pct']}%</div>
                </div>
                <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #DCFCE7;">
                    <div style="color: #64748B; font-size: 0.8rem; text-transform: uppercase; font-weight: 600;">Mean Optimal R²</div>
                    <div style="font-size: 1.5rem; font-weight: 700; color: #D97706;">{summary['avg_r2_optimal']:.4f}</div>
                </div>
                <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #DCFCE7;">
                    <div style="color: #64748B; font-size: 0.8rem; text-transform: uppercase; font-weight: 600;">Mean Compression Loss (ΔR²)</div>
                    <div style="font-size: 1.5rem; font-weight: 700; color: #475569;">{summary['avg_delta_r2']:.4f}</div>
                </div>
                <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #DCFCE7;">
                    <div style="color: #64748B; font-size: 0.8rem; text-transform: uppercase; font-weight: 600;">High Fidelity Rate</div>
                    <div style="font-size: 1.5rem; font-weight: 700; color: #059669;">{summary['high_fidelity_pct']}%</div>
                </div>
            </div>
            <p style="margin: 15px 0 0 0; color: #15803D; font-size: 0.9rem;">
                All 4 master deliverable files and the consolidated ZIP archive are ready for download below.
            </p>
        </div>
        """

        return (
            summary_html,
            df_es,
            df_mat,
            df_r2,
            df_anc,
            summary['opt_es_result_path'],
            summary['opt_signs_r2_path'],
            summary['es_combinations_csv_path'],
            summary['anchor_xlsx_path'],
            summary['zip_bundle_path'],
            "\n".join(log_history)
        )

    except Exception as e:
        import traceback
        err_detail = traceback.format_exc()
        error_html = f"""
        <div style="background-color: #FEF2F2; border: 1px solid #F87171; border-radius: 8px; padding: 16px; color: #991B1B;">
            <h4 style="margin: 0 0 8px 0;">❌ Pipeline Execution Error</h4>
            <p style="margin: 0; font-family: monospace; font-size: 0.85rem;">{str(e)}</p>
        </div>
        """
        return (
            error_html,
            pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(),
            None, None, None, None, None,
            f"Error during execution:\n{err_detail}"
        )

def build_app():
    custom_css = """
    .gradio-container { max-width: 1300px !important; margin: auto !important; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Arial, sans-serif; }
    .header-box { background: linear-gradient(135deg, #0F172A 0%, #1E293B 100%); color: white; padding: 28px; border-radius: 12px; margin-bottom: 24px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
    .header-title { font-size: 1.85rem; font-weight: 700; margin: 0 0 8px 0; color: #F8FAFC; }
    .header-subtitle { font-size: 0.95rem; color: #94A3B8; margin: 0; line-height: 1.5; }
    .info-card { background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 10px; padding: 20px; margin-bottom: 20px; }
    .section-title { font-size: 1.15rem; font-weight: 600; color: #1E293B; margin-top: 0; margin-bottom: 12px; border-bottom: 2px solid #E2E8F0; padding-bottom: 6px; }
    .btn-run { background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%) !important; color: white !important; font-size: 1.05rem !important; font-weight: 600 !important; border-radius: 8px !important; padding: 12px 24px !important; box-shadow: 0 2px 6px rgba(37,99,235,0.3) !important; }
    .btn-run:hover { background: linear-gradient(135deg, #1D4ED8 0%, #1E40AF 100%) !important; }
    .code-block { background: #0F172A; color: #38BDF8; padding: 12px; border-radius: 6px; font-family: monospace; font-size: 0.85rem; }
    """

    with gr.Blocks(title="Cancer Research TME Pipeline", css=custom_css, theme=gr.themes.Default(primary_hue="blue")) as demo:
        
        # Header Box
        gr.HTML("""
        <div class="header-box">
            <h1 class="header-title">Cancer Research: Tumor Microenvironment (TME) Analysis Platform</h1>
            <p class="header-subtitle">
                Surprisal Analysis Matrix Ingestion &bull; Combinatorial Exact Sign Sequence (ES) Mining &bull; Greedy Multi-Anchor Set Cover &bull; Single-Cell Reconstruction (R²)
            </p>
        </div>
        """)

        # User Guide Accordion
        with gr.Accordion("📘 Complete User Manual: Input File Formats, Scientific Methodology & Metric Definitions", open=True):
            gr.HTML("""
            <div class="info-card">
                <h3 class="section-title">1. Expected Dataset Formats & Directory Layout</h3>
                <p style="margin-top: 0; color: #334155; line-height: 1.6;">
                    The platform automatically detects and ingests two standard scientific data formats. You can upload individual files, multiple files simultaneously, a compressed <code>.zip</code> archive, or provide a local folder path:
                </p>
                <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 16px;">
                    <div style="background: white; border: 1px solid #CBD5E1; border-radius: 8px; padding: 16px;">
                        <h4 style="margin: 0 0 8px 0; color: #0F172A;">Format A: Surprisal Analysis Spreadsheets (Matrix Format)</h4>
                        <p style="font-size: 0.88rem; color: #475569; margin: 0 0 10px 0;">
                            An Excel workbook (<code>.xlsx</code>) containing two specific sheets:
                        </p>
                        <ul style="font-size: 0.85rem; color: #334155; margin: 0 0 10px 0; padding-left: 20px;">
                            <li><strong><code>Gvalues</code> Sheet:</strong> Protein loading matrix where rows represent proteins and columns represent biological components (<code>G0, G1, ..., Gk</code>). One column must contain protein names (e.g. <code>MARKER</code> or first text column).</li>
                            <li><strong><code>Lambavalues</code> Sheet:</strong> Single-cell amplitude matrix where rows represent cells and columns represent amplitude weights (<code>L0, L1, ..., Lk</code>).</li>
                        </ul>
                        <div style="font-size: 0.8rem; background: #F1F5F9; padding: 8px; border-radius: 4px; color: #0F172A;">
                            <em>Example: Clinical PDE data files with <code>Gvalues</code> and <code>Lambavalues</code>.</em>
                        </div>
                    </div>
                    <div style="background: white; border: 1px solid #CBD5E1; border-radius: 8px; padding: 16px;">
                        <h4 style="margin: 0 0 8px 0; color: #0F172A;">Format B: Tail Proteins Spreadsheets (Extracted Tails)</h4>
                        <p style="font-size: 0.88rem; color: #475569; margin: 0 0 10px 0;">
                            An Excel workbook (<code>.xlsx</code>) containing pre-extracted protein tails:
                        </p>
                        <ul style="font-size: 0.85rem; color: #334155; margin: 0 0 10px 0; padding-left: 20px;">
                            <li>Columns formatted as <code>G1+, G1-, G2+, G2-, ..., Gk+, Gk-</code>.</li>
                            <li>Under each column, list the protein names classified as significantly over-expressed (<code>+</code>) or under-expressed (<code>-</code>) in that component.</li>
                        </ul>
                        <div style="font-size: 0.8rem; background: #F1F5F9; padding: 8px; border-radius: 4px; color: #0F172A;">
                            <em>Example: <code>ROI_XX_tail_proteins.xlsx</code> tables from IMC cohorts.</em>
                        </div>
                    </div>
                </div>

                <h3 class="section-title">2. Scientific Metrics & Optimization Objectives</h3>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 14px; font-size: 0.88rem; color: #334155;">
                    <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #E2E8F0;">
                        <strong>Exact Sign Sequence (ES):</strong> A coordinated multi-protein signature with strict directional signs (e.g., <code>CKIT(+) | EGFR(+) | IGFR(+)</code>). Captures co-functional protein complexes without sign contradictions.
                    </div>
                    <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #E2E8F0;">
                        <strong>Universe G Coverage:</strong> The cumulative percentage of all microenvironment states (G-vectors) matched by the selected set of ES modules with zero sign errors.
                    </div>
                    <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #E2E8F0;">
                        <strong>Kinetic Variance Explained:</strong> The proportion of the total physical deviation energy &Sigma;(L<sub>k</sub>)<sup>2</sup> accounted for by the components selected in the optimal cover.
                    </div>
                    <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #E2E8F0;">
                        <strong>Reconstruction Fidelity (R²):</strong> Measure of single-cell expression matrix reconstruction quality. <code>R²_full</code> is the baseline across all components; <code>R²_optimal</code> is the reconstructed quality from the compact ES set.
                    </div>
                    <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #E2E8F0;">
                        <strong>Compression Penalty (&Delta;R²):</strong> Information loss incurred by data compression: &Delta;R² = R²_full - R²_optimal. Loss &lt; 0.05 indicates high-fidelity preservation.
                    </div>
                    <div style="background: white; padding: 12px; border-radius: 6px; border: 1px solid #E2E8F0;">
                        <strong>Conflict-Free Allocation Matrix:</strong> A strict 1-to-1 disjoint assignment ensuring no single biological G-vector is counted multiple times across ES modules.
                    </div>
                </div>
            </div>
            """)

        # Data Ingestion & Controls Section
        with gr.Row():
            with gr.Column(scale=3):
                gr.Markdown("### 1. Ingest Dataset Files")
                upload_input = gr.File(
                    label="Upload Dataset Files (.xlsx spreadsheets or .zip archive)",
                    file_count="multiple",
                    file_types=[".xlsx", ".xls", ".zip"]
                )
                folder_input = gr.Textbox(
                    label="Or Specify Dataset Directory Path (Local folder on disk)",
                    placeholder="e.g., C:/Projects/Cancer_TME_Data/PDE_Cohort",
                    lines=1
                )

            with gr.Column(scale=2):
                gr.Markdown("### 2. Execution Parameters")
                strategy_select = gr.Dropdown(
                    label="Primary Optimization Anchor Strategy",
                    choices=[
                        "Hybrid Multi-Objective (Variance x Count)",
                        "SVD Variance Driver",
                        "Max Frequency / Universality",
                        "Deep Specificity",
                        "Biological Hallmark (Seed Co-expression)"
                    ],
                    value="Hybrid Multi-Objective (Variance x Count)",
                    info="Mathematical objective used by the Greedy Set Cover solver."
                )

                with gr.Row():
                    min_len_input = gr.Slider(minimum=3, maximum=6, value=3, step=1, label="Min Sequence Length")
                    max_len_input = gr.Slider(minimum=4, maximum=10, value=8, step=1, label="Max Sequence Length")

                with gr.Row():
                    threshold_input = gr.Slider(minimum=0.01, maximum=0.20, value=0.05, step=0.01, label="Cutoff Loading (|G|)")
                    min_rec_input = gr.Slider(minimum=1, maximum=5, value=2, step=1, label="Min G Recurrence")

        # Action Button
        run_btn = gr.Button("🚀 Run Algorithm & Generate 4 Master Deliverables", variant="primary", elem_classes=["btn-run"])

        # Results & Summary Area
        summary_display = gr.HTML()

        # Master Deliverables Download Box
        with gr.Accordion("📥 Download Master Deliverable Files", open=True):
            gr.Markdown("#### Direct 1-Click File Downloads:")
            with gr.Row():
                dl_opt_es = gr.File(label="1. Optimal_ES_Result.xlsx")
                dl_opt_r2 = gr.File(label="2. Optimal_ES_Result_Signs_R2.xlsx")
                dl_combos = gr.File(label="3. ES_Combinations.csv")
                dl_anchors = gr.File(label="4. Anchor_Strategies_Comparison.xlsx")
                dl_zip = gr.File(label="📦 Download All (Deliverables_Bundle.zip)")

        # Result Preview Tabs
        with gr.Tabs():
            with gr.TabItem("📊 Selected Optimal ES Modules"):
                gr.Markdown("#### Ranked Exact Sign Sequence (ES) Modules Covering the Cohort")
                es_table = gr.DataFrame(headers=["Rank", "ES_ID", "Length", "New_Gs_Covered", "Direct (+)", "Inverse (-)", "Cumulative_Gs", "Coverage_%", "Exact_Sign_Sequence"])

            with gr.TabItem("🧩 Conflict-Free Allocation Matrix"):
                gr.Markdown("#### 100% Disjoint Sample &times; ES Allocation Matrix (Zero Duplicate Gs)")
                mat_table = gr.DataFrame()

            with gr.TabItem("🎯 Reconstruction Fidelity (R²)"):
                gr.Markdown("#### Single-Cell Expression Matrix Reconstruction & Compression Loss (&Delta;R²)")
                r2_table = gr.DataFrame()

            with gr.TabItem("⚖️ Anchor Strategies Comparison"):
                gr.Markdown("#### Benchmark Overview Across 6 Greedy Anchor Strategies")
                anc_table = gr.DataFrame()

            with gr.TabItem("📋 Execution Console Log"):
                log_box = gr.Textbox(label="Step-by-Step Processing Logs", lines=10, max_lines=25)

        # Connect callback
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
                es_table,
                mat_table,
                r2_table,
                anc_table,
                dl_opt_es,
                dl_opt_r2,
                dl_combos,
                dl_anchors,
                dl_zip,
                log_box
            ]
        )

    return demo

def main():
    app = build_app()
    app.launch(server_name="127.0.0.1", server_port=7860, share=False)

create_web_application = build_app

if __name__ == '__main__':
    main()
