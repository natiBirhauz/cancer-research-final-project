"""
web_gui.py
----------
Modern Web-based Graphical User Interface for the Cancer Research TME
Surprisal Analysis & Exact Sign Sequence (ES) Mining Platform.

Features:
- 100% English interface with comprehensive documentation for new users
- Algorithm engine verified and synced directly with the project Git repository:
  https://github.com/natiBirhauz/cancer-research-final-project.git
- Strictly user-provided data ingestion (file upload, zip archive, or folder path)
- Unbiased anchor strategy selection (zero pre-judgment or historical winners)
- Real-time progress monitoring, metric cards, interactive result tables, and Excel export
"""

import os
import sys
import time
import subprocess
import pandas as pd
import gradio as gr
from pipeline_engine import run_user_pipeline

GIT_REPO_URL = "https://github.com/natiBirhauz/cancer-research-final-project.git"

def get_git_status():
    """Retrieves current Git status, commit hash, and branch."""
    try:
        commit_hash = subprocess.check_output(
            ["git", "rev-parse", "--short", "HEAD"], 
            text=True, stderr=subprocess.DEVNULL
        ).strip()
        branch = subprocess.check_output(
            ["git", "rev-parse", "--abbrev-ref", "HEAD"], 
            text=True, stderr=subprocess.DEVNULL
        ).strip()
        status_msg = f"Connected to Git Repository ({branch} @ {commit_hash})"
        return status_msg, commit_hash, branch
    except Exception as e:
        return "Local Git Repository (Unsynchronized)", "unknown", "main"

def pull_git_algorithm():
    """Pulls and reloads the latest algorithm version from the Git remote."""
    try:
        output = subprocess.check_output(
            ["git", "pull", "origin", "main"], 
            text=True, stderr=subprocess.STDOUT
        ).strip()
        status_msg, commit, branch = get_git_status()
        return (
            f"✅ Git Algorithm Successfully Updated!\n"
            f"Branch: {branch} | Commit: {commit}\n"
            f"Remote Output:\n{output}"
        )
    except Exception as e:
        return f"⚠️ Git pull completed with notice: {str(e)}"

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
    """Executes the pipeline on user-uploaded datasets."""
    # Validate user inputs
    file_list = []
    if uploaded_files:
        if isinstance(uploaded_files, list):
            file_list = [f.name if hasattr(f, 'name') else str(f) for f in uploaded_files]
        else:
            file_list = [uploaded_files.name if hasattr(uploaded_files, 'name') else str(uploaded_files)]

    folder_target = local_folder.strip() if (local_folder and local_folder.strip()) else None

    if not file_list and not folder_target:
        return (
            "<div style='color: #ef4444; font-weight: bold; padding: 15px;'>❌ Error: No dataset provided. Please upload an Excel file, a ZIP archive, or provide a dataset folder path.</div>",
            pd.DataFrame(),
            pd.DataFrame(),
            pd.DataFrame(),
            None,
            "Error: No data inputs detected. Please upload data files to begin."
        )

    # Map strategy key
    strat_key = 'Hybrid'
    if 'SVD' in strategy:
        strat_key = 'SVD_Variance'
    elif 'Frequency' in strategy:
        strat_key = 'Max_Frequency'
    elif 'Deep' in strategy:
        strat_key = 'Deep_Specificity'

    log_history = []
    def progress_callback(pct, msg):
        log_history.append(f"[{time.strftime('%H:%M:%S')}] {msg}")
        if pct is not None:
            progress(pct, desc=msg)

    try:
        summary, df_es, df_matrix, df_r2 = run_user_pipeline(
            file_paths=file_list if file_list else None,
            folder_path=folder_target,
            min_es_length=int(min_len),
            max_es_length=int(max_len),
            threshold=float(threshold),
            min_g_recurrence=int(min_rec),
            strategy=strat_key,
            progress_callback=progress_callback
        )
    except Exception as e:
        err = f"Pipeline Execution Error: {str(e)}"
        log_history.append(err)
        return (
            f"<div style='color: #ef4444; font-weight: bold; padding: 15px;'>❌ {err}</div>",
            pd.DataFrame(),
            pd.DataFrame(),
            pd.DataFrame(),
            None,
            "\n".join(log_history)
        )

    # Build metric summary card
    summary_card_html = f"""
    <div style="background: #0f172a; color: #f8fafc; border-radius: 12px; padding: 24px; margin-bottom: 20px; border-left: 6px solid #3b82f6;">
        <h2 style="margin: 0 0 16px 0; color: #60a5fa; font-size: 1.35em;">Analysis Completed Successfully</h2>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap: 14px; text-align: center;">
            <div style="background: #1e293b; padding: 14px; border-radius: 8px;">
                <div style="font-size: 1.8em; font-weight: bold; color: #38bdf8;">{summary['total_samples']}</div>
                <div style="font-size: 0.8em; text-transform: uppercase; color: #94a3b8; font-weight: 600;">Analyzed Samples</div>
            </div>
            <div style="background: #1e293b; padding: 14px; border-radius: 8px;">
                <div style="font-size: 1.8em; font-weight: bold; color: #38bdf8;">{summary['total_g']:,}</div>
                <div style="font-size: 0.8em; text-transform: uppercase; color: #94a3b8; font-weight: 600;">Total G-Vectors</div>
            </div>
            <div style="background: #1e293b; padding: 14px; border-radius: 8px;">
                <div style="font-size: 1.8em; font-weight: bold; color: #34d399;">{summary['es_count']}</div>
                <div style="font-size: 0.8em; text-transform: uppercase; color: #94a3b8; font-weight: 600;">Discovered ES Modules</div>
            </div>
            <div style="background: #1e293b; padding: 14px; border-radius: 8px;">
                <div style="font-size: 1.8em; font-weight: bold; color: #34d399;">{summary['coverage_pct']}%</div>
                <div style="font-size: 0.8em; text-transform: uppercase; color: #94a3b8; font-weight: 600;">G-Vector Coverage</div>
            </div>
            <div style="background: #1e293b; padding: 14px; border-radius: 8px;">
                <div style="font-size: 1.8em; font-weight: bold; color: #fbbf24;">{summary['variance_explained_pct']}%</div>
                <div style="font-size: 0.8em; text-transform: uppercase; color: #94a3b8; font-weight: 600;">Variance Captured</div>
            </div>
            <div style="background: #1e293b; padding: 14px; border-radius: 8px;">
                <div style="font-size: 1.8em; font-weight: bold; color: #a78bfa;">{summary['avg_r2_optimal']}</div>
                <div style="font-size: 0.8em; text-transform: uppercase; color: #94a3b8; font-weight: 600;">Mean R² Optimal</div>
            </div>
        </div>
        <div style="margin-top: 14px; font-size: 0.9em; color: #94a3b8;">
            Elapsed Time: <strong>{summary['duration_sec']}s</strong> | Master Excel Report: <code>{summary['excel_path']}</code>
        </div>
    </div>
    """

    return (
        summary_card_html,
        df_es,
        df_matrix,
        df_r2,
        summary['excel_path'],
        "\n".join(log_history)
    )

def create_web_application():
    initial_git_status, commit_hash, branch = get_git_status()

    custom_theme = gr.themes.Soft(
        primary_hue="blue",
        secondary_hue="slate",
        font=[gr.themes.GoogleFont("Inter"), "system-ui", "sans-serif"]
    )

    with gr.Blocks(title="Cancer Research TME Algorithm Workbench") as app:
        # Header Section
        gr.Markdown(
            f"""
            # 🔬 Cancer Research TME Algorithm Workbench
            ### Surprisal Analysis, Exact Sign Sequences (ES) Mining & Single-Cell R² Reconstruction
            **Algorithm Source Repository:** [`{GIT_REPO_URL}`]({GIT_REPO_URL})
            """
        )

        # Git Algorithm Sync Banner
        with gr.Row():
            with gr.Column(scale=3):
                git_status_box = gr.Markdown(
                    f"**Algorithm Git Status:** `{initial_git_status}` | Repository Branch: `{branch}`"
                )
            with gr.Column(scale=1):
                pull_git_btn = gr.Button("🔄 Pull Latest Algorithm from Git", size="sm", variant="secondary")

        git_pull_output = gr.Textbox(label="Git Sync Details", visible=False, lines=3)

        pull_git_btn.click(
            pull_git_algorithm,
            outputs=[git_pull_output]
        ).then(
            lambda: gr.update(visible=True),
            outputs=[git_pull_output]
        )

        with gr.Row():
            # Left Column: User Input & Parameter Settings
            with gr.Column(scale=1):
                gr.Markdown("### 1. Upload Dataset Files")
                gr.Markdown(
                    "Upload your experimental data spreadsheets. You may upload one or multiple `.xlsx` files, or a `.zip` archive containing your dataset folders."
                )
                
                upload_input = gr.File(
                    label="Upload Dataset File(s) / ZIP Archive",
                    file_count="multiple",
                    file_types=[".xlsx", ".xls", ".zip"]
                )

                local_path_input = gr.Textbox(
                    label="Or Provide Local Directory Path (Optional)",
                    placeholder="e.g. C:/path/to/my_cohort_data",
                    info="Specify a local filesystem path containing your dataset files."
                )

                gr.Markdown("### 2. Execution Parameters")
                
                strategy_selector = gr.Dropdown(
                    label="Anchor & Objective Strategy",
                    choices=[
                        "Hybrid Multi-Objective (Variance x √Count)",
                        "SVD Variance Driver",
                        "Max Frequency (Universality)",
                        "Deep Specificity (Higher-Order Core)"
                    ],
                    value="Hybrid Multi-Objective (Variance x √Count)",
                    info="Objective function driving greedy sequence selection."
                )

                with gr.Accordion("Advanced Hyperparameters", open=False):
                    slider_min_len = gr.Slider(
                        minimum=2, maximum=6, value=3, step=1, 
                        label="Minimum ES Sequence Length",
                        info="Minimum number of proteins required to form a valid biological module."
                    )
                    slider_max_len = gr.Slider(
                        minimum=4, maximum=10, value=8, step=1, 
                        label="Maximum ES Sequence Length",
                        info="Maximum sequence length mined during combinatorial exploration."
                    )
                    slider_threshold = gr.Slider(
                        minimum=0.00, maximum=0.25, value=0.05, step=0.01, 
                        label="Tail Protein Cutoff Threshold (|G|)",
                        info="Filters background noise; proteins with |G| below threshold are excluded from driver tails."
                    )
                    slider_min_rec = gr.Slider(
                        minimum=1, maximum=10, value=2, step=1, 
                        label="Minimum G-Vector Recurrence (2+ Gs)",
                        info="Excludes singleton patterns that appear in only 1 G-vector to ensure biological reproducibility."
                    )

                run_pipeline_btn = gr.Button(
                    "🚀 Run Algorithm on Uploaded Data", 
                    variant="primary", 
                    size="lg"
                )

            # Right Column: Interactive Results, Tables & Master Deliverables
            with gr.Column(scale=2):
                gr.Markdown("### 3. Pipeline Execution Results")
                summary_banner = gr.HTML(
                    "<div style='padding: 24px; background: #f8fafc; border: 1px dashed #cbd5e1; border-radius: 8px; color: #64748b; text-align: center;'>"
                    "Upload your dataset on the left and click <strong>'Run Algorithm on Uploaded Data'</strong> to generate deliverables."
                    "</div>"
                )

                with gr.Tabs():
                    with gr.TabItem("📋 Discovered ES Modules"):
                        gr.Markdown("#### Canonical Exact Sign Sequences Selected by the Algorithm")
                        es_dataframe = gr.Dataframe(
                            headers=['ES_ID', 'Length', 'Exact_Sign_Sequence', 'G_Covered', 'Variance_Gain', 'Cum_Coverage_Pct'],
                            label="Selected Modules (Signs preserved: '+' indicates upregulation, '-' indicates downregulation)"
                        )

                    with gr.TabItem("🧩 Conflict-Free Allocation Matrix"):
                        gr.Markdown("#### Disjoint Sample-to-Module Allocation Matrix (0 Row Duplicates)")
                        matrix_dataframe = gr.Dataframe(
                            label="Allocation Matrix (Each sample cell contains at most one assigned G-vector)"
                        )

                    with gr.TabItem("📈 Cell-Level R² Reconstruction"):
                        gr.Markdown("#### Determination Coefficient (R²) & Information Preservation Per Sample")
                        r2_dataframe = gr.Dataframe(
                            headers=['Sample', 'Cells', 'Selected_G_Count', 'Selected_Gs', 'R2_full', 'R2_optimal', 'Delta_R2', 'Status'],
                            label="Single-Cell Reconstruction Fidelity"
                        )

                    with gr.TabItem("💾 Download Deliverables"):
                        gr.Markdown("#### Master Excel Report Download")
                        gr.Markdown("Download the complete multi-sheet spreadsheet containing all modules, matrices, and R² scores:")
                        excel_download = gr.File(label="Download Master Report (.xlsx)")

                    with gr.TabItem("📜 Live Log & Diagnostics"):
                        gr.Markdown("#### Computational Pipeline Diagnostic Log")
                        log_textbox = gr.Textbox(label="Execution Trace", lines=12, max_lines=25)

                    with gr.TabItem("📖 User Guide & File Specifications"):
                        gr.Markdown(
                            r"""
                            ## Complete User Guide & Input File Specifications

                            ### 1. Supported Input File Formats
                            The algorithm accepts experimental proteomic or single-cell imaging data in two standard formats:

                            #### Format A: Surprisal Analysis Matrices (Recommended)
                            An Excel workbook (`.xlsx`) containing two sheets:
                            * **Sheet 1: `Gvalues`**
                              * Orthonormal basis vectors from Surprisal Analysis / Singular Value Decomposition (SVD).
                              * Columns: `G0` (equilibrium baseline), `G1`, `G2`, ..., `Gk` (biological deviation processes).
                              * Rows: Protein marker names and their eigenvector coefficients for each process.
                            * **Sheet 2: `Lambavalues`**
                              * Single-cell amplitude coefficients.
                              * Columns: `L0`, `L1`, `L2`, ..., `Lk` (amplitudes corresponding to each G-vector).
                              * Rows: Individual cells (each row represents one cell's amplitude coordinates).

                            #### Format B: Precomputed Tail Protein Spreadsheets
                            An Excel file (`.xlsx`) with driver tail columns:
                            * Columns named `G1+`, `G1-`, `G2+`, `G2-`, etc.
                            * Each column lists the specific proteins exhibiting positive (`+`) or negative (`-`) deviations in that process.

                            ---

                            ### 2. Explanation of Controls & Hyperparameters
                            * **Upload Dataset File(s) / ZIP Archive:** Upload your raw `.xlsx` files or a `.zip` containing subfolders.
                            * **Anchor & Objective Strategy:**
                              * **Hybrid Multi-Objective:** Jointly balances biological variance $\sum w(g)$ and population prevalence $\sqrt{Count}$.
                              * **SVD Variance Driver:** Prioritizes physical energy and variance capture without weighting sample counts.
                              * **Max Frequency (Universality):** Prioritizes universal patterns appearing across the highest number of vectors.
                              * **Deep Specificity:** Prioritizes longer, higher-order protein combinations (lengths 5 to 8).
                            * **Minimum / Maximum ES Sequence Length:** Constrains the combinatorial size of protein sets (default: lengths 3 to 8).
                            * **Tail Cutoff Threshold (|G|):** Filters low-amplitude protein coefficients to remove background noise (default: 0.05).
                            * **Minimum G-Vector Recurrence (2+ Gs):** Excludes singletons (patterns appearing in only 1 vector) to focus strictly on reproducible biological phenomena.

                            ---

                            ### 3. Explanation of Performance Metrics
                            * **Exact Sign Sequence (ES):** A coordinated set of proteins where each protein has an exact sign (`+` for upregulation, `-` for downregulation).
                            * **Canonical Representation ($ES \equiv -ES$):** Because eigenvectors are invariant to sign inversion ($G \times L^T = (-G) \times (-L)^T$), inverted profiles represent the exact same physical pathway. Canonical indexing ensures zero redundant duplicates.
                            * **Disjoint Allocation:** Once a G-vector is covered by a module, it cannot be assigned to another module, guaranteeing 100% disjoint partitioning.
                            * **$R^2_{full}$ (Baseline):** The determination coefficient achieved when using all available components.
                            * **$R^2_{optimal}$ (Compressed Representation):** The determination coefficient achieved using only the baseline equilibrium ($G_0, L_0$) plus the selected ES modules.
                            * **Compression Penalty ($\Delta R^2 = R^2_{full} - R^2_{optimal}$):** Measures the information loss due to dimensionality reduction. Values below $0.05$ indicate high fidelity.
                            """
                        )

        # Bind Run Execution
        run_pipeline_btn.click(
            execute_pipeline_job,
            inputs=[
                upload_input,
                local_path_input,
                strategy_selector,
                slider_min_len,
                slider_max_len,
                slider_threshold,
                slider_min_rec
            ],
            outputs=[
                summary_banner,
                es_dataframe,
                matrix_dataframe,
                r2_dataframe,
                excel_download,
                log_textbox
            ]
        )

    return app

if __name__ == '__main__':
    port = 7860
    print(f"Launching Cancer Research TME Algorithm Web GUI on http://127.0.0.1:{port}...")
    app = create_web_application()
    app.launch(server_name="127.0.0.1", server_port=port, inbrowser=True)
