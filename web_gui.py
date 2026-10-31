"""
web_gui.py
----------
Modern Web-based Graphical User Interface for the Cancer Research TME
Surprisal Analysis & Exact Sign Sequence (ES) Mining Platform.

Built with Gradio. Allows users to:
1. Select dataset: Evaluation Cohort (6 samples), Original 120-ROI Dataset, or Custom Folder.
2. Fine-tune parameters (min length, threshold, anchor strategy, min recurrence).
3. Execute the full end-to-end pipeline with live status updates.
4. Inspect discovered ES modules, allocation matrices, and R2 reconstruction metrics.
5. Download exported Excel and HTML reports directly.
"""

import os
import sys
import time
import pandas as pd
import gradio as gr
from pipeline_engine import run_pipeline

def execute_gui_job(
    dataset_choice,
    custom_dir,
    strategy,
    min_len,
    max_len,
    threshold,
    min_rec,
    progress=gr.Progress(track_tqdm=False)
):
    source_key = 'evaluation'
    if 'Original' in dataset_choice:
        source_key = 'original'
    elif 'Custom' in dataset_choice:
        source_key = 'custom'
        
    strat_key = 'Hybrid'
    if 'SVD' in strategy:
        strat_key = 'SVD_Variance'
    elif 'Frequency' in strategy:
        strat_key = 'Max_Frequency'
    elif 'Deep' in strategy:
        strat_key = 'Deep_Specificity'

    log_messages = []
    def p_cb(pct, msg):
        log_messages.append(f"[{time.strftime('%H:%M:%S')}] {msg}")
        if pct is not None:
            progress(pct, desc=msg)

    try:
        summary, df_es, df_matrix, df_r2 = run_pipeline(
            data_source=source_key,
            custom_path=custom_dir if source_key == 'custom' else None,
            min_es_length=int(min_len),
            max_es_length=int(max_len),
            threshold=float(threshold),
            min_g_recurrence=int(min_rec),
            strategy=strat_key,
            progress_callback=p_cb
        )
    except Exception as e:
        err_msg = f"Error during execution: {str(e)}"
        log_messages.append(err_msg)
        return (
            f"❌ {err_msg}",
            pd.DataFrame(),
            pd.DataFrame(),
            pd.DataFrame(),
            None,
            "\n".join(log_messages)
        )

    # Summary HTML card
    summary_html = f"""
    <div style="background: linear-gradient(135deg, #1f2937, #111827); color: white; padding: 24px; border-radius: 12px; margin-bottom: 20px;">
        <h2 style="margin: 0 0 16px 0; color: #60a5fa; font-size: 1.4em;">Pipeline Execution Successful! 🚀</h2>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 16px; text-align: center;">
            <div style="background: rgba(255,255,255,0.08); padding: 12px; border-radius: 8px;">
                <div style="font-size: 1.8em; font-weight: bold; color: #38bdf8;">{summary['total_rois']}</div>
                <div style="font-size: 0.85em; text-transform: uppercase; color: #94a3b8;">Datasets / ROIs</div>
            </div>
            <div style="background: rgba(255,255,255,0.08); padding: 12px; border-radius: 8px;">
                <div style="font-size: 1.8em; font-weight: bold; color: #38bdf8;">{summary['total_g']:,}</div>
                <div style="font-size: 0.85em; text-transform: uppercase; color: #94a3b8;">Total G-Vectors</div>
            </div>
            <div style="background: rgba(255,255,255,0.08); padding: 12px; border-radius: 8px;">
                <div style="font-size: 1.8em; font-weight: bold; color: #4ade80;">{summary['es_count']}</div>
                <div style="font-size: 0.85em; text-transform: uppercase; color: #94a3b8;">Discovered ES Modules</div>
            </div>
            <div style="background: rgba(255,255,255,0.08); padding: 12px; border-radius: 8px;">
                <div style="font-size: 1.8em; font-weight: bold; color: #4ade80;">{summary['coverage_pct']}%</div>
                <div style="font-size: 0.85em; text-transform: uppercase; color: #94a3b8;">G Space Coverage</div>
            </div>
            <div style="background: rgba(255,255,255,0.08); padding: 12px; border-radius: 8px;">
                <div style="font-size: 1.8em; font-weight: bold; color: #fbbf24;">{summary['variance_explained_pct']}%</div>
                <div style="font-size: 0.85em; text-transform: uppercase; color: #94a3b8;">Variance Explained</div>
            </div>
            <div style="background: rgba(255,255,255,0.08); padding: 12px; border-radius: 8px;">
                <div style="font-size: 1.8em; font-weight: bold; color: #a78bfa;">{summary['avg_delta_r2']}</div>
                <div style="font-size: 0.85em; text-transform: uppercase; color: #94a3b8;">Mean Compression ΔR²</div>
            </div>
        </div>
        <div style="margin-top: 14px; font-size: 0.9em; color: #cbd5e1;">
            Completed in <strong>{summary['duration_sec']}s</strong>. Output file: <code>{summary['excel_path']}</code>
        </div>
    </div>
    """

    return (
        summary_html,
        df_es,
        df_matrix,
        df_r2,
        summary['excel_path'],
        "\n".join(log_messages)
    )

def create_app():
    with gr.Blocks(title="Cancer Research TME Algorithm Workbench") as demo:
        gr.Markdown(
            """
            # 🔬 Cancer Research TME Algorithm Workbench
            ### ניתוח Surprisal Analysis, כריית רצפי סימנים מדויקים (ES) ושחזור R²
            מערכת להרצת האלגוריתם על הדאטא המקורי (120 סקשנים), דאטאות הבחינה (Data for algorithm evaluation), או כל תקיית דאטא מותאמת אישית.
            """
        )

        with gr.Row():
            # Controls Sidebar
            with gr.Column(scale=1):
                gr.Markdown("### ⚙️ הגדרות הרצה ודאטא (Configuration)")
                
                dataset_choice = gr.Radio(
                    label="בחירת דאטא סט (Dataset Source)",
                    choices=[
                        "Evaluation Cohort (Data for algorithm evaluation - 6 PDE Samples)",
                        "Original 120-ROI IMC Dataset (All_Tail_Proteins_Tables & All Patients)",
                        "Custom Directory Path (תיקייה מותאמת אישית)"
                    ],
                    value="Evaluation Cohort (Data for algorithm evaluation - 6 PDE Samples)"
                )
                
                custom_dir = gr.Textbox(
                    label="נתיב תקייה מותאם אישית (Custom Directory Path)",
                    placeholder="e.g. C:/data/my_pde_cohort",
                    visible=False
                )

                strategy = gr.Dropdown(
                    label="אסטרטגיית עוגן לבחירה חמדנית (Anchor Strategy)",
                    choices=[
                        "Hybrid (Multi-Objective: Variance x √Count) [WINNER / מומלץ]",
                        "SVD Variance Driver (מקסימום שונות פיזיקלית)",
                        "Max Frequency (אוניברסליות ושכיחות מקסימלית)",
                        "Deep Specificity (רצפים ארוכים 5-8 בעדיפות)"
                    ],
                    value="Hybrid (Multi-Objective: Variance x √Count) [WINNER / מומלץ]"
                )

                with gr.Accordion("הגדרות מתקדמות (Advanced Parameters)", open=False):
                    min_len = gr.Slider(minimum=2, maximum=6, value=3, step=1, label="אורך רצף ES מינימלי (Min Length)")
                    max_len = gr.Slider(minimum=4, maximum=8, value=8, step=1, label="אורך רצף ES מקסימלי (Max Length)")
                    threshold = gr.Slider(minimum=0.00, maximum=0.20, value=0.05, step=0.01, label="סף זנב לסינון רעשים (Tail Threshold)")
                    min_rec = gr.Slider(minimum=1, maximum=5, value=2, step=1, label="חזרתיות מינימלית לווקטורים (Min Recurrence / 2+ Gs)")

                run_btn = gr.Button("🚀 הרץ אלגוריתם (Run Algorithm)", variant="primary", size="lg")

            # Results View
            with gr.Column(scale=2):
                gr.Markdown("### 📊 תוצאות ומדדי ביצועים (Results & Metrics)")
                summary_output = gr.HTML("<div style='padding: 20px; color: #64748b;'>בחר דאטאסט ולחץ 'הרץ אלגוריתם' להפקת דוחות מלאים.</div>")

                with gr.Tabs():
                    with gr.TabItem("📋 מודולי ה-ES שנבחרו (Selected ES Modules)"):
                        es_table = gr.Dataframe(
                            headers=['ES_ID', 'Length', 'Exact_Sign_Sequence', 'G_Covered', 'Variance_Gain', 'Cum_Coverage_Pct'],
                            label="טבלת מודולי ה-ES המנצחים (Direct/Inverse Exact Signs)"
                        )

                    with gr.TabItem("🧩 מטריצת הקצאה נקייה (Conflict-Free Matrix)"):
                        matrix_table = gr.Dataframe(
                            label="הקצאה ללא כפילויות (0 Row Duplicates)"
                        )

                    with gr.TabItem("📈 איכות שחזור R² לכל דאטא (R² Reconstruction)"):
                        r2_table = gr.Dataframe(
                            headers=['ROI', 'Cells', 'Selected_G_Count', 'Selected_Gs', 'R2_full', 'R2_optimal', 'Delta_R2', 'Status'],
                            label="מדדי איכות ברמת התא הבודד"
                        )

                    with gr.TabItem("💾 הורדת דוחות (Deliverables Download)"):
                        excel_file = gr.File(label="הורדת קובץ אקסל מרובה-גיליונות (Master Excel Report)")

                    with gr.TabItem("📜 יומן ריצה (Execution Log)"):
                        log_output = gr.Textbox(label="יומן שלבי הצינור החישובי (Log)", lines=12, max_lines=20)

        # Toggle visibility of custom_dir based on radio selection
        def on_ds_change(val):
            return gr.update(visible=("Custom" in val))
            
        dataset_choice.change(on_ds_change, inputs=[dataset_choice], outputs=[custom_dir])

        # Bind run button
        run_btn.click(
            execute_gui_job,
            inputs=[dataset_choice, custom_dir, strategy, min_len, max_len, threshold, min_rec],
            outputs=[summary_output, es_table, matrix_table, r2_table, excel_file, log_output]
        )

    return demo

if __name__ == '__main__':
    port = 7860
    print(f"Starting Cancer Research Algorithm Web GUI on http://127.0.0.1:{port}...")
    app = create_app()
    app.launch(server_name="127.0.0.1", server_port=port, inbrowser=True)
