"""
app_gui.py
----------
Native Desktop Graphical User Interface (Tkinter) for the Cancer Research TME
Surprisal Analysis & Exact Sign Sequence (ES) Mining Platform.

Runs as a native Windows desktop application with zero external browser dependencies.
Includes:
- Data source selector & folder browser
- Parameter tuning
- Non-blocking threaded execution with live progress bar
- Tabbed results viewer (ES Modules, Allocation Matrix, R2 Metrics)
- One-click launcher to open exported Excel and HTML deliverables directly
"""

import os
import sys
import threading
import time
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from pipeline_engine import run_pipeline

class CancerResearchGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Cancer Research TME Algorithm Workbench")
        self.geometry("1100x750")
        self.minsize(950, 600)
        
        self.last_excel_path = None
        self.is_running = False
        
        self.init_style()
        self.create_widgets()
        
    def init_style(self):
        self.style = ttk.Style(self)
        # Attempt to use 'clam' or 'vista' for modern native look
        available_themes = self.style.theme_names()
        if 'vista' in available_themes:
            self.style.theme_use('vista')
        elif 'clam' in available_themes:
            self.style.theme_use('clam')
            
        self.style.configure('Header.TLabel', font=('Segoe UI', 15, 'bold'), foreground='#1F4E79')
        self.style.configure('SubHeader.TLabel', font=('Segoe UI', 10), foreground='#555555')
        self.style.configure('Primary.TButton', font=('Segoe UI', 11, 'bold'), padding=6)
        self.style.configure('Action.TButton', font=('Segoe UI', 10), padding=4)
        self.style.configure('MetricVal.TLabel', font=('Segoe UI', 16, 'bold'), foreground='#1F4E79')
        self.style.configure('MetricLbl.TLabel', font=('Segoe UI', 8), foreground='#666666')

    def create_widgets(self):
        # Top Header Banner
        header_frame = ttk.Frame(self, padding=(15, 10, 15, 5))
        header_frame.pack(fill=tk.X)
        
        ttk.Label(
            header_frame, 
            text="🔬 Cancer Research TME Algorithm Workbench", 
            style='Header.TLabel'
        ).pack(anchor=tk.W)
        ttk.Label(
            header_frame, 
            text="Surprisal Analysis, Exact Sign Sequences (ES) Mining & Cell-Level R² Reconstruction", 
            style='SubHeader.TLabel'
        ).pack(anchor=tk.W)
        ttk.Separator(self, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=(5, 10))

        # Main Layout: Left Sidebar (Controls) | Right Content (Results)
        main_paned = ttk.PanedWindow(self, orient=tk.HORIZONTAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Left Control Panel
        control_frame = ttk.Frame(main_paned, padding=10, width=340)
        main_paned.add(control_frame, weight=0)

        # Right Content Panel
        content_frame = ttk.Frame(main_paned, padding=10)
        main_paned.add(content_frame, weight=1)

        self.setup_controls(control_frame)
        self.setup_content(content_frame)

    def setup_controls(self, parent):
        # 1. Dataset Selection
        grp_ds = ttk.LabelFrame(parent, text=" 📂 Data Source / מקור הנתונים ", padding=10)
        grp_ds.pack(fill=tk.X, pady=(0, 10))

        self.var_ds = tk.StringVar(value="evaluation")
        ttk.Radiobutton(
            grp_ds, 
            text="Evaluation Cohort (6 Samples / 48k Cells)", 
            variable=self.var_ds, 
            value="evaluation",
            command=self.toggle_custom_path
        ).pack(anchor=tk.W, pady=2)
        
        ttk.Radiobutton(
            grp_ds, 
            text="Original 120-ROI IMC Dataset", 
            variable=self.var_ds, 
            value="original",
            command=self.toggle_custom_path
        ).pack(anchor=tk.W, pady=2)

        ttk.Radiobutton(
            grp_ds, 
            text="Custom Dataset Folder...", 
            variable=self.var_ds, 
            value="custom",
            command=self.toggle_custom_path
        ).pack(anchor=tk.W, pady=2)

        self.frame_custom = ttk.Frame(grp_ds)
        self.entry_custom = ttk.Entry(self.frame_custom, width=24)
        self.entry_custom.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 4))
        self.btn_browse = ttk.Button(self.frame_custom, text="Browse...", width=8, command=self.browse_folder)
        self.btn_browse.pack(side=tk.RIGHT)

        # 2. Strategy Selection
        grp_strat = ttk.LabelFrame(parent, text=" 🎯 Anchor Strategy / אסטרטגיית עוגן ", padding=10)
        grp_strat.pack(fill=tk.X, pady=(0, 10))

        self.var_strat = tk.StringVar(value="Hybrid")
        strat_combo = ttk.Combobox(
            grp_strat, 
            textvariable=self.var_strat, 
            state='readonly',
            values=[
                "Hybrid (Variance x √Count) [WINNER]",
                "SVD_Variance (Pure Physical Variance)",
                "Max_Frequency (Max G Coverage)",
                "Deep_Specificity (Lengths 5-8 Core)"
            ]
        )
        strat_combo.current(0)
        strat_combo.pack(fill=tk.X, pady=2)

        # 3. Parameters
        grp_params = ttk.LabelFrame(parent, text=" ⚙️ Hyperparameters / פרמטרים ", padding=10)
        grp_params.pack(fill=tk.X, pady=(0, 10))

        # Min ES Length
        row_min = ttk.Frame(grp_params)
        row_min.pack(fill=tk.X, pady=2)
        ttk.Label(row_min, text="Min ES Length:").pack(side=tk.LEFT)
        self.spin_min_len = ttk.Spinbox(row_min, from_=2, to=6, width=5)
        self.spin_min_len.set(3)
        self.spin_min_len.pack(side=tk.RIGHT)

        # Max ES Length
        row_max = ttk.Frame(grp_params)
        row_max.pack(fill=tk.X, pady=2)
        ttk.Label(row_max, text="Max ES Length:").pack(side=tk.LEFT)
        self.spin_max_len = ttk.Spinbox(row_max, from_=4, to=8, width=5)
        self.spin_max_len.set(8)
        self.spin_max_len.pack(side=tk.RIGHT)

        # Tail Cutoff
        row_th = ttk.Frame(grp_params)
        row_th.pack(fill=tk.X, pady=2)
        ttk.Label(row_th, text="Tail Cutoff Threshold:").pack(side=tk.LEFT)
        self.spin_th = ttk.Spinbox(row_th, from_=0.0, to=0.20, increment=0.01, width=5)
        self.spin_th.set(0.05)
        self.spin_th.pack(side=tk.RIGHT)

        # Min Recurrence
        row_rec = ttk.Frame(grp_params)
        row_rec.pack(fill=tk.X, pady=2)
        ttk.Label(row_rec, text="Min G Recurrence (2+):").pack(side=tk.LEFT)
        self.spin_rec = ttk.Spinbox(row_rec, from_=1, to=5, width=5)
        self.spin_rec.set(2)
        self.spin_rec.pack(side=tk.RIGHT)

        # 4. Action Button
        self.btn_run = ttk.Button(
            parent, 
            text="🚀 Run Pipeline / הרץ אלגוריתם", 
            style='Primary.TButton',
            command=self.start_pipeline_thread
        )
        self.btn_run.pack(fill=tk.X, pady=(10, 5))

        # Progress bar & status
        self.progress_bar = ttk.Progressbar(parent, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=5)

        self.lbl_status = ttk.Label(parent, text="Ready", foreground="#444444")
        self.lbl_status.pack(anchor=tk.W)

    def setup_content(self, parent):
        # Metric Banner Cards
        self.cards_frame = ttk.Frame(parent)
        self.cards_frame.pack(fill=tk.X, pady=(0, 10))

        self.metric_cards = {}
        metrics = [
            ('rois', 'Datasets / ROIs', '-'),
            ('g_cov', 'G Coverage', '-'),
            ('es_cnt', 'ES Modules', '-'),
            ('var', 'Variance Explained', '-'),
            ('r2', 'Mean R² Optimal', '-')
        ]
        for key, title, init_v in metrics:
            f = ttk.Frame(self.cards_frame, padding=8, relief='solid', borderwidth=1)
            f.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=3)
            val_lbl = ttk.Label(f, text=init_v, style='MetricVal.TLabel', anchor='center')
            val_lbl.pack(fill=tk.X)
            lbl = ttk.Label(f, text=title, style='MetricLbl.TLabel', anchor='center')
            lbl.pack(fill=tk.X)
            self.metric_cards[key] = val_lbl

        # Tabbed Viewer
        self.notebook = ttk.Notebook(parent)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Tab 1: ES Modules
        self.tab_es = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_es, text="📋 Selected ES Modules")
        self.tree_es = self.create_treeview(
            self.tab_es, 
            columns=('ES_ID', 'Length', 'Exact_Sign_Sequence', 'G_Covered', 'Variance_Gain', 'Cum_Coverage_Pct'),
            headings=('ES ID', 'Len', 'Exact Sign Sequence Pattern', 'Gs', 'Gain', 'Cum %')
        )

        # Tab 2: Allocation Matrix
        self.tab_matrix = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_matrix, text="🧩 Conflict-Free Matrix")
        self.text_matrix = tk.Text(self.tab_matrix, wrap=tk.NONE, font=('Consolas', 10))
        scroll_my = ttk.Scrollbar(self.tab_matrix, orient=tk.VERTICAL, command=self.text_matrix.yview)
        scroll_mx = ttk.Scrollbar(self.tab_matrix, orient=tk.HORIZONTAL, command=self.text_matrix.xview)
        self.text_matrix.configure(xscrollcommand=scroll_mx.set, yscrollcommand=scroll_my.set)
        scroll_my.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_mx.pack(side=tk.BOTTOM, fill=tk.X)
        self.text_matrix.pack(fill=tk.BOTH, expand=True)

        # Tab 3: R2 Reconstruction
        self.tab_r2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_r2, text="📈 R² Reconstruction Quality")
        self.tree_r2 = self.create_treeview(
            self.tab_r2,
            columns=('ROI', 'Cells', 'Selected_G_Count', 'Selected_Gs', 'R2_full', 'R2_optimal', 'Delta_R2', 'Status'),
            headings=('ROI/Dataset', 'Cells', 'Gs', 'Selected G Components', 'R² Full', 'R² Optimal', 'ΔR²', 'Status')
        )

        # Tab 4: Execution Log
        self.tab_log = ttk.Frame(self.notebook)
        self.notebook.add(self.tab_log, text="📜 Live Execution Log")
        self.text_log = tk.Text(self.tab_log, wrap=tk.WORD, font=('Consolas', 9), bg='#F8F9FA')
        scroll_ly = ttk.Scrollbar(self.tab_log, orient=tk.VERTICAL, command=self.text_log.yview)
        self.text_log.configure(yscrollcommand=scroll_ly.set)
        scroll_ly.pack(side=tk.RIGHT, fill=tk.Y)
        self.text_log.pack(fill=tk.BOTH, expand=True)

        # Bottom Bar: Open Deliverable Button
        bottom_bar = ttk.Frame(parent, padding=(0, 8, 0, 0))
        bottom_bar.pack(fill=tk.X)
        
        self.btn_open_excel = ttk.Button(
            bottom_bar, 
            text="📊 Open Master Excel Deliverable (Docs/*.xlsx)", 
            style='Action.TButton',
            state=tk.DISABLED,
            command=self.open_excel_file
        )
        self.btn_open_excel.pack(side=tk.LEFT, padx=4)

        self.btn_open_html = ttk.Button(
            bottom_bar, 
            text="🌐 Open HTML Summary Report", 
            style='Action.TButton',
            command=self.open_html_file
        )
        self.btn_open_html.pack(side=tk.LEFT, padx=4)

    def create_treeview(self, parent, columns, headings):
        tree = ttk.Treeview(parent, columns=columns, show='headings')
        for col, head in zip(columns, headings):
            tree.heading(col, text=head)
            tree.column(col, anchor=tk.CENTER, width=80)
            
        tree.column(columns[2], width=350, anchor=tk.W)  # Wider for sequence pattern
        
        vsb = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=tree.yview)
        hsb = ttk.Scrollbar(parent, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(xscrollcommand=hsb.set, yscrollcommand=vsb.set)
        
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        tree.pack(fill=tk.BOTH, expand=True)
        return tree

    def toggle_custom_path(self):
        if self.var_ds.get() == "custom":
            self.frame_custom.pack(fill=tk.X, pady=(4, 0))
        else:
            self.frame_custom.pack_forget()

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select Dataset Directory")
        if folder:
            self.entry_custom.delete(0, tk.END)
            self.entry_custom.insert(0, folder)

    def log_message(self, msg):
        self.text_log.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {msg}\n")
        self.text_log.see(tk.END)
        self.lbl_status.config(text=msg[:60])

    def start_pipeline_thread(self):
        if self.is_running:
            return
            
        ds = self.var_ds.get()
        custom_path = self.entry_custom.get().strip() if ds == "custom" else None
        if ds == "custom" and not os.path.isdir(custom_path or ''):
            messagebox.showerror("Error", "Please select a valid custom dataset directory.")
            return

        strat_val = self.var_strat.get()
        strat_key = 'Hybrid'
        if 'SVD' in strat_val: strat_key = 'SVD_Variance'
        elif 'Frequency' in strat_val: strat_key = 'Max_Frequency'
        elif 'Deep' in strat_val: strat_key = 'Deep_Specificity'

        min_len = int(self.spin_min_len.get())
        max_len = int(self.spin_max_len.get())
        th = float(self.spin_th.get())
        min_rec = int(self.spin_rec.get())

        self.is_running = True
        self.btn_run.config(state=tk.DISABLED)
        self.progress_bar['value'] = 0
        self.text_log.delete('1.0', tk.END)

        def worker():
            def progress_cb(pct, msg):
                self.after(0, lambda: self.update_progress(pct, msg))

            try:
                summary, df_es, df_matrix, df_r2 = run_pipeline(
                    data_source=ds,
                    custom_path=custom_path,
                    min_es_length=min_len,
                    max_es_length=max_len,
                    threshold=th,
                    min_g_recurrence=min_rec,
                    strategy=strat_key,
                    progress_callback=progress_cb
                )
                self.after(0, lambda: self.on_pipeline_success(summary, df_es, df_matrix, df_r2))
            except Exception as e:
                err = str(e)
                self.after(0, lambda: self.on_pipeline_error(err))

        threading.Thread(target=worker, daemon=True).start()

    def update_progress(self, pct, msg):
        if pct is not None:
            self.progress_bar['value'] = int(pct * 100)
        self.log_message(msg)

    def on_pipeline_success(self, summary, df_es, df_matrix, df_r2):
        self.is_running = False
        self.btn_run.config(state=tk.NORMAL)
        self.progress_bar['value'] = 100
        self.last_excel_path = summary['excel_path']
        self.btn_open_excel.config(state=tk.NORMAL)

        # Update metric cards
        self.metric_cards['rois'].config(text=str(summary['total_rois']))
        self.metric_cards['g_cov'].config(text=f"{summary['coverage_pct']}%")
        self.metric_cards['es_cnt'].config(text=str(summary['es_count']))
        self.metric_cards['var'].config(text=f"{summary['variance_explained_pct']}%")
        self.metric_cards['r2'].config(text=str(summary['avg_r2_optimal']))

        # Populate ES treeview
        for item in self.tree_es.get_children():
            self.tree_es.delete(item)
        for _, row in df_es.iterrows():
            self.tree_es.insert('', tk.END, values=list(row))

        # Populate Matrix text
        self.text_matrix.delete('1.0', tk.END)
        self.text_matrix.insert(tk.END, df_matrix.to_string(index=False))

        # Populate R2 treeview
        for item in self.tree_r2.get_children():
            self.tree_r2.delete(item)
        for _, row in df_r2.iterrows():
            vals = [
                row['ROI'],
                f"{row['Cells']:,}",
                row['Selected_G_Count'],
                row['Selected_Gs'],
                f"{row['R2_full']:.4f}",
                f"{row['R2_optimal']:.4f}",
                f"{row['Delta_R2']:.4f}",
                row['Status']
            ]
            self.tree_r2.insert('', tk.END, values=vals)

        messagebox.showinfo("Success", f"Pipeline completed successfully in {summary['duration_sec']}s!\nSaved to: {summary['excel_path']}")

    def on_pipeline_error(self, err_msg):
        self.is_running = False
        self.btn_run.config(state=tk.NORMAL)
        self.log_message(f"ERROR: {err_msg}")
        messagebox.showerror("Pipeline Error", err_msg)

    def open_excel_file(self):
        if self.last_excel_path and os.path.exists(self.last_excel_path):
            os.startfile(self.last_excel_path)
        else:
            cand = os.path.join('Docs', 'Evaluation_Datasets_Report.xlsx')
            if os.path.exists(cand):
                os.startfile(cand)
            else:
                messagebox.showwarning("Notice", "Excel report file not found yet.")

    def open_html_file(self):
        cand = os.path.join('Docs', 'Evaluation_Datasets_Report.html')
        if not os.path.exists(cand):
            cand = os.path.join('Docs', 'Algorithm_Full_Summary.html')
        if os.path.exists(cand):
            os.startfile(cand)
        else:
            messagebox.showwarning("Notice", "HTML report not found.")

def main():
    app = CancerResearchGUI()
    app.mainloop()

if __name__ == '__main__':
    main()
