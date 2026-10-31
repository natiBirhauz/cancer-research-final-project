# Cancer Research: TME Surprisal Analysis & Exact Sign Sequence (ES) Platform

A mathematically grounded computational framework for deciphering cellular heterogeneity in tumor microenvironments (TME). 

The platform decomposes multiplexed single-cell proteomic expression matrices into discrete biological processes, identifies coordinated driver protein modules with strict directional signs (**Exact Sign Sequences - ES**), eliminates regulatory contradictions, and validates reconstruction fidelity at single-cell resolution using the **Determination Coefficient ($R^2$)**.

---

## ⚡ Quick Start: Running the Web GUI

The platform includes a modern, browser-based user interface.

```bash
# 1. Clone the repository
git clone https://github.com/natiBirhauz/cancer-research-final-project.git
cd cancer-research-final-project

# 2. Install dependencies
pip install numpy pandas openpyxl gradio python-docx

# 3. Launch the Web GUI
python run_gui.py
```
Then open your browser at `http://127.0.0.1:7860`.

---

## 🔬 Core Algorithmic Pipeline

The framework executes five automated mathematical stages:

```text
[Single-Cell Proteomic Data]
             │
             ▼
[Stage 1: Surprisal Analysis Decomposition] ──► ln(X) = G × Lᵀ  (Eigenvectors & Variance Weights)
             │
             ▼
[Stage 2: Driver Tails & Canonical ES Mining]──► G+/G- Tails & Sign Normalization (ES ≡ -ES)
             │
             ▼
[Stage 3: Zero-Contradiction Sign Matching]  ──► Strict Direct (+) & Inverse (-) Polarity Matching
             │
             ▼
[Stage 4: Multi-Objective Greedy Set Cover]  ──► Score = [Σ w(g)] × √(Count) & Disjoint Partitioning
             │
             ▼
[Stage 5: Conflict-Free Matrix & Single-Cell R²]► 1-to-1 Disjoint Allocation & Single-Cell R² Verification
```

### 1. Thermodynamic Surprisal Analysis Decomposition
Multiplexed single-cell expression matrices $X$ are decomposed into orthogonal state variables:
$$\ln(X_{i,j}) = \sum_{\alpha=0}^K G_{i,\alpha} \cdot L_{j,\alpha} = \left( G \times L^T \right)_{i,j}$$
- **$G_\alpha$ (Process Vectors):** Orthonormal protein basis vectors representing coordinated biological programs.
- **$L_\alpha$ (Cellular Amplitudes):** Amplitude coordinate of process $\alpha$ across individual cells.
- **Kinetic Variance Weight $w(g)$:** Normalized physical energy of deviation driving microenvironment states:
$$w(g) = \frac{\|L_k\|^2}{\sum_{\alpha \ge 1} \|L_\alpha\|^2}$$

### 2. Driver Tail Extraction & Canonical Sign Sequences
- **Deviation Tails ($G+ / G-$):** Proteins with loadings exceeding the cutoff threshold $|G_{ip}| \ge 0.05$ represent active processes ($G+$ = upregulation, $G-$ = targeted repression).
- **Canonical Inversion Invariance ($ES \equiv -ES$):** Eigenvectors are invariant under sign inversion ($G \times L^T = (-G) \times (-L)^T$). Signatures are canonically sorted and polarized to prevent redundant representations.
- **Combinatorial Mining:** Mines candidate signatures across lengths 3 to 8, filtering singletons to focus strictly on recurring co-expression modules ($\ge 2$ processes).

### 3. Anti-Contradiction Rule
When evaluating whether a candidate ES module matches a biological process:
1. **Overlap:** Must share at least 2 proteins with the process.
2. **Direct Match (`+`):** All shared proteins exhibit identical signs.
3. **Inverse Match (`-`):** All shared proteins exhibit opposite signs.
4. **Zero Contradictions:** Any contradictory sign (e.g., $P(+)$ vs $P(-)$) immediately invalidates the match.

### 4. Multi-Objective Greedy Set Cover
To balance kinetic energy with biological universality, modules are selected greedily via:
$$\text{Score}(ES) = \left[ \sum_{g \in \text{New}(ES)} w(g) \right] \times \sqrt{|\text{New}(ES)|}$$
Vectors matched by a selected module are permanently locked, guaranteeing a **100% disjoint allocation** with zero duplicate assignments.

### 5. Conflict-Free Matrix & Single-Cell $R^2$ Reconstruction
- **Conflict-Free Matrix:** Produces a strictly disjoint Sample $\times$ Module allocation matrix with zero intra-row duplicates.
- **Single-Cell Determination Coefficient ($R^2$):**
$$R^2 = 1 - \frac{\sum_i (X_{i,j} - \hat{X}_{i,j})^2}{\sum_i (X_{i,j} - \bar{X}_j)^2}$$
- **Compression Penalty ($\Delta R^2$):** $\Delta R^2 = R^2_{full} - R^2_{optimal}$. Measures preservation of original proteomic variance (loss $< 0.05$ indicates high fidelity).

---

## 📂 Supported Input Data Formats

The platform ingests two standardized formats (upload as `.xlsx` files, a `.zip` archive, or specify a directory path):

### Format A: Surprisal Analysis Workbooks
An Excel file (`.xlsx`) containing two sheets:
- **`Gvalues`:** Protein loading matrix ($G_0, G_1, \dots, G_k$) across protein marker rows.
- **`Lambavalues`:** Single-cell amplitude matrix ($L_0, L_1, \dots, L_k$) across individual cell rows.

### Format B: Tail Proteins Spreadsheets
An Excel file (`.xlsx`) with columns named `G1+, G1-, G2+, G2-, ..., Gk+, Gk-`, listing the participating protein markers per process tail.

---

## 📦 Generated Master Deliverables

Every analysis run automatically generates 4 standardized deliverable files:

| Deliverable File | Description |
| :--- | :--- |
| **`Optimal_ES_Result.xlsx`** | Master Set Cover results, conflict-free allocation matrix, and coverage curves. |
| **`Optimal_ES_Result_Signs_R2.xlsx`** | Comprehensive 8-sheet master report including single-cell $R^2$ reconstruction fidelity. |
| **`ES_Combinations.csv`** | Catalog of all candidate mined ES signatures, frequency counts, and participating processes. |
| **`Anchor_Strategies_Comparison.xlsx`** | Benchmark comparison across 6 greedy anchor strategies (Coverage, Variance, $R^2$, Run time). |
| **`Deliverables_Bundle.zip`** | Consolidated ZIP archive containing all 4 master deliverables for 1-click download. |

---

## 📁 Repository Structure

```text
cancer-research-final-project/
├── run_gui.py                     # Primary launcher for the Web GUI workbench
├── web_gui.py                     # Modern Gradio web interface
├── pipeline_engine.py             # Core mathematical pipeline & deliverable engine
├── generate_cohort_comparison.py  # Comparative analysis generator
│
├── Docs/                          # Documentation, reports & deliverables
│   ├── Algorithm_Explanation_Guide.md   # Mathematical specification guide
│   ├── Cohort_Deliverables_Comparison_Report.xlsx # Cross-cohort comparison workbook
│   ├── Cohort_Deliverables_Comparison_Report.md   # Cross-cohort findings summary
│   └── algorithm_workflow_schema.svg             # Pipeline architecture vector diagram
│
├── .gitignore                     # Strict exclusion of raw patient data and binaries
└── README.md                      # Project documentation & algorithm guide
```

---

## 📄 License & Attribution
Academic and research use only. Developed for advanced computational oncology and multiplexed single-cell tissue analysis.
