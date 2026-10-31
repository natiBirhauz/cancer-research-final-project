# Cancer Research: Surprisal Analysis, Exact Sign Sequences (ES) & Single-Cell R² Reconstruction

An advanced agentic, mathematically grounded computational framework for deciphering cellular heterogeneity in tumor microenvironments through multiplexed single-cell proteomic imaging.

The algorithm decomposes high-dimensional cellular expression matrices into discrete, orthogonal biological processes, extracts active driver protein signatures (**Exact Sign Sequences - ES**), preserves strict regulatory polarity (`+` for upregulation / `-` for downregulation), enforces zero biological contradictions, and reconstructs original proteomic states with minimal information loss verified at single-cell resolution via the **Determination Coefficient ($R^2$)**.

---

## 🔬 Mathematical & Algorithmic Foundations

The framework operates as a rigorous 5-stage computational pipeline:

```text
[Multiplexed Expression Data]
           │
           ▼
[Stage 1: Surprisal Analysis SVD Decomposition] ──► ln(X) = G × Lᵀ & Variance Weights w(g)
           │
           ▼
[Stage 2: Driver Tails & Canonical ES Mining]   ──► G+/G- Tails & Inversion Invariant ES ≡ -ES
           │
           ▼
[Stage 3: Anti-Contradiction Match Model]       ──► Direct (+), Inverse (-), Zero Contradictions
           │
           ▼
[Stage 4: Multi-Objective Greedy Set Cover]     ──► Score = [Σ w(g)] × √(Count) & Disjoint Allocation
           │
           ▼
[Stage 5: Conflict-Free Matrix & Single-Cell R²]──► 0 Row Duplicates, Cell-Level R² & ΔR² Metric
```

### Stage 1: Thermodynamic Surprisal Analysis Decomposition
Multiplexed single-cell proteomic expression matrices $X$ ($N_{proteins} \times N_{cells}$) are decomposed into orthogonal biological state variables:
$$\ln(X_{i,j}) = \sum_{\alpha=0}^K G_{i,\alpha} \cdot L_{j,\alpha} = \left( G \times L^T \right)_{i,j}$$
* **$G_\alpha$ (Process Vectors):** Orthonormal basis vectors representing coordinated multi-protein biological programs.
* **$L_\alpha$ (Cellular Amplitudes):** The amplitude coordinate of process $\alpha$ in each individual cell.
* **Baseline Equilibrium ($G_0, L_0$):** State of maximum entropy / unperturbed tissue baseline.
* **Variance Weights $w(g)$:** Physical energy ratio driving tissue deviation:
$$w(g) = \frac{\|L_k\|^2}{\sum_{\alpha \ge 1} \|L_\alpha\|^2}$$

### Stage 2: Driver Tail Extraction & Canonical Normalization
* **Tails ($G+$ / $G-$):** Proteins with extreme deviation coefficients drive the process ($G+$ = active overexpression, $G-$ = targeted repression), while near-zero coefficients represent non-participating background noise.
* **Canonical Form ($ES \equiv -ES$):** Eigenvectors are invariant under sign inversion ($G \times L^T = (-G) \times (-L)^T$). To prevent redundant representation of identical physical pathways, signatures are sorted alphabetically and normalized so the leading protein carries a positive sign.
* **Combinatorial Mining:** Mines all valid sub-combinations from length 3 up to the maximum panel length.
* **Singleton Filtering:** Discards singleton combinations (appearing in only 1 vector) to eliminate measurement noise and focus strictly on reproducible multi-sample biological attractors ($\ge 2$ vectors).

### Stage 3: Biological Compatibility & Zero Contradiction Rule
When evaluating whether a candidate ES module matches a biological process vector $G$:
1. **Overlap Requirement:** Requires a common overlap of at least 2 proteins: $\text{Common} = \text{Prots}(ES) \cap \text{Prots}(G)$.
2. **Direct Match (`+`):** All common proteins share identical signs.
3. **Inverse Match (`-`):** All common proteins share inverted signs.
4. **Strict Rejection (Zero Contradictions):** If even a single shared protein contradicts in sign (e.g., $P_1(+)$ vs. $P_1(-)$), the match is immediately disqualified.

### Stage 4: Multi-Objective Greedy Optimization & Disjoint Partitioning
Standard Set Cover heuristics fail in proteomic systems because deviation processes carry radically different variance energies. The framework applies a balanced Multi-Objective scoring function:
$$\text{Score}(ES) = \left[ \sum_{g \in \text{New}(ES)} w(g) \right] \times \sqrt{|\text{New}(ES)|}$$
* **Variance Term ($\sum w(g)$):** Maximizes physical variance explained and single-cell reconstruction fidelity.
* **Prevalence Term ($\sqrt{\text{Count}}$):** Favors broadly conserved biological processes across the population while avoiding generic short motifs.
* **Disjoint Allocation:** Vectors captured by a selected module are permanently locked and removed from subsequent selection, guaranteeing $100\%$ disjoint assignment with zero duplicate allocations across modules.

### Stage 5: Conflict-Free Matrix & Cell-Level $R^2$ Validation
For each analyzed sample and each module column:
* At most one representative vector is assigned per sample per module, locking out intra-row duplicates (0 duplicate row assignments).
* **Single-Cell Determination Coefficient ($R^2$):**
$$R^2_j = 1 - \frac{\sum_i (X_{i,j} - \hat{X}_{i,j})^2}{\sum_i (X_{i,j} - \bar{X}_j)^2}$$
* **$R^2_{full}$ (Baseline):** Reconstruction using all available decomposition components.
* **$R^2_{optimal}$ (Compressed Representation):** Reconstruction using the unperturbed baseline ($G_0, L_0$) plus only the selected ES modules.
* **Compression Penalty ($\Delta R^2$):** Quantifies exact biological information loss:
$$\Delta R^2 = R^2_{full} - R^2_{optimal}$$
High-fidelity preservation is defined by $\Delta R^2 < 0.05$ (minimal loss) and $0.05 \le \Delta R^2 \le 0.15$ (acceptable trade-off).

---

## 🖥️ Web-Based Graphical User Interface (Web GUI)

The platform provides a browser-based interactive workbench. The algorithm engine is synced directly with this Git repository:

### Launching the Web GUI
```bash
# Launch interactive web application
python run_gui.py

# Launch on custom port (default: 7860)
python run_gui.py --port 8080
```

### Key GUI Features
1. **Direct Git Integration:** Connects to `https://github.com/natiBirhauz/cancer-research-final-project.git`. Features a one-click `Pull Latest Algorithm from Git` button to synchronize and verify the active algorithm version.
2. **User-Provided Data Ingestion:** Purely user-driven input. Users upload their own Excel spreadsheets (`.xlsx`), multi-sample ZIP archives, or specify local directory paths. No data is hardcoded or pre-assumed.
3. **Unbiased Strategy Selection:** Choose neutrally between multiple objective strategies (Hybrid Multi-Objective, SVD Variance Driver, Max Frequency, Deep Specificity) without historical bias.
4. **Interactive Deliverables & Inspection:**
   * Dynamic summary cards (samples, vectors, modules, coverage %, variance %, mean $R^2$).
   * Interactive table of discovered canonical ES modules with exact polarities (`+`/`-`).
   * Conflict-free sample-to-module allocation matrix (0 row duplicates).
   * Single-cell $R^2$ fidelity scores per sample.
   * Direct download of the comprehensive master Excel report (`.xlsx`).

---

## 📂 Supported Input Data Formats

The algorithm accepts proteomic data structured in standard formats:

### Format 1: Surprisal Analysis Workbooks (Recommended)
An Excel file (`.xlsx`) containing two sheets:
* **`Gvalues`:** Process eigenvectors. Columns $G_0, G_1, \dots, G_k$ across protein marker rows.
* **`Lambavalues`:** Cellular coordinates. Columns $L_0, L_1, \dots, L_k$ across individual cell rows.

### Format 2: Precomputed Tail Protein Tables
An Excel file (`.xlsx` or `.csv`) with driver tail columns named `G1+`, `G1-`, `G2+`, `G2-`, etc., specifying positive and negative driver markers per process.

---

## 🛠️ Repository Architecture

```text
cancer-research-final-project/
├── run_gui.py                     # Primary launcher for the Web GUI workbench
├── web_gui.py                     # Web interface implementation (Gradio-based)
├── pipeline_engine.py             # Core mathematical pipeline engine for user data
├── generate_cohort_comparison.py  # Comparative analysis generator across experimental designs
│
├── Docs/                          # Documentation, vector schemas & reports
│   ├── algorithm_workflow_schema.svg # Vector architecture flow diagram
│   ├── Algorithm_Explanation_Guide.md# Comprehensive mathematical guide
│   ├── Algorithm_Explanation_Guide.html# Interactive HTML guide with KaTeX formulas
│   ├── Algorithm_Explanation_Guide.docx# Formatted Microsoft Word deliverable
│   ├── Cohort_Comparison_Report.xlsx # Multi-sheet comparative benchmark spreadsheet
│   ├── Cohort_Comparison_Report.html # Standalone HTML comparative report
│   └── Cohort_Comparison_Report.md   # Markdown summary of cross-cohort findings
│
├── Archive_Old_Project/           # Archived historical scripts and legacy workflows
├── .gitignore                     # Strict exclusion of all data, binary, and cache files
└── README.md                      # Technical project documentation & methodology
```

---

## 📋 Software Requirements & Installation

```bash
# Clone the repository
git clone https://github.com/natiBirhauz/cancer-research-final-project.git
cd cancer-research-final-project

# Install core dependencies
pip install numpy pandas openpyxl gradio python-docx
```
