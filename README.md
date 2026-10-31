# Tumor Microenvironment (TME) Surprisal Analysis & Exact Sign Sequence (ES) Platform

A computational framework and interactive Web GUI for analyzing cellular heterogeneity in tumor microenvironments. The platform decomposes single-cell proteomic expression matrices into coordinated biological programs, identifies robust protein co-expression modules with strict regulatory signs (**Exact Sign Sequences - ES**), eliminates directional contradictions, and evaluates reconstruction fidelity at single-cell resolution using determination coefficients ($R^2$).

---

## ⚡ Quick Start: Running the Web GUI

### Option A: 1-Click Launcher (Windows)
Simply double-click `Launch_GUI.bat` in the project root folder. The browser will open automatically at `http://127.0.0.1:7860`.

### Option B: Command Line (Cross-Platform)
```bash
# 1. Clone the repository
git clone https://github.com/natiBirhauz/cancer-research-final-project.git
cd cancer-research-final-project

# 2. Install dependencies
pip install numpy pandas openpyxl gradio

# 3. Launch the Web application
python run_gui.py
```
Open your browser at `http://127.0.0.1:7860`.

---

## 🔬 Algorithm Overview

The framework processes proteomic matrices through five stages:

```text
[Input Proteomic Data]
          │
          ▼
[1. Surprisal Analysis]       ──► ln(X) = G × Lᵀ (Basis Vectors & Variance Weights)
          │
          ▼
[2. Tail & ES Module Mining]  ──► Extraction of G+/G- Tails & Canonical Inversion (ES ≡ -ES)
          │
          ▼
[3. Sign Consistency Filter]  ──► Anti-Contradiction Rule (Strict Direct + / Inverse - Alignment)
          │
          ▼
[4. Greedy Set Cover]         ──► Multi-Objective Optimization: Score = Variance × √(Count)
          │
          ▼
[5. Single-Cell R² Fidelity]  ──► Conflict-Free Matrix & Cell-Level Determination Coefficient (R²)
```

### 1. Surprisal Analysis Decomposition
Single-cell protein expression matrices $X$ are decomposed into orthogonal state variables:
$$\ln(X_{i,j}) = \sum_{\alpha=0}^K G_{i,\alpha} \cdot L_{j,\alpha}$$
- **$G_\alpha$ (Process Vectors):** Orthonormal protein basis vectors representing coordinated programs.
- **$L_\alpha$ (Cellular Amplitudes):** Cell-level state coordinates.
- **Variance Weights $w(g)$:** Normalized kinetic energy carried by each process.

### 2. Driver Tail Extraction & Canonical Sign Sequences
- **Deviation Tails ($G+ / G-$):** Proteins with loadings $|G| \ge 0.05$ represent active regulatory drivers ($+$ = upregulation, $-$ = repression).
- **Inversion Invariance ($ES \equiv -ES$):** Eigenvectors are invariant under sign inversion ($G \times L^T = (-G) \times (-L)^T$). Modules are canonically sorted and polarized.
- **Combinatorial Mining:** Mines candidate signatures across lengths 3 to 8, filtering singletons to focus on recurring co-expression patterns.

### 3. Anti-Contradiction Matching
When evaluating whether a candidate ES module covers a biological vector:
- **Direct Match (`+`):** All shared proteins share identical regulatory signs.
- **Inverse Match (`-`):** All shared proteins share reversed regulatory signs.
- **Zero Contradictions:** Any discordant sign immediately disqualifies the module.

### 4. Multi-Objective Greedy Set Cover
Modules are iteratively selected by optimizing:
$$\text{Score}(ES) = \left[ \sum_{g \in \text{New}(ES)} w(g) \right] \times \sqrt{|\text{New}(ES)|}$$
Vectors covered by selected modules are locked, ensuring a strictly disjoint allocation with zero duplicate assignments.

### 5. Conflict-Free Allocation & Single-Cell $R^2$ Fidelity
- **Conflict-Free Matrix:** Produces a 1-to-1 disjoint allocation matrix where each vector is assigned to at most one module.
- **Reconstruction Fidelity ($R^2$):** Measures how accurately the selected ES modules reconstruct single-cell protein expression compared to the full matrix.
- **Compression Loss ($\Delta R^2$):** $\Delta R^2 = R^2_{full} - R^2_{optimal}$. A loss $< 0.05$ indicates high-fidelity preservation.

---

## 📂 Input Data Formats

The Web GUI accepts files via drag-and-drop (`.xlsx` spreadsheets or `.zip` archives containing multiple files):

### Format A: Surprisal Matrix Workbooks
An Excel workbook (`.xlsx`) containing two sheets:
- **`Gvalues`:** Protein loading matrix (rows = proteins, columns = $G_0, G_1, \dots, G_k$).
- **`Lambavalues`:** Single-cell amplitude matrix (rows = cells, columns = $L_0, L_1, \dots, L_k$).

### Format B: Tail Proteins Spreadsheets
An Excel workbook (`.xlsx`) with columns named `G1+, G1-, G2+, G2-, ..., Gk+, Gk-`, listing participating proteins per process tail.

---

## 📦 Generated Deliverables

Every execution produces 4 standardized deliverables and a consolidated download bundle:

| Deliverable File | Description |
| :--- | :--- |
| **`Optimal_ES_Result.xlsx`** | Set Cover results, conflict-free allocation matrix, and cumulative coverage curves. |
| **`Optimal_ES_Result_Signs_R2.xlsx`** | Master report containing single-cell $R^2$ reconstruction fidelity and protein sign tables. |
| **`ES_Combinations.csv`** | Comprehensive catalog of all mined candidate ES combinations and recurrence statistics. |
| **`Anchor_Strategies_Comparison.xlsx`** | Multi-strategy benchmark comparing all 5 anchor algorithms with dedicated detail sheets for each. |
| **`Deliverables_Bundle.zip`** | Consolidated ZIP archive containing all master deliverables. |

---

## 📁 Repository Structure

```text
cancer-research-final-project/
├── Launch_GUI.bat                 # 1-Click launcher for Windows
├── run_gui.py                     # Primary launcher for the Web GUI workbench
├── web_gui.py                     # Gradio web application interface
├── pipeline_engine.py             # Core algorithm engine (mining, set cover, R² evaluation)
├── test_gui_execution.py          # Automated verification test suite
├── Sample_Input_Template.xlsx     # Standardized input data template for users
│
├── Docs/                          # Documentation & reference materials
│   ├── Deliverables_Data_Dictionary.xlsx  # Complete data dictionary for all deliverables
│   ├── Algorithm_Explanation_Guide.md     # In-depth algorithmic methodology guide
│   ├── ES_Combinations_Specification.md   # Exact sign sequence technical specification
│   └── algorithm_workflow_schema.svg     # Algorithm workflow diagram
│
├── .gitignore                     # Git exclusion rules
└── README.md                      # Project documentation
```
