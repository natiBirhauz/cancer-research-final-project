# Algorithmic & Mathematical Specification: Tumor Microenvironment (TME) Analysis

A comprehensive guide to the mathematical modeling, algorithmic formulation, and single-cell reconstruction principles powering the platform.

---

## Table of Contents
1. [Scientific Context & The Computational Challenge](#1-scientific-context--the-computational-challenge)
2. [Thermodynamic Surprisal Analysis Decomposition](#2-thermodynamic-surprisal-analysis-decomposition)
3. [Process Driver Tails (G+ and G-)](#3-process-driver-tails-g-and-g-)
4. [Exact Sign Sequences (ES) & Inversion Invariance](#4-exact-sign-sequences-es--inversion-invariance)
5. [Combinatorial Sequence Mining](#5-combinatorial-sequence-mining)
6. [Anti-Contradiction Matching Logic](#6-anti-contradiction-matching-logic)
7. [Multi-Objective Greedy Set Cover Optimization](#7-multi-objective-greedy-set-cover-optimization)
8. [Anchor Strategy Benchmarks](#8-anchor-strategy-benchmarks)
9. [Conflict-Free Allocation Matrix](#9-conflict-free-allocation-matrix)
10. [Single-Cell Reconstruction Fidelity (R² and ΔR²)](#10-single-cell-reconstruction-fidelity-r-and-r)
11. [Output Deliverables Architecture](#11-output-deliverables-architecture)

---

## 1. Scientific Context & The Computational Challenge

In modern oncology research, tumors are complex ecosystems termed the **Tumor Microenvironment (TME)**. TMEs comprise heterogeneous populations: malignant cells, immune infiltrating cells, vascular endothelial cells, and cancer-associated fibroblasts (CAFs). 

Multiplexed single-cell spatial proteomics measures dozens of biomarkers across thousands of individual cells while preserving cellular topography. The computational challenge is to:
1. Decompose noisy, high-dimensional protein expression into independent biological processes.
2. Identify a minimal, non-redundant set of recurring protein co-expression modules (**Exact Sign Sequences - ES**).
3. Strictly preserve directional polarity (`+` for upregulation, `-` for downregulation).
4. Reconstruct single-cell expression matrices with high fidelity ($R^2$) and minimal compression loss ($\Delta R^2$).

---

## 2. Thermodynamic Surprisal Analysis Decomposition

Surprisal Analysis models single-cell proteomic expression matrices through information-theoretic state decomposition. For an expression matrix $X$ ($N_{proteins} \times N_{cells}$), log-expression is decomposed into orthogonal components:

$$\ln(X_{i,j}) = \sum_{\alpha=0}^K G_{i,\alpha} \cdot L_{j,\alpha} = \left( G \times L^T \right)_{i,j}$$

- **State 0 ($G_0, L_0$):** Baseline thermodynamic equilibrium shared across cells.
- **Deviation Processes $\alpha \ge 1$ ($G_\alpha, L_\alpha$):** Biological constraints and phenotypic alterations (e.g., hypoxia, immune activation, apoptosis resistance).
- **Process Vectors $G_\alpha$:** Orthonormal protein signatures representing biological programs.
- **Cellular Amplitudes $L_\alpha$:** State coordinates across individual segmented cells.

### Kinetic Variance Weight $w(g)$
The normalized kinetic energy contributed by process $k$ is calculated from cellular amplitudes:

$$\|L_k\|^2 = \sum_{j=1}^{N_{cells}} L_{j,k}^2$$

$$w(g) = \frac{\|L_k\|^2}{\sum_{\alpha \ge 1} \|L_\alpha\|^2}$$

This ensures that $\sum_{\alpha \ge 1} w(g_\alpha) = 1.0$ per sample.

---

## 3. Process Driver Tails (G+ and G-)

For each process vector $G_\alpha$, proteins with loading values exceeding the biological significance threshold $|G_{ip}| \ge 0.05$ define active driver tails:
- **$G+$ Tail:** Upregulated proteins ($G_{ip} \ge +0.05$).
- **$G-$ Tail:** Repressed proteins ($G_{ip} \le -0.05$).

Proteins below the cutoff represent background noise and are excluded from driver tails.

---

## 4. Exact Sign Sequences (ES) & Inversion Invariance

An **Exact Sign Sequence (ES)** is an ordered combination of proteins with fixed regulatory polarity:
$$\text{ES} = \{ P_1(s_1), P_2(s_2), \dots, P_m(s_m) \}, \quad s_i \in \{+, -\}$$

### Inversion Invariance Principle
In SVD/eigen-decomposition, process vectors are sign-invariant:
$$G \times L^T = (-G) \times (-L)^T$$
A biological module where protein A is up and protein B is down in direct orientation ($A(+) | B(-)$) is mathematically equivalent to its inverted conjugate ($A(-) | B(+)$). 

To prevent redundant duplicate signatures, candidate modules are converted into **Canonical Form**:
1. Identify the first protein alphabetically.
2. If its sign is `-`, invert all signs in the module.
3. Sort proteins alphabetically.

---

## 5. Combinatorial Sequence Mining

Candidate signatures of lengths $L \in [3, 8]$ are exhaustively mined across all process tails.
- **Singletons:** Signatures occurring in only 1 process vector are flagged as private noise.
- **Recurring Modules:** Signatures occurring in $\ge 2$ distinct process vectors represent shared biology and are retained for Set Cover optimization.

---

## 6. Anti-Contradiction Matching Logic

When evaluating whether candidate module $\text{ES}$ matches process vector $G$:
1. **Overlap Requirement:** $\text{ES}$ and $G$ must share at least 2 proteins.
2. **Direct Match (`+`):** For all shared proteins, $\text{sign}_{\text{ES}}(P) = \text{sign}_G(P)$.
3. **Inverse Match (`-`):** For all shared proteins, $\text{sign}_{\text{ES}}(P) = -\text{sign}_G(P)$.
4. **Zero Contradictions:** If any shared protein has contradictory signs while another matches, the candidate is strictly disqualified.

---

## 7. Multi-Objective Greedy Set Cover Optimization

The Set Cover problem is solved using a greedy objective function balancing explained variance and biological recurrence:

$$\text{Score}(ES) = \left[ \sum_{g \in \text{New}(ES)} w(g) \right] \times \sqrt{|\text{New}(ES)|}$$

At each iteration:
1. The module maximizing $\text{Score}(ES)$ over uncovered vectors is selected.
2. Its newly covered vectors are permanently locked.
3. The process repeats until complete coverage is reached.

---

## 8. Anchor Strategy Benchmarks

The platform benchmarks 5 anchor objective strategies simultaneously:
1. **SVD Variance Driver:** Maximizes cumulative kinetic energy ($\sum w(g)$).
2. **Max Frequency:** Maximizes vector count without weighting.
3. **Deep Specificity:** Prioritizes complex long modules (lengths 5-8).
4. **Biological Seed:** Seeds optimization around the most co-expressed protein pair.
5. **Hybrid Multi-Objective:** Optimizes variance weighted by the square root of frequency count.

All 5 strategies are evaluated in parallel and reported in `Anchor_Strategies_Comparison.xlsx`.

---

## 9. Conflict-Free Allocation Matrix

To prevent regulatory overlap:
- Each biological vector $g$ is assigned to exactly **one** primary ES module (1-to-1 disjoint allocation).
- A Sample $\times$ Module binary allocation matrix is exported with verified 0 intra-row duplicates.

---

## 10. Single-Cell Reconstruction Fidelity (R² and ΔR²)

To ensure compressed ES modules faithfully retain biological variance:
1. **Full Reconstruction ($R^2_{full}$):** Reconstructs single-cell expression using all available singular components.
2. **Optimal ES Reconstruction ($R^2_{opt}$):** Reconstructs expression using only the selected ES modules.
3. **Compression Loss ($\Delta R^2$):**
   $$\Delta R^2 = R^2_{full} - R^2_{opt}$$
   - **High Fidelity:** $\Delta R^2 < 0.05$ (near-lossless compression).
   - **Acceptable:** $0.05 \le \Delta R^2 \le 0.15$.
   - **Alert:** $\Delta R^2 > 0.15$.

---

## 11. Output Deliverables Architecture

Execution generates 4 master deliverables:
- **`Optimal_ES_Result.xlsx`:** Set Cover selections and conflict-free allocation matrix.
- **`Optimal_ES_Result_Signs_R2.xlsx`:** Multi-sheet master report with single-cell $R^2$ fidelity scores.
- **`ES_Combinations.csv`:** Full catalog of mined candidate ES signatures.
- **`Anchor_Strategies_Comparison.xlsx`:** Side-by-side benchmark overview plus 5 dedicated strategy detail sheets.
- **`Deliverables_Bundle.zip`:** Consolidated archive of all deliverables.
