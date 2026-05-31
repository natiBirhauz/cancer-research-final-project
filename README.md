# Cancer Research Final Project

the code uses protein values from cancer paitents and compares them to find the maximal shared biological sequence

example data:

| G1+   | G1+Value    | G1-  | G1-Value    |
|-------|-------------|------|-------------|
| PD1   | 0.116862203 | LAG3 | -0.82758781 |
| CD123 | 0.128121789 |      |             |
| CD83  | 0.177881823 |      |             |
| HLADR | 0.447148572 |      |             |


what does it do?
Vector Correlation: Uses Cosine Similarity to find matching biological patterns across different patients and ROIs (Regions of Interest).
Anti-Correlation Detection: Preserves true raw signs to identify both Direct (+) and Inverse (-) biological processes.

MaxSharedCompare.py - The main analytical engine. It compares all patient profiles, applies a greedy 1-to-1 matching algorithm to find maximal shared sequences, and exports a detailed Maximal_Sequences_Report.xlsx.

example compare:
| G18 <-> G16 (86.4%, Inverse (-)) | CD8 (0.77, -0.76) | CXCR3 (-0.31, 0.25) | RORgT (-0.28, 0.21) |
| G17 <-> G15 (86.0%, Direct (+)) | Foxp3 (-0.90, -0.81) | CD8 (0.21, 0.35) | Histone (-0.12, 0.13) |


count.py - is a fast utility script that scans the entire dataset to generate a global summary of the top 20 most frequently occurring proteins across all files.

