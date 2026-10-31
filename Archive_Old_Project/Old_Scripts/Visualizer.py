import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import re
import glob
import os

def load_and_parse_data():
    """
    Finds and loads the most recently generated Maximal Sequences Report.
    """
    list_of_files = glob.glob('Maximal_Sequences_Report_*.xlsx')
    if not list_of_files:
        print("Error: Could not find any Maximal_Sequences_Report Excel files.")
        return None
    
    latest_file = max(list_of_files, key=os.path.getctime)
    print(f"\n--- Loading data from: {latest_file} ---\n")
    return pd.read_excel(latest_file)

def extract_all_matches(details_text):
    """
    Parses the detailed text cell from the Excel report to extract 
    all G-components, match types, and individual protein statistics.
    """
    parsed_data = []
    lines = str(details_text).split('\n')
    
    for line in lines:
        line = line.strip()
        if not line: continue
        
        # Extract the header of the match (e.g., G7 <-> G7 (98.2%, Direct (+)))
        header_match = re.match(r"(G\d+)\s*<->\s*(G\d+)\s*\(([\d\.]+)%,\s*([^)]+)\)", line)
        if not header_match: continue
            
        ga, gb, sim, match_type = header_match.groups()
        
        # Extract all proteins and their corresponding normalized values
        pattern = r"([A-Za-z0-9_\-]+)\s*\(([\d\.]+)%\s*Match\s*\|\s*(?:Norm_)?A:\s*([+-]?[\d\.]+),\s*(?:Norm_)?B:\s*([+-]?[\d\.]+)\)"
        proteins = re.findall(pattern, line)
        
        parsed_data.append({
            'G_A': ga,
            'G_B': gb,
            'Sim': float(sim),
            'Type': match_type,
            'Proteins': proteins
        })
        
    return parsed_data

def build_dashboard(pA, pB, parsed_data):
    """
    Generates a 4-panel dashboard containing advanced visualizations 
    for the selected patient pair.
    """
    fig, axs = plt.subplots(2, 2, figsize=(16, 11))
    fig.suptitle(f"TME Comprehensive Match Analysis: {pA} vs {pB}", fontsize=18, fontweight='bold')
    
    # Map G-component names to numeric indices for plotting coordinates
    unique_ga = sorted(list(set([d['G_A'] for d in parsed_data])))
    unique_gb = sorted(list(set([d['G_B'] for d in parsed_data])))
    map_ga = {k: v for v, k in enumerate(unique_ga)}
    map_gb = {k: v for v, k in enumerate(unique_gb)}

    # ==========================================
    # 1. Bubble Chart (Match Strength)
    # ==========================================
    ax1 = axs[0, 0]
    for d in parsed_data:
        color = '#1f77b4' if 'Direct' in d['Type'] else '#d62728'
        # Bubble size correlates with the similarity percentage
        bubble_size = (d['Sim'] - 70) * 40 
        ax1.scatter(map_ga[d['G_A']], map_gb[d['G_B']], s=bubble_size, color=color, alpha=0.6, edgecolors='black')
        ax1.text(map_ga[d['G_A']], map_gb[d['G_B']], f"{d['Sim']}%", ha='center', va='center', fontsize=9, fontweight='bold')

    ax1.set_xticks(range(len(unique_ga)))
    ax1.set_xticklabels(unique_ga)
    ax1.set_yticks(range(len(unique_gb)))
    ax1.set_yticklabels(unique_gb)
    ax1.set_title("1. Bubble Chart (Match Strength)", fontsize=14, pad=10)
    ax1.set_xlabel(pA, fontweight='bold')
    ax1.set_ylabel(pB, fontweight='bold')
    ax1.grid(True, linestyle='--', alpha=0.3)

    # ==========================================
    # 2. Bipartite Network (Component Wiring)
    # ==========================================
    ax2 = axs[0, 1]
    # Draw nodes for Patient A
    for i, ga in enumerate(unique_ga):
        ax2.scatter(0, i, s=300, color='lightblue', zorder=5, edgecolors='black')
        ax2.text(-0.15, i, ga, ha='right', va='center', fontsize=12, fontweight='bold')
        
    # Draw nodes for Patient B
    for i, gb in enumerate(unique_gb):
        ax2.scatter(1, i, s=300, color='lightpink', zorder=5, edgecolors='black')
        ax2.text(1.15, i, gb, ha='left', va='center', fontsize=12, fontweight='bold')

    # Draw connecting edges (thickness correlates with similarity)
    for d in parsed_data:
        y1, y2 = map_ga[d['G_A']], map_gb[d['G_B']]
        color = 'blue' if 'Direct' in d['Type'] else 'red'
        line_width = (d['Sim'] - 70) / 5
        ax2.plot([0, 1], [y1, y2], color=color, linewidth=line_width, alpha=0.5, zorder=1)

    ax2.set_xlim(-0.5, 1.5)
    ax2.axis('off')
    ax2.set_title("2. Bipartite Network (Component Wiring)", fontsize=14, pad=10)

    # ==========================================
    # Aggregate protein data for plots 3 and 4
    # ==========================================
    prot_stats = {}
    for d in parsed_data:
        for p in d['Proteins']:
            p_name, pct, va, vb = p[0], float(p[1]), float(p[2]), float(p[3])
            if p_name not in prot_stats:
                prot_stats[p_name] = {'A': [], 'B': [], 'Match': []}
            prot_stats[p_name]['A'].append(va)
            prot_stats[p_name]['B'].append(vb)
            prot_stats[p_name]['Match'].append(pct)

    # Select the top 15 most frequently occurring proteins in these matches
    top_prots = sorted(prot_stats.keys(), key=lambda k: len(prot_stats[k]['A']), reverse=True)[:15]

    # ==========================================
    # 3. Dumbbell Plot (Avg Protein Expression Shift)
    # ==========================================
    ax3 = axs[1, 0]
    y_pos = np.arange(len(top_prots))
    
    for i, p in enumerate(top_prots):
        mean_a = np.mean(prot_stats[p]['A'])
        mean_b = np.mean(prot_stats[p]['B'])
        ax3.plot([mean_a, mean_b], [i, i], color='gray', zorder=1)
        ax3.scatter(mean_a, i, color='skyblue', s=100, zorder=2, edgecolors='black', label=pA if i==0 else "")
        ax3.scatter(mean_b, i, color='salmon', s=100, zorder=2, edgecolors='black', label=pB if i==0 else "")

    ax3.set_yticks(y_pos)
    ax3.set_yticklabels(top_prots)
    ax3.set_title("3. Dumbbell Plot (Avg Protein Expression Shift)", fontsize=14, pad=10)
    ax3.set_xlabel("Normalized Expression [-1.0 to 1.0]")
    ax3.axvline(0, color='black', linestyle='--', alpha=0.5)
    ax3.legend()
    ax3.grid(axis='x', linestyle='--', alpha=0.3)

    # ==========================================
    # 4. Heatmap (Protein Match % per G-Pair)
    # ==========================================
    ax4 = axs[1, 1]
    matrix = np.zeros((len(top_prots), len(parsed_data)))
    col_labels = []
    
    for j, d in enumerate(parsed_data):
        col_labels.append(f"{d['G_A']}-{d['G_B']}")
        prot_dict = {p[0]: float(p[1]) for p in d['Proteins']}
        for i, p in enumerate(top_prots):
            matrix[i, j] = prot_dict.get(p, 0.0)

    im = ax4.imshow(matrix, cmap='Greens', aspect='auto', vmin=0, vmax=100)
    
    ax4.set_xticks(np.arange(len(col_labels)))
    ax4.set_xticklabels(col_labels, rotation=45, ha='right')
    ax4.set_yticks(np.arange(len(top_prots)))
    ax4.set_yticklabels(top_prots)
    ax4.set_title("4. Heatmap (Protein Match % per G-Pair)", fontsize=14, pad=10)
    
    cbar = fig.colorbar(im, ax=ax4, fraction=0.046, pad=0.04)
    cbar.set_label('Match Percentage (%)')

    fig.tight_layout(pad=3.0)
    plt.show()

def interactive_menu():
    """
    Main CLI function allowing the user to search for specific patient pairs.
    """
    df = load_and_parse_data()
    if df is None: return
    
    # Collect all unique patients for reference
    all_patients = sorted(list(set(df['Patient A'].dropna()) | set(df['Patient B'].dropna())))
    
    print("=" * 60)
    print(" 🧬 TME Maximal Sequences - Interactive Dashboard 🧬")
    print("=" * 60)
    print(f"Loaded {len(df)} total matches involving {len(all_patients)} unique ROIs.")
    print("Example format: P13_ROI_71")
    print("-" * 60)
    
    while True:
        pA = input("\nEnter Patient A (or 'q' to quit, 'list' to see all): ").strip()
        if pA.lower() == 'q':
            print("Exiting...")
            break
        if pA.lower() == 'list':
            print("\nAvailable Patients/ROIs:")
            print(", ".join(all_patients))
            continue
            
        pB = input("Enter Patient B: ").strip()
        
        # Look for the exact pair in either order
        match_row = df[((df['Patient A'] == pA) & (df['Patient B'] == pB)) | 
                       ((df['Patient A'] == pB) & (df['Patient B'] == pA))]
        
        if not match_row.empty:
            print(f"Match found! Generating dashboard for {pA} vs {pB}...")
            details = match_row.iloc[0][[col for col in df.columns if 'Match Details' in col][0]]
            parsed = extract_all_matches(details)
            
            if parsed:
                # Execution halts here until the user closes the matplotlib window
                build_dashboard(pA, pB, parsed)
            else:
                print("Error: Could not parse the proteins for this match.")
        else:
            print(f"No match found between '{pA}' and '{pB}' in the report.")
            print("Tip: Check for typos or type 'list' to see all available names.")

if __name__ == "__main__":
    interactive_menu()