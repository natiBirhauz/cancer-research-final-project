"""
create_sample_template.py
-------------------------
Generates an official sample Excel template (Sample_Input_Template.xlsx)
demonstrating the exact format expected by the Cancer Research TME Algorithm.
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

TEMPLATE_PATH = os.path.join('Docs', 'Sample_Input_Template.xlsx')

def generate_template():
    os.makedirs('Docs', exist_ok=True)
    wb = openpyxl.Workbook()

    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )
    header_fill = PatternFill(start_color='2C3E50', end_color='2C3E50', fill_type='solid')
    header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    title_font = Font(name='Calibri', size=14, bold=True, color='000000')
    subtitle_font = Font(name='Calibri', size=11, italic=True, color='555555')
    bold_font = Font(name='Calibri', size=11, bold=True, color='000000')
    regular_font = Font(name='Calibri', size=11, color='000000')

    # -------------------------------------------------------------
    # Sheet 1: Instructions & Guidelines
    # -------------------------------------------------------------
    ws_readme = wb.active
    ws_readme.title = "Instructions"
    ws_readme.views.sheetView[0].showGridLines = True

    ws_readme.append(["CANCER RESEARCH TME ALGORITHM - INPUT DATA TEMPLATE"])
    ws_readme.append(["Standard Format Specification for Surprisal Analysis Spreadsheets"])
    ws_readme.append([])
    ws_readme.append(["Section", "Description & Rules"])

    guidelines = [
        ("Workbook Structure", "The workbook must contain two sheets: 'Gvalues' and 'Lambavalues'."),
        ("Sheet 'Gvalues'", "Represents the protein process matrix (G-vectors)."),
        (" - Protein Markers", "One column must contain protein names (e.g. named 'MARKER' or placed in column A)."),
        (" - Process Vectors", "Numerical columns must be named G0, G1, G2, ..., Gk (G0 is baseline equilibrium, G1..Gk are deviation processes)."),
        ("Sheet 'Lambavalues'", "Represents the single-cell amplitude coordinates (L-values)."),
        (" - Cellular Columns", "Columns must be named L0, L1, L2, ..., Lk corresponding directly to G0, G1, ..., Gk."),
        (" - Cells / Rows", "Each row corresponds to an individual segmented cell in the sample (e.g., hundreds or thousands of cells)."),
        ("Single / Multi-sample", "You can upload a single Excel file for one sample, multiple Excel files at once, a ZIP archive containing sample folders, or specify a directory path."),
        ("Tail Format (Alternative)", "If using tail tables, columns must be named G1+, G1-, G2+, G2-, etc., listing over- and under-expressed proteins.")
    ]

    for sec, desc in guidelines:
        ws_readme.append([sec, desc])

    # -------------------------------------------------------------
    # Sheet 2: Gvalues (Sample Matrix)
    # -------------------------------------------------------------
    ws_g = wb.create_sheet(title="Gvalues")
    ws_g.views.sheetView[0].showGridLines = True

    # Header row
    g_headers = ["MARKER", "G0", "G1", "G2", "G3", "G4", "G5", "G6", "G7", "G8", "G9"]
    ws_g.append(g_headers)

    sample_proteins = [
        ("CMET", 0.2826,  0.2048,  0.1519, -0.1982,  0.0712,  0.3120, -0.1354, -0.6965,  0.0972,  0.1205),
        ("HER2", 0.3150, -0.1820,  0.2240,  0.0950, -0.1420,  0.0890,  0.2150,  0.1450, -0.2100, -0.0540),
        ("CD44", 0.2980,  0.3120, -0.1850,  0.1400,  0.2250, -0.0950, -0.1650,  0.0850,  0.1750,  0.2300),
        ("CKIT", 0.3050,  0.0950,  0.3200, -0.2100, -0.1850,  0.1750,  0.0950, -0.1200,  0.1450, -0.1800),
        ("CD24", 0.2890, -0.2250,  0.1150,  0.3100, -0.0950, -0.2150,  0.1850,  0.0950, -0.1650,  0.1100),
        ("EPCAM",0.3400, -0.1450, -0.2100, -0.0950,  0.3100,  0.1450, -0.0850,  0.2200, -0.0950, -0.1450),
        ("MUC1", 0.3250,  0.1850, -0.0950,  0.2450, -0.1650, -0.1100,  0.2450, -0.0950,  0.1850,  0.0950),
        ("EGFR", 0.3100,  0.2450,  0.1650, -0.1200,  0.1950, -0.1850, -0.1450,  0.1850, -0.1100,  0.1650),
        ("IGFR", 0.2950,  0.1650,  0.2850, -0.1750, -0.1100,  0.2300, -0.1950, -0.1450,  0.0950, -0.1200),
        ("PDL1", 0.2750, -0.2950, -0.1450,  0.1850,  0.1650, -0.1450,  0.1100,  0.0750,  0.2450, -0.1950)
    ]

    for row in sample_proteins:
        ws_g.append(list(row))

    # -------------------------------------------------------------
    # Sheet 3: Lambavalues (Sample Cells)
    # -------------------------------------------------------------
    ws_l = wb.create_sheet(title="Lambavalues")
    ws_l.views.sheetView[0].showGridLines = True

    l_headers = ["Cell_ID", "L0", "L1", "L2", "L3", "L4", "L5", "L6", "L7", "L8", "L9"]
    ws_l.append(l_headers)

    import numpy as np
    np.random.seed(42)
    for c_idx in range(1, 26):
        c_name = f"Cell_{c_idx:04d}"
        l0_val = round(float(np.random.normal(5.5, 0.4)), 4)
        l_deviations = [round(float(np.random.normal(0.0, 0.8)), 4) for _ in range(9)]
        ws_l.append([c_name, l0_val] + l_deviations)

    # -------------------------------------------------------------
    # Styling All Sheets (Clean Black & White / Corporate)
    # -------------------------------------------------------------
    for ws in [ws_readme, ws_g, ws_l]:
        ws.cell(1, 1).font = title_font
        if ws == ws_readme:
            ws.cell(2, 1).font = subtitle_font
            for col in range(1, 3):
                c = ws.cell(4, col)
                c.fill = PatternFill(start_color='E5E7EB', end_color='E5E7EB', fill_type='solid')
                c.font = bold_font
            for row in range(5, ws.max_row + 1):
                ws.cell(row, 1).font = bold_font
                ws.cell(row, 2).font = regular_font
        else:
            for col in range(1, ws.max_column + 1):
                c = ws.cell(1, col)
                c.fill = PatternFill(start_color='E5E7EB', end_color='E5E7EB', fill_type='solid')
                c.font = bold_font
                c.alignment = Alignment(horizontal='center', vertical='center')

            for row in range(2, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row, col)
                    cell.font = regular_font
                    cell.border = thin_border
                    if col > 1 and ws != ws_readme:
                        cell.alignment = Alignment(horizontal='right')

        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    wb.save(TEMPLATE_PATH)
    # Also save in root directory for easy access
    root_template = "Sample_Input_Template.xlsx"
    wb.save(root_template)
    print(f"Sample template created: {TEMPLATE_PATH} and {root_template}")

if __name__ == '__main__':
    generate_template()

