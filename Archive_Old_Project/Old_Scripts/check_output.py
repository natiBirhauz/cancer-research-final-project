import openpyxl
import os

xlsx = r'Docs\Optimal_ES_Result_20260915_134232.xlsx'
wb = openpyxl.load_workbook(xlsx)

print(f"File: {xlsx}")
print(f"Size: {os.path.getsize(xlsx):,} bytes")
print(f"\nSheets ({len(wb.sheetnames)}):")
for name in wb.sheetnames:
    ws = wb[name]
    print(f"  - {name:<30}  ({ws.max_row} rows x {ws.max_column} cols)")

print("\n" + "="*70)
print("Summary sheet preview:")
print("="*70)
ws = wb['Summary']
for row in ws.iter_rows(min_row=1, max_row=2, values_only=True):
    print(row)

print("\n" + "="*70)
print("Per_ROI_Coverage sample (first 5 ROIs):")
print("="*70)
ws = wb['Per_ROI_Coverage']
for i, row in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True)):
    print(row)

print("\n" + "="*70)
print("Coverage_Curve sample (first 10 + last 5 ES):")
print("="*70)
ws = wb['Coverage_Curve']
rows = list(ws.iter_rows(min_row=1, max_row=11, values_only=True))
for row in rows:
    print(row)
print("...")
for row in ws.iter_rows(min_row=ws.max_row-4, max_row=ws.max_row, values_only=True):
    print(row)
