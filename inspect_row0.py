import pandas as pd
import openpyxl

excel_file = "I SEM DIGIVAL PURPOSE-25 AB F+RV.xlsx"
wb = openpyxl.load_workbook(excel_file, data_only=True)

for sheet in wb.sheetnames:
    df = pd.read_excel(excel_file, sheet_name=sheet)
    if df.empty or len(df.columns) < 3:
        continue
    
    print(f"\n--- SHEET: {sheet} ---")
    row0 = df.iloc[0]
    print("Row 0 (Credits row):")
    for col in df.columns:
        val0 = row0[col]
        if pd.notna(val0):
            print(f"  {col[:45]:45s} -> Row0: {val0}")
