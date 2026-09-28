import pandas as pd
import openpyxl

excel_file = "I SEM DIGIVAL PURPOSE-25 AB F+RV.xlsx"
wb = openpyxl.load_workbook(excel_file, data_only=True)

for sheet in wb.sheetnames:
    df = pd.read_excel(excel_file, sheet_name=sheet)
    if df.empty or len(df.columns) < 3:
        print(f"Sheet '{sheet}' is empty or header only.")
        continue
    
    print(f"\n==========================================")
    print(f"SHEET: {sheet}")
    print(f"==========================================")
    print("Columns:", list(df.columns))
    
    # Find student rows (rows with valid HALL TICKET or SUC NUMBER)
    hall_ticket_col = None
    for col in df.columns:
        if "HALL" in str(col).upper() or "TICKET" in str(col).upper():
            hall_ticket_col = col
            break
    
    if hall_ticket_col:
        valid_rows = df[df[hall_ticket_col].notna()]
        print(f"Valid student rows in '{sheet}': {len(valid_rows)}")
        if not valid_rows.empty:
            sample = valid_rows.iloc[0]
            print("\nSample Student Record:")
            for c, val in sample.items():
                print(f"  {c:55s}: {val}")
