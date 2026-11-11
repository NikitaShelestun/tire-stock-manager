import pandas as pd
import os

# Path to your Excel file
excel_path = r'.\tires_data.xlsx'

if not os.path.exists(excel_path):
    print(f"File not found at: {excel_path}")
    print("Please check the path is correct.")
else:
    df = pd.read_excel(excel_path)
    
    numeric_cols = ['ID', 'WIDTH', 'PROFILE', 'DIAMETER', 'FLOOR', 'QUANTITY']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)
    
    df.to_excel(excel_path, index=False)
    print("SUCCESS! File fixed.")

input("Press Enter to exit...")
