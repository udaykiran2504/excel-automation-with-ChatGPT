import pandas as pd
import os

# Path to your file
file_path = r"U:\excel automation\combined_data.xlsx"

# Read the Excel file
df = pd.read_excel(file_path)

# Check for duplicates
duplicates = df[df.duplicated('Customer Name')]

if not duplicates.empty:
    print(f"Found {len(duplicates)} duplicate rows.")
    print(duplicates)
else:
    print("No duplicates found.")

