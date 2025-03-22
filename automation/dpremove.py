import pandas as pd
import os

# Path to your file
file_path = r"U:\excel automation\combined_data.xlsx"

# Read the Excel file
df = pd.read_excel(file_path)

# Remove exact duplicates
cleaned_df = df.drop_duplicates()

# Save the cleaned data
output_path = os.path.join(r"U:\excel automation", "cleaned_data.xlsx")
cleaned_df.to_excel(output_path, index=False)

print(f"Duplicates removed. Cleaned data saved to: {output_path}")
print(f"Total rows before: {len(df)} | Total rows after: {len(cleaned_df)}")
