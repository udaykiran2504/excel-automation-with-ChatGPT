import pandas as pd

# Path to the cleaned file
file_path = r"U:\excel automation\cleaned_data.xlsx"

# Read the Excel file
df = pd.read_excel(file_path)

# Check for missing values
missing_values = df.isnull().sum()

print("Missing Values in Each Column:")
print(missing_values)

# Total number of missing values
total_missing = missing_values.sum()
print(f"\nTotal Missing Values: {total_missing}")
