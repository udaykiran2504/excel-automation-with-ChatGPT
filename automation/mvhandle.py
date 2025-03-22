import pandas as pd

# Path to the cleaned file
file_path = r"U:\excel automation\cleaned_data.xlsx"

# Read the Excel file
df = pd.read_excel(file_path)

# Fill missing values in the 'Region' column with Mode
mode_region = df['Region'].mode()[0]
df.loc[:, 'Region'] = df['Region'].fillna(mode_region)

print(f"Missing values filled with Mode: {mode_region}")

# Save the updated data
output_path = r"U:\excel automation\cleaned_data_filled.xlsx"
df.to_excel(output_path, index=False)
print(f"Data saved to: {output_path}")
