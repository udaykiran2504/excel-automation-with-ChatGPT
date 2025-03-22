import pandas as pd
import os
import subprocess

# Path to your Excel files
folder_path = r"U:\excel automation"
file_names = [f"excel_data_month_{i}.xlsx" for i in range(1, 13)]

# Read and concatenate all files
dataframes = [pd.read_excel(os.path.join(folder_path, file)) for file in file_names]
combined_df = pd.concat(dataframes, ignore_index=True)

# Save the combined data to a single Excel file
output_path = os.path.join(folder_path, "combined_data.xlsx")
combined_df.to_excel(output_path, index=False)

print(f"Files concatenated successfully! Combined data saved at: {output_path}")

# Open the Excel file
subprocess.run(["start", "", output_path], shell=True)
