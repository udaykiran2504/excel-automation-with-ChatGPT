import pandas as pd
import os
import subprocess

# Path to your Excel files
folder_path = r"U:\excel automation"
file_names = [f"excel_data_month_{i}.xlsx" for i in range(1, 13)]

# Read and combine Excel files
dataframes = [pd.read_excel(os.path.join(folder_path, file)) for file in file_names]
combined_df = pd.concat(dataframes, ignore_index=True)

# Calculate Total Sales
total_sales = combined_df['Total Sales'].sum()
print(f"Total Sales: {total_sales}")

# Create a Total Sales row with empty values except in the 'Total Sales' column
total_sales_row = {col: '' for col in combined_df.columns}
total_sales_row['Total Sales'] = total_sales  # Optional if you have a Month column

# Append the row to the DataFrame
combined_df = pd.concat([combined_df, pd.DataFrame([total_sales_row])], ignore_index=True)

# Path for the output file
output_path = os.path.join(folder_path, "combined_data.xlsx")

# Check if the file exists and rename if necessary
if os.path.exists(output_path):
    print("File is in use or cannot be overwritten. Saving with a new name.")
    output_path = os.path.join(folder_path, "combined_data_new.xlsx")

# Save the DataFrame
combined_df.to_excel(output_path, index=False)
print(f"Data saved to: {output_path}")

# Open the Excel file
subprocess.run(["start", "", output_path], shell=True)

