import pandas as pd

# Path to the cleaned file
file_path = r"U:\excel automation\cleaned_data_filled.xlsx"

# Read the Excel file
df = pd.read_excel(file_path)

# Check data types before conversion
print("Data Types Before Conversion:")
print(df.dtypes)

# Convert 'Order Date' to DateTime
df['Order Date'] = pd.to_datetime(df['Order Date'], errors='coerce')

# Convert numeric columns to float (handling any errors)
numeric_cols = ['Quantity', 'Unit Price', 'Total Sales']
for col in numeric_cols:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# Convert categorical columns to string
string_cols = ['Customer Name', 'Product', 'Category', 'Region']
for col in string_cols:
    df[col] = df[col].astype(str)

# Check data types after conversion
print("\nData Types After Conversion:")
print(df.dtypes)

# Save the cleaned data
output_path = r"U:\excel automation\cleaned_data_final.xlsx"
df.to_excel(output_path, index=False)
print(f"\nData types fixed and saved to: {output_path}")
