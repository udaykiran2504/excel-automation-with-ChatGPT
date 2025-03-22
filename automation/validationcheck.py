import pandas as pd

# Path to the cleaned file
file_path = r"U:\excel automation\cleaned_data_final.xlsx"

# Read the Excel file
df = pd.read_excel(file_path)

# Track issues
issues = []

### 1. Check for Invalid Dates
if 'Order Date' in df.columns:
    invalid_dates = df[(df['Order Date'].isnull()) | (df['Order Date'] < '2024-01-01') | (df['Order Date'] > '2024-12-31')]
    if not invalid_dates.empty:
        issues.append(f"Invalid Dates Found:\n{invalid_dates}\n")

### 2. Check for Negative or Zero Values
for col in ['Quantity', 'Unit Price', 'Total Sales']:
    if col in df.columns:
        invalid_values = df[df[col] <= 0]
        if not invalid_values.empty:
            issues.append(f"Invalid {col} Values Found:\n{invalid_values}\n")

### 3. Validate Categories and Regions
valid_categories = ['Electronics', 'Furniture', 'Clothing', 'Accessories', 'Home Appliances']
valid_regions = ['North', 'South', 'East', 'West', 'Central']

if 'Category' in df.columns:
    invalid_categories = df[~df['Category'].isin(valid_categories)]
    if not invalid_categories.empty:
        issues.append(f"Invalid Categories Found:\n{invalid_categories}\n")

if 'Region' in df.columns:
    invalid_regions = df[~df['Region'].isin(valid_regions)]
    if not invalid_regions.empty:
        issues.append(f"Invalid Regions Found:\n{invalid_regions}\n")

### Display Results
if issues:
    print("Validation Errors Found:")
    for issue in issues:
        print(issue)
else:
    print("No Validation Errors Found. Data is Clean!")
