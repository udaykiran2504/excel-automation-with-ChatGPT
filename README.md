# Excel Automation for Data Analysis

This project automates various data analysis tasks using Python and Excel. It covers end-to-end data operations including data cleaning, validation, transformation, and aggregation. The purpose of this automation is to streamline repetitive tasks and enhance efficiency in business analysis.

## Features

- **Data Consolidation:** Combines multiple Excel files (monthly sales data) into a single file.
- **Data Cleaning:** Identifies and handles missing values and duplicate records.
- **Data Validation:** Checks for correct data types and validates values.
- **Data Transformation:** Extracts insights using calculated columns (e.g., Profit, Cost).
- **Conditional Logic:** Classifies data using business rules.
- **Profit Calculation:** Applies a fixed 25% profit margin to determine profit from sales.

## Prerequisites

Ensure you have the following installed:

- Python 3.11+
- Pandas
- OpenPyXL
- Git
- Excel or any compatible spreadsheet software

### Install Required Packages

```bash
pip install pandas openpyxl
```

## Folder Structure

```
excel-automation/
|-- automation/
|   |-- data/                 # Excel files for each month
|   |-- cleaned_data_final.xlsx
|   |-- transformed_data.xlsx
|   |-- concat.py             # Combines monthly data
|   |-- sum.py                # Calculates Total Sales and Profit
|   |-- mvhandle.py           # Handles missing values
|   |-- validation.py         # Performs data validation
|-- README.md
```

## How to Use

1. **Clone the Repository:**

```bash
git clone <repository_url>
```

2. **Navigate to the Directory:**

```bash
cd excel-automation/automation
```

3. **Combine Monthly Data:**

```bash
python concat.py
```

4. **Check and Calculate Total Sales and Profit:**

```bash
python sum.py
```

5. **Handle Missing Values:**

```bash
python mvhandle.py
```

6. **Perform Data Validation:**

```bash
python validation.py
```

7. **View Results:**

- Check `transformed_data.xlsx` for consolidated results.
- Check `cleaned_data_final.xlsx` for the cleaned data.

## Contribution

Feel free to contribute by submitting issues or pull requests. Ensure code is clean and well-commented.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

its free to use 
