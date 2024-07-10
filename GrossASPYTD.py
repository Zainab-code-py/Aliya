import pandas as pd
from datetime import datetime, timedelta

# Load the data
gross_sales_path = 'C:/Users/Aliya/Downloads/yearly_gross_sales.csv'
gross_quantity_path = 'C:/Users/Aliya/Downloads/yearly_gross_quantity.csv'

gross_sales = pd.read_csv(gross_sales_path)
gross_quantity = pd.read_csv(gross_quantity_path)

# Define the date range for 2nd April to 30th June for this year and last year
specified_start = datetime(2024, 4, 2)
specified_end = datetime(2024, 6, 30)
specified_start_ly = specified_start - timedelta(days=365)
specified_end_ly = specified_end - timedelta(days=365)

# Ensure the brand columns are strings
gross_sales['Brand'] = gross_sales['Brand'].astype(str)
gross_quantity['Brand'] = gross_quantity['Brand'].astype(str)

# Strip any leading/trailing spaces from column names
gross_sales.columns = gross_sales.columns.str.strip()
gross_quantity.columns = gross_quantity.columns.str.strip()

# Find common brands in both sheets
common_brands = set(gross_sales['Brand']).intersection(set(gross_quantity['Brand']))

# Filter the sheets to only include common brands
gross_sales_filtered = gross_sales[gross_sales['Brand'].isin(common_brands)].set_index('Brand')
gross_quantity_filtered = gross_quantity[gross_quantity['Brand'].isin(common_brands)].set_index('Brand')

# Ensure 'TY' and 'LY' columns exist in both dataframes
if 'TY' not in gross_sales_filtered.columns or 'LY' not in gross_sales_filtered.columns:
    raise KeyError("Columns 'TY' or 'LY' are missing in gross_sales_filtered")
if 'TY Quantity' not in gross_quantity_filtered.columns or 'LY Quantity' not in gross_quantity_filtered.columns:
    raise KeyError("Columns 'TY Quantity' or 'LY Quantity' are missing in gross_quantity_filtered")

# Calculate ASP for TY and LY
gross_asp = pd.DataFrame()
gross_asp['TY ASP'] = gross_sales_filtered['TY'] / gross_quantity_filtered['TY Quantity']
gross_asp['LY ASP'] = gross_sales_filtered['LY'] / gross_quantity_filtered['LY Quantity']

# Replace NaN values resulting from division by zero with 0
gross_asp = gross_asp.fillna(0)

# Calculate vs LY
gross_asp['vs LY'] = ((gross_asp['TY ASP'] - gross_asp['LY ASP']) / gross_asp['LY ASP']) * 100

# Round vs LY to 2 decimal places and add %
gross_asp['vs LY'] = gross_asp['vs LY'].round(2).astype(str) + '%'

# Reset the index to make Brand a column
gross_asp = gross_asp.reset_index()

# Reorder columns to place Brand first
gross_asp = gross_asp[['Brand', 'TY ASP', 'LY ASP', 'vs LY']]

# Save the result to a CSV file and Excel file
csv_output_path = 'C:/Users/Aliya/Downloads/yearly_gross_asp.csv'
excel_output_path = 'C:/Users/Aliya/Downloads/yearly_gross_asp.xlsx'
gross_asp.to_csv(csv_output_path, index=False)
gross_asp.to_excel(excel_output_path, index=False)

# Display the result
print("Gross ASP data has been successfully saved.")
print(gross_asp)
