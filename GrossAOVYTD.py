import pandas as pd
from datetime import datetime, timedelta

# Load the data
gross_sales = pd.read_csv('C:/Users/Aliya/Downloads/yearly_gross_sales.csv')
gross_transactions = pd.read_csv('C:/Users/Aliya/Downloads/yearly_gross_transactions.csv')

# Print the columns to debug and inspect the structure
print("Gross Sales Columns: ", gross_sales.columns)
print("Gross Transactions Columns: ", gross_transactions.columns)

# Define the date range for 2nd April to 30th June for this year and last year
specified_start = datetime(2024, 4, 2)
specified_end = datetime(2024, 6, 30)
specified_start_ly = specified_start - timedelta(days=365)
specified_end_ly = specified_end - timedelta(days=365)
# Ensure the brand columns are strings
gross_sales['Brand'] = gross_sales['Brand'].astype(str)
gross_transactions['Brand'] = gross_transactions['Brand'].astype(str)

# Strip any leading/trailing spaces from column names
gross_sales.columns = gross_sales.columns.str.strip()
gross_transactions.columns = gross_transactions.columns.str.strip()

# Find common brands in both sheets
common_brands = set(gross_sales['Brand']).intersection(set(gross_transactions['Brand']))

# Filter the sheets to only include common brands
gross_sales_filtered = gross_sales[gross_sales['Brand'].isin(common_brands)].set_index('Brand')
gross_transactions_filtered = gross_transactions[gross_transactions['Brand'].isin(common_brands)].set_index('Brand')

# Ensure 'TY' and 'LY' columns exist in both dataframes
if 'TY' not in gross_sales_filtered.columns or 'LY' not in gross_sales_filtered.columns:
    raise KeyError("Columns 'TY' or 'LY' are missing in gross_sales_filtered")
if 'TY' not in gross_transactions_filtered.columns or 'LY' not in gross_transactions_filtered.columns:
    raise KeyError("Columns 'TY' or 'LY' are missing in gross_transactions_filtered")

# Calculate AOV for TY and LY
gross_aov = pd.DataFrame()
gross_aov['TY'] = gross_sales_filtered['TY'] / gross_transactions_filtered['TY']
gross_aov['LY'] = gross_sales_filtered['LY'] / gross_transactions_filtered['LY']

# Replace NaN values resulting from division by zero with 0
gross_aov = gross_aov.fillna(0)

# Calculate vs LY
gross_aov['vs LY'] = ((gross_aov['TY'] - gross_aov['LY']) / gross_aov['LY']) * 100

# Round vs LY to 2 decimal places and add %
gross_aov['vs LY'] = gross_aov['vs LY'].round(2).astype(str) + '%'

# Reset the index to make Brand a column
gross_aov = gross_aov.reset_index()

# Reorder columns to place Brand first
gross_aov = gross_aov[['Brand', 'TY', 'LY', 'vs LY']]

# Save the result to a CSV file
gross_aov.to_csv('C:/Users/Aliya/Downloads/GrossAOV-YTD.csv', index=False)

# Display the result
print("Gross AOV data has been successfully saved.")
print(gross_aov)
