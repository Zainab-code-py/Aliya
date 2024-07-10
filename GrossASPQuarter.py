import pandas as pd
from datetime import datetime, timedelta

# Load the data
gross_sales_path = 'C:/Users/Aliya/Downloads/quarterly_gross_sales.csv'
gross_quantity_path = 'C:/Users/Aliya/Downloads/quarterly_gross_quantity.csv'
bcodes_path = 'C:/Users/Aliya/Downloads/BCodes.csv'

gross_sales = pd.read_csv(gross_sales_path)
gross_quantity = pd.read_csv(gross_quantity_path)
bcodes = pd.read_csv(bcodes_path)

# Define the date range for the fiscal quarter (2nd April to 30th June) for this year and last year
specified_start = datetime(2024, 4, 2)
specified_end = datetime(2024, 6, 30)
specified_start_ly = specified_start - timedelta(days=365)
specified_end_ly = specified_end - timedelta(days=365)

# Merge the gross sales and gross quantity data
merged_data = gross_sales.merge(gross_quantity, on='Brand', suffixes=('_Sales', '_Quantity'))

# Calculate the gross ASP for TY and LY
merged_data['TY ASP'] = merged_data['TY'] / merged_data['TY Quantity']
merged_data['LY ASP'] = merged_data['LY'] / merged_data['LY Quantity']

# Calculate vs LY as percentage
merged_data['vs LY'] = ((merged_data['TY ASP'] - merged_data['LY ASP']) / merged_data['LY ASP']) * 100

# Select relevant columns
result = merged_data[['Brand', 'TY ASP', 'LY ASP', 'vs LY']]

# Replace brand codes with brand names
result = result.merge(bcodes, left_on='Brand', right_on='Code', how='left')
result = result.drop(columns=['Brand', 'Code'])
result = result.rename(columns={'Name': 'Brand'})

# Save the result to a CSV file and Excel file
csv_output_path = 'C:/Users/Aliya/Downloads/quarterly_gross_asp.csv'
excel_output_path = 'C:/Users/Aliya/Downloads/quarterly_gross_asp.xlsx'
result.to_csv(csv_output_path, index=False)
result.to_excel(excel_output_path, index=False)

print("Gross ASP data has been successfully saved.")
