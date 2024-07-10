from datetime import timedelta, datetime

import pandas as pd

# Load the provided CSV files
yearly_gross_transactions = pd.read_csv('C:/Users/Aliya/Downloads/yearly_gross_transactions.csv', dtype={'Brand': str}, low_memory=False)
yearly_gross_quantity = pd.read_csv('C:/Users/Aliya/Downloads/yearly_gross_quantity.csv', dtype={'Brand': str}, low_memory=False)

# Define the date range for 2nd April to 30th June for this year and last year
specified_start = datetime(2024, 4, 2)
specified_end = datetime(2024, 6, 30)
specified_start_ly = specified_start - timedelta(days=365)
specified_end_ly = specified_end - timedelta(days=365)
# Merging the datasets and calculating the required fields
merged_quantity_transactions = yearly_gross_quantity.merge(yearly_gross_transactions, on='Brand', suffixes=('_Quantity', '_Transactions'))

# Calculating the gross UPT for TY and LY
merged_quantity_transactions['TY UPT'] = merged_quantity_transactions['TY Quantity'] / merged_quantity_transactions['TY']
merged_quantity_transactions['LY UPT'] = merged_quantity_transactions['LY Quantity'] / merged_quantity_transactions['LY']

# Calculating vs LY as percentage
merged_quantity_transactions['vs LY'] = ((merged_quantity_transactions['TY UPT'] - merged_quantity_transactions['LY UPT']) / merged_quantity_transactions['LY UPT']) * 100

# Selecting relevant columns
result = merged_quantity_transactions[['Brand', 'TY UPT', 'LY UPT', 'vs LY']]

# Saving the result to a CSV file and Excel file
csv_output_path = 'C:/Users/Aliya/Downloads/yearly_gross_upt.csv'
excel_output_path = 'C:/Users/Aliya/Downloads/yearly_gross_upt.xlsx'
result.to_csv(csv_output_path, index=False)
result.to_excel(excel_output_path, index=False)

# Display the result
print("Gross UPT data has been successfully saved.")
print(result)
