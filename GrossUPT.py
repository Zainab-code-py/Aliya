import pandas as pd
from datetime import datetime, timedelta

# Load the data with specified dtypes to avoid mixed type warning
quarterly_gross_transactions = pd.read_csv('C:/Users/Aliya//Downloads/quarterly_gross_transactions.csv', dtype={'Brand': str}, low_memory=False)
quarterly_gross_quantity = pd.read_csv('C:/Users/Aliya/Downloads/quarterly_gross_quantity.csv', dtype={'Brand': str}, low_memory=False)
bcodes = pd.read_csv('C:/Users/Aliya/Downloads/BCodes.csv')

# Merge the quantity and transactions data
merged_quantity_transactions = quarterly_gross_quantity.merge(quarterly_gross_transactions, on='Brand', suffixes=('_Quantity', '_Transactions'))

# Calculate the gross UPT for TY and LY
merged_quantity_transactions['TY UPT'] = merged_quantity_transactions['TY Quantity'] / merged_quantity_transactions['TY ']
merged_quantity_transactions['LY UPT'] = merged_quantity_transactions['LY Quantity'] / merged_quantity_transactions['LY ']

# Calculate vs LY as percentage
merged_quantity_transactions['vs LY'] = ((merged_quantity_transactions['TY UPT'] - merged_quantity_transactions['LY UPT']) / merged_quantity_transactions['LY UPT']) * 100

# Replace brand codes with brand names
result = merged_quantity_transactions.merge(bcodes, left_on='Brand', right_on='Code', how='left')
result = result.drop(columns=[ 'Code'])
#result = result.rename(columns={'Name': 'Brand'})

# Select relevant columns
result = result[['Brand', 'TY UPT', 'LY UPT', 'vs LY']]

# Save the result to a CSV file and Excel file
csv_output_path = 'C:/Users/Aliya/Downloads/quarterly_gross_upt.csv'
excel_output_path = 'C:/Users/Aliya/Downloads/quarterly_gross_upt.xlsx'
result.to_csv(csv_output_path, index=False)
result.to_excel(excel_output_path, index=False)

print("Gross UPT data has been successfully saved.")
