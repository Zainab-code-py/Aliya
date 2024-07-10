import pandas as pd

# Load the CSV files
net_sales_path = 'C:/Users/Aliya/Downloads/yearly_net_sales.csv'
gross_sales_path = 'C:/Users/Aliya/Downloads/yearly_gross_sales.csv'
bcodes_path = 'C:/Users/Aliya/Downloads/BCodes.csv'

# Load the data
net_sales = pd.read_csv(net_sales_path, dtype={'Brand': str, 'TY': float, 'LY': float}, low_memory=False)
gross_sales = pd.read_csv(gross_sales_path, dtype={'Brand': str, 'TY': float, 'LY': float}, low_memory=False)
bcodes = pd.read_csv(bcodes_path, dtype={'Code': str, 'Name': str}, low_memory=False)

# Summarize the net sales and gross sales by brand for TY and LY
net_sales_summary = net_sales[['Brand', 'TY', 'LY']].copy()
gross_sales_summary = gross_sales[['Brand', 'TY', 'LY']].copy()

# Merge the net sales and gross sales data
merged_data = net_sales_summary.merge(gross_sales_summary, on='Brand', suffixes=('_NET', '_GROSS'))

# Calculate the Gross to Net percentage for both TY and LY
merged_data['Gross_to_Net%_TY'] = merged_data['TY_NET'] / merged_data['TY_GROSS']
merged_data['Gross_to_Net%_LY'] = merged_data['LY_NET'] / merged_data['LY_GROSS']

# Replace brand codes with brand names
result = merged_data.merge(bcodes, left_on='Brand', right_on='Name', how='left')

# Select relevant columns
result = result[['Brand', 'TY_NET', 'TY_GROSS', 'Gross_to_Net%_TY', 'LY_NET', 'LY_GROSS', 'Gross_to_Net%_LY']]

# Save the result to a CSV file and Excel file
result.to_csv('C:/Users/Aliya/Downloads/gross_to_net_percentage-YTD.csv', index=False)
result.to_excel('C:/Users/Aliya/Downloads/gross_to_net_percentage-YTD.xlsx', index=False)

print("Gross to Net percentage data for YTD has been successfully saved.")
