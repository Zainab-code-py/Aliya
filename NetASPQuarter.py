import pandas as pd

def safe_divide(numerator, denominator):
    return numerator / denominator if denominator != 0 and numerator != 0 else 0

# Load the data
net_sales = pd.read_csv('C:/Users/Aliya/Downloads/quarterly_net_sales.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
net_transactions = pd.read_csv('C:/Users/Aliya/Downloads/quarterly_net_quantity.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
bcodes = pd.read_csv('C:/Users/Aliya/Downloads/BCodes.csv')

# Ensure the brand columns are strings
net_sales['Brand'] = net_sales['Brand'].astype(str)
net_transactions['Brand'] = net_transactions['Brand'].astype(str)

# Strip any leading/trailing spaces from column names
net_sales.columns = net_sales.columns.str.strip()
net_transactions.columns = net_transactions.columns.str.strip()

# Find common brands in both datasets
common_brands = set(net_sales['Brand']).intersection(set(net_transactions['Brand']))

# Filter the data to include only common brands
net_sales = net_sales[net_sales['Brand'].isin(common_brands)]
net_transactions = net_transactions[net_transactions['Brand'].isin(common_brands)]

# Group and sum the net sales by brand and count unique orders for TY
net_sales_specified = net_sales.groupby('Brand').agg({'TY': 'sum'}).reset_index()
net_sales_specified.columns = ['Brand', 'TY_Net_Sales']

net_transactions_specified = net_transactions.groupby('Brand').agg({'TY': 'sum'}).reset_index()
net_transactions_specified.columns = ['Brand', 'TY_Orders']

# Group and sum the net sales by brand and count unique orders for LY
net_sales_specified_ly = net_sales.groupby('Brand').agg({'LY': 'sum'}).reset_index()
net_sales_specified_ly.columns = ['Brand', 'LY_Net_Sales']

net_transactions_specified_ly = net_transactions.groupby('Brand').agg({'LY': 'sum'}).reset_index()
net_transactions_specified_ly.columns = ['Brand', 'LY_Orders']

# Merge the TY and LY data for net sales and transactions
merged_sales = net_sales_specified.merge(net_sales_specified_ly, on='Brand', how='left')
merged_transactions = net_transactions_specified.merge(net_transactions_specified_ly, on='Brand', how='left')

# Merge the sales and transactions data
merged_data = merged_sales.merge(merged_transactions, on='Brand', how='left')

# Calculate the Net AOV for TY and LY using safe_divide
merged_data['TY_Net_ASP'] = merged_data.apply(lambda row: safe_divide(row['TY_Net_Sales'], row['TY_Orders']), axis=1)
merged_data['LY_Net_ASP'] = merged_data.apply(lambda row: safe_divide(row['LY_Net_Sales'], row['LY_Orders']), axis=1)

# Calculate vs LY as percentage using safe_divide
merged_data['vs_LY'] = merged_data.apply(lambda row: safe_divide(row['TY_Net_ASP'] - row['LY_Net_ASP'], row['LY_Net_ASP']) * 100, axis=1)

# Select relevant columns
result = merged_data[['Brand', 'TY_Net_ASP', 'LY_Net_ASP', 'vs_LY']]

# Replace brand codes with brand names (assuming BCodes.csv has columns 'Code' and 'Name')
bcodes = bcodes.rename(columns={'Code': 'Brand', 'Name': 'Brand_Name'})
result = result.merge(bcodes, on='Brand', how='left')

# If the merge results in NaN Brand_Name, it means the Brand in the data is not present in BCodes
result['Brand'] = result['Brand_Name'].combine_first(result['Brand'])
result = result.drop(columns=['Brand_Name'])

# Save the result to a CSV file and Excel file
result.to_csv('C:/Users/Aliya/Downloads/quarterly_net_asp.csv', index=False)
result.to_excel('C:/Users/Aliya/Downloads/quarterly_net_asp.xlsx', index=False)

print("Net ASP data has been successfully saved.")
