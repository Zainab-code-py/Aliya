import pandas as pd
from datetime import datetime

def safe_divide(numerator, denominator):
    return numerator / denominator if denominator != 0 and numerator != 0 else 0

# Load the data
net_sales = pd.read_csv('C:/Users/Aliya/Downloads/yearly_net_sales.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
net_transactions = pd.read_csv('C:/Users/Aliya/Downloads/yearly_net_transactions.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)

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

# Group and sum the net sales by brand for TY and LY
net_sales_specified = net_sales.groupby('Brand').agg({'TY': 'sum', 'LY': 'sum'}).reset_index()
net_sales_specified.columns = ['Brand', 'TY_Net_Sales', 'LY_Net_Sales']

# Group and sum the net transactions by brand for TY and LY
net_transactions_specified = net_transactions.groupby('Brand').agg({'TY Sales': 'sum', 'LY Sales': 'sum'}).reset_index()
net_transactions_specified.columns = ['Brand', 'TY_Orders', 'LY_Orders']

# Merge the sales and transactions data
merged_data = net_sales_specified.merge(net_transactions_specified, on='Brand', how='left')

# Calculate the Net AOV for TY and LY using safe_divide
merged_data['TY_Net_AOV'] = merged_data.apply(lambda row: safe_divide(row['TY_Net_Sales'], row['TY_Orders']), axis=1)
merged_data['LY_Net_AOV'] = merged_data.apply(lambda row: safe_divide(row['LY_Net_Sales'], row['LY_Orders']), axis=1)

# Calculate vs LY as percentage using safe_divide
merged_data['vs_LY'] = merged_data.apply(lambda row: safe_divide(row['TY_Net_AOV'] - row['LY_Net_AOV'], row['LY_Net_AOV']) * 100, axis=1)

# Select relevant columns
result = merged_data[['Brand', 'TY_Net_AOV', 'LY_Net_AOV', 'vs_LY']]

# Save the result to a CSV file and Excel file
result.to_csv('C:/Users/Aliya/Downloads/yearly_net_aov_ytd.csv', index=False)
result.to_excel('C:/Users/Aliya/Downloads/yearly_net_aov_ytd.xlsx', index=False)


