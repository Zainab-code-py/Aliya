import pandas as pd
from datetime import datetime

def safe_divide(numerator, denominator):
    return numerator / denominator if denominator != 0 and numerator != 0 else 0

# Load the data
net_quantity = pd.read_csv('C:/Users/Aliya/Downloads/yearly_net_quantity.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
net_transactions = pd.read_csv('C:/Users/Aliya/Downloads/yearly_net_transactions.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
bcodes = pd.read_csv('C:/Users/Aliya/Downloads/BCodes.csv', dtype={'Code': str}, low_memory=False)

# Ensure the brand columns are strings
net_quantity['Brand'] = net_quantity['Brand'].astype(str)
net_transactions['Brand'] = net_transactions['Brand'].astype(str)

# Strip any leading/trailing spaces from column names
net_quantity.columns = net_quantity.columns.str.strip()
net_transactions.columns = net_transactions.columns.str.strip()

# Define the date range for YTD (1st January to 30th June) for this year and last year
start_date_ytd = datetime(datetime.now().year, 1, 1)
end_date_ytd = datetime(datetime.now().year, 6, 30)
start_date_ytd_ly = datetime(datetime.now().year - 1, 1, 1)
end_date_ytd_ly = datetime(datetime.now().year - 1, 6, 30)

# Filter the data to include only the YTD range
if 'Date' in net_quantity.columns:
    net_quantity['Date'] = pd.to_datetime(net_quantity['Date'])
    net_quantity_ytd = net_quantity[(net_quantity['Date'] >= start_date_ytd) & (net_quantity['Date'] <= end_date_ytd)]
    net_quantity_ytd_ly = net_quantity[(net_quantity['Date'] >= start_date_ytd_ly) & (net_quantity['Date'] <= end_date_ytd_ly)]
else:
    net_quantity_ytd = net_quantity
    net_quantity_ytd_ly = net_quantity

if 'Date' in net_transactions.columns:
    net_transactions['Date'] = pd.to_datetime(net_transactions['Date'])
    net_transactions_ytd = net_transactions[(net_transactions['Date'] >= start_date_ytd) & (net_transactions['Date'] <= end_date_ytd)]
    net_transactions_ytd_ly = net_transactions[(net_transactions['Date'] >= start_date_ytd_ly) & (net_transactions['Date'] <= end_date_ytd_ly)]
else:
    net_transactions_ytd = net_transactions
    net_transactions_ytd_ly = net_transactions

# Find common brands in both datasets
common_brands = set(net_quantity_ytd['Brand']).intersection(set(net_transactions_ytd['Brand']))

# Filter the data to include only common brands
net_quantity_ytd = net_quantity_ytd[net_quantity_ytd['Brand'].isin(common_brands)]
net_transactions_ytd = net_transactions_ytd[net_transactions_ytd['Brand'].isin(common_brands)]
net_quantity_ytd_ly = net_quantity_ytd_ly[net_quantity_ytd_ly['Brand'].isin(common_brands)]
net_transactions_ytd_ly = net_transactions_ytd_ly[net_transactions_ytd_ly['Brand'].isin(common_brands)]

# Group and sum the net quantity and transactions by brand for TY and LY
net_quantity_specified = net_quantity_ytd.groupby('Brand').agg({'TY Quantity': 'sum'}).reset_index()
net_quantity_specified.columns = ['Brand', 'TY_Quantity']

net_transactions_specified = net_transactions_ytd.groupby('Brand').agg({'TY Sales': 'sum'}).reset_index()
net_transactions_specified.columns = ['Brand', 'TY_Orders']

net_quantity_specified_ly = net_quantity_ytd_ly.groupby('Brand').agg({'LY Quantity': 'sum'}).reset_index()
net_quantity_specified_ly.columns = ['Brand', 'LY_Quantity']

net_transactions_specified_ly = net_transactions_ytd_ly.groupby('Brand').agg({'LY Sales': 'sum'}).reset_index()
net_transactions_specified_ly.columns = ['Brand', 'LY_Orders']

# Merge the TY and LY data for net quantity and transactions
merged_quantity = net_quantity_specified.merge(net_quantity_specified_ly, on='Brand', how='left')
merged_transactions = net_transactions_specified.merge(net_transactions_specified_ly, on='Brand', how='left')

# Merge the quantity and transactions data
merged_data = merged_quantity.merge(merged_transactions, on='Brand', how='left')

# Calculate the Net UPT for TY and LY using safe_divide
merged_data['TY_Net_UPT'] = merged_data.apply(lambda row: safe_divide(row['TY_Quantity'], row['TY_Orders']), axis=1)
merged_data['LY_Net_UPT'] = merged_data.apply(lambda row: safe_divide(row['LY_Quantity'], row['LY_Orders']), axis=1)

# Calculate vs LY as percentage using safe_divide
merged_data['vs_LY'] = merged_data.apply(lambda row: safe_divide(row['TY_Net_UPT'] - row['LY_Net_UPT'], row['LY_Net_UPT']) * 100, axis=1)

# Select relevant columns
result = merged_data[['Brand', 'TY_Net_UPT', 'LY_Net_UPT', 'vs_LY']]

# Replace brand codes with brand names (assuming BCodes.csv has columns 'Code' and 'Name')
bcodes = bcodes.rename(columns={'Code': 'Brand', 'Name': 'Brand_Name'})
result = result.merge(bcodes, on='Brand', how='left')

# If the merge results in NaN Brand_Name, it means the Brand in the data is not present in BCodes
result['Brand'] = result['Brand_Name'].combine_first(result['Brand'])
result = result.drop(columns=['Brand_Name'])

# Save the result to a CSV file and Excel file
result.to_csv('C:/Users/Aliya/Downloads/yearly_net_upt_ytd.csv', index=False)
result.to_excel('C:/Users/Aliya/Downloads/yearly_net_upt_ytd.xlsx', index=False)

print("Yearly Net UPT data has been successfully saved.")
