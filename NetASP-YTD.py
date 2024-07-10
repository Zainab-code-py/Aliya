import pandas as pd
from datetime import datetime

def safe_divide(numerator, denominator):
    return numerator / denominator if denominator != 0 and numerator != 0 else 0

# Load the data
net_sales = pd.read_csv('C:/Users/Aliya/Downloads/yearly_net_sales.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
net_quantity = pd.read_csv('C:/Users/Aliya/Downloads/yearly_net_quantity.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
bcodes = pd.read_csv('C:/Users/Aliya/Downloads/BCodes.csv', dtype={'Code': str}, low_memory=False)

# Ensure the brand columns are strings
net_sales['Brand'] = net_sales['Brand'].astype(str)
net_quantity['Brand'] = net_quantity['Brand'].astype(str)

# Strip any leading/trailing spaces from column names
net_sales.columns = net_sales.columns.str.strip()
net_quantity.columns = net_quantity.columns.str.strip()

# Define the date range for YTD (1st January to 30th June) for this year and last year
start_date_ytd = datetime(datetime.now().year, 1, 1)
end_date_ytd = datetime(datetime.now().year, 6, 30)
start_date_ytd_ly = datetime(datetime.now().year - 1, 1, 1)
end_date_ytd_ly = datetime(datetime.now().year - 1, 6, 30)

# Convert 'Date' columns to datetime if they exist in the datasets
if 'Date' in net_sales.columns:
    net_sales['Date'] = pd.to_datetime(net_sales['Date'])
if 'Date' in net_quantity.columns:
    net_quantity['Date'] = pd.to_datetime(net_quantity['Date'])

# Filter the data to include only the YTD range
if 'Date' in net_sales.columns:
    net_sales_ytd = net_sales[(net_sales['Date'] >= start_date_ytd) & (net_sales['Date'] <= end_date_ytd)]
    net_sales_ytd_ly = net_sales[(net_sales['Date'] >= start_date_ytd_ly) & (net_sales['Date'] <= end_date_ytd_ly)]
else:
    net_sales_ytd = net_sales
    net_sales_ytd_ly = net_sales

if 'Date' in net_quantity.columns:
    net_quantity_ytd = net_quantity[(net_quantity['Date'] >= start_date_ytd) & (net_quantity['Date'] <= end_date_ytd)]
    net_quantity_ytd_ly = net_quantity[(net_quantity['Date'] >= start_date_ytd_ly) & (net_quantity['Date'] <= end_date_ytd_ly)]
else:
    net_quantity_ytd = net_quantity
    net_quantity_ytd_ly = net_quantity

# Group and sum the net sales and quantity by brand for TY and LY
net_sales_specified = net_sales_ytd.groupby('Brand').agg({'TY': 'sum'}).reset_index()
net_sales_specified.columns = ['Brand', 'TY_Net_Sales']

net_quantity_specified = net_quantity_ytd.groupby('Brand').agg({'TY Quantity': 'sum'}).reset_index()
net_quantity_specified.columns = ['Brand', 'TY_Quantity']

net_sales_specified_ly = net_sales_ytd_ly.groupby('Brand').agg({'LY': 'sum'}).reset_index()
net_sales_specified_ly.columns = ['Brand', 'LY_Net_Sales']

net_quantity_specified_ly = net_quantity_ytd_ly.groupby('Brand').agg({'LY Quantity': 'sum'}).reset_index()
net_quantity_specified_ly.columns = ['Brand', 'LY_Quantity']

# Merge the TY and LY data for net sales and quantity
merged_sales = net_sales_specified.merge(net_sales_specified_ly, on='Brand', how='left')
merged_quantity = net_quantity_specified.merge(net_quantity_specified_ly, on='Brand', how='left')

# Merge the sales and quantity data
merged_data = merged_sales.merge(merged_quantity, on='Brand', how='left')

# Calculate the Net ASP for TY and LY using safe_divide
merged_data['TY_Net_ASP'] = merged_data.apply(lambda row: safe_divide(row['TY_Net_Sales'], row['TY_Quantity']), axis=1)
merged_data['LY_Net_ASP'] = merged_data.apply(lambda row: safe_divide(row['LY_Net_Sales'], row['LY_Quantity']), axis=1)

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
result.to_csv('C:/Users/Aliya/Downloads/yearly_net_asp_ytd.csv', index=False)
result.to_excel('C:/Users/Aliya/Downloads/yearly_net_asp_ytd.xlsx', index=False)

# Print the result
print("Result DataFrame:\n", result)
