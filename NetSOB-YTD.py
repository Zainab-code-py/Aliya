import pandas as pd

# Load the CSV files
file_path1 = 'C:/Users/Aliya/Downloads/quarterly_net_sales.csv'
file_path2 = 'C:/Users/Aliya/Downloads/Filtered_OfflineSales.csv'

quarterly_net_sales = pd.read_csv(file_path1)
quarterly_sales = pd.read_csv(file_path2)

# Convert relevant columns to numeric, forcing errors to NaN and then filling NaN with 0
quarterly_net_sales['TY'] = pd.to_numeric(quarterly_net_sales['TY'], errors='coerce').fillna(0)
quarterly_net_sales['LY'] = pd.to_numeric(quarterly_net_sales['LY'], errors='coerce').fillna(0)

# Convert 'Transaction Date' to datetime
quarterly_sales['Transaction Date'] = pd.to_datetime(quarterly_sales['Transaction Date'], format='%d-%b-%y')

# Filter the date range for TY (This Year)
filtered_sales_TY = quarterly_sales[(quarterly_sales['Transaction Date'] >= '2024-04-02') &
                                    (quarterly_sales['Transaction Date'] <= '2024-06-30')]

# Filter the date range for LY (Last Year)
filtered_sales_LY = quarterly_sales[(quarterly_sales['Transaction Date'] >= '2023-04-02') &
                                    (quarterly_sales['Transaction Date'] <= '2023-06-30')]

# Aggregate the filtered sales data by 'Brand' for TY and LY
aggregated_sales_TY = filtered_sales_TY.groupby('Brand').agg({'Revenue': 'sum'}).reset_index()
aggregated_sales_TY.columns = ['Brand', 'TY_sales']

aggregated_sales_LY = filtered_sales_LY.groupby('Brand').agg({'Revenue': 'sum'}).reset_index()
aggregated_sales_LY.columns = ['Brand', 'LY_sales']

# Merge datasets on 'Brand'
merged_data = pd.merge(quarterly_net_sales, aggregated_sales_TY, on='Brand', how='left')
merged_data = pd.merge(merged_data, aggregated_sales_LY, on='Brand', how='left')

# Convert the 'TY_sales' and 'LY_sales' columns to numeric, forcing errors to NaN and then filling NaN with 0
merged_data['TY_sales'] = pd.to_numeric(merged_data['TY_sales'], errors='coerce').fillna(0)
merged_data['LY_sales'] = pd.to_numeric(merged_data['LY_sales'], errors='coerce').fillna(0)

# Calculate the Net Sales SOB for TY and LY
merged_data['Net Sales SOB TY'] = merged_data['TY'] / (merged_data['TY_sales'] + merged_data['TY'])
merged_data['Net Sales SOB LY'] = merged_data['LY'] / (merged_data['LY_sales'] + merged_data['LY'])

# Prepare the output DataFrame with the required columns
output_data = merged_data[['Brand', 'TY', 'LY', 'TY_sales', 'LY_sales']].copy()
output_data.columns = ['Brand', 'TY', 'LY', 'Offline Sales TY', 'Offline Sales LY']

# Calculate 'vs LY'
output_data['vs LY'] = (output_data['TY'] - output_data['LY']) / output_data['LY'] * 100

# Rearrange columns to match the required output format
output_data = output_data[['Brand', 'TY', 'LY', 'vs LY']]

# Paths for the output files
output_csv_path = 'C:/Users/Aliya/Downloads/Net_Sales_SOB_Q2_2024YTD.csv'
output_excel_path = 'C:/Users/Aliya/Downloads/Net_Sales_SOB_Q2_2024YTD.xlsx'

# Save the output to a new CSV file
output_data.to_csv(output_csv_path, index=False)

# Save the output to a new Excel file
output_data.to_excel(output_excel_path, index=False)

print(f"CSV file saved to: {output_csv_path}")
print(f"Excel file saved to: {output_excel_path}")
