import pandas as pd

# Load the pre-generated NET SALES and GROSS SALES files
net_sales_path = 'C:/Users/Aliya/Downloads/quarterly_net_sales.csv'
gross_sales_path = 'C:/Users/Aliya/Downloads/quarterly_gross_sales.csv'
bcodes_path = 'C:/Users/Aliya/Downloads/BCodes.csv'

# Load the data
net_sales = pd.read_csv(net_sales_path, dtype={'Brand': str, 'NET SALES': float}, low_memory=False)
gross_sales = pd.read_csv(gross_sales_path, dtype={'Brand': str, 'GROSS SALES': float}, low_memory=False)
bcodes = pd.read_csv(bcodes_path, dtype={'Code': str, 'Name': str}, low_memory=False)

# Summarize the net sales and gross sales by brand
net_sales_summary = net_sales.groupby('Brand')['NET SALES'].sum().reset_index()
gross_sales_summary = gross_sales.groupby('Brand')['GROSS SALES'].sum().reset_index()

# Merge the net sales and gross sales data
merged_data = net_sales_summary.merge(gross_sales_summary, on='Brand', how='left')

# Calculate the Gross to Net percentage
merged_data['Gross_to_Net%'] = merged_data['NET SALES'] / merged_data['GROSS SALES']

# Replace brand codes with brand names
result = merged_data.merge(bcodes, left_on='Brand', right_on='Code', how='left')
result = result.drop(columns=['Brand', 'Code'])
result = result.rename(columns={'Name': 'Brand'})

# Select relevant columns
result = result[['Brand', 'NET SALES', 'GROSS SALES', 'Gross_to_Net%']]

# Save the result to a CSV file and Excel file
result.to_csv('C:/Users/Aliya/Downloads/gross_to_net_percentage.csv', index=False)
result.to_excel('C:/Users/Aliya/Downloads/gross_to_net_percentage.xlsx', index=False)

print("Gross to Net percentage data has been successfully saved.")
