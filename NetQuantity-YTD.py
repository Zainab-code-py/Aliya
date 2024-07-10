import pandas as pd
from datetime import datetime

# Load the data with specified dtypes to avoid mixed type warning
net_sales = pd.read_csv('C:/Users/Aliya/Downloads/NetSales_modified.csv')
bcodes = pd.read_csv('C:/Users/Aliya/Downloads/BCodes.csv')
currencies = pd.read_excel('C:/Users/Aliya/Downloads/Currencies.xlsx')

# Define the date range for YTD (1st January to 30th June) for this year and last year
start_date_ytd = datetime(datetime.now().year, 1, 1)
end_date_ytd = datetime(datetime.now().year, 6, 30)
start_date_ytd_ly = datetime(datetime.now().year - 1, 1, 1)
end_date_ytd_ly = datetime(datetime.now().year - 1, 6, 30)

# Function to parse dates with different formats
def parse_dates(date_str):
    for fmt in ("%d/%m/%Y %H:%M", "%m/%d/%Y %I:%M:%S %p", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return pd.NaT

# Apply the date parsing function to the Finance_Date column
net_sales['Finance_Date'] = net_sales['Finance_Date'].apply(parse_dates)

# Filter net sales data for the YTD date range
filtered_net_sales_ytd = net_sales[
    (net_sales['Finance_Date'] >= start_date_ytd) &
    (net_sales['Finance_Date'] <= end_date_ytd)
].copy()

filtered_net_sales_ytd_ly = net_sales[
    (net_sales['Finance_Date'] >= start_date_ytd_ly) &
    (net_sales['Finance_Date'] <= end_date_ytd_ly)
].copy()

# Convert Quantity to appropriate data type
filtered_net_sales_ytd['Quantity'] = pd.to_numeric(filtered_net_sales_ytd['Quantity'], errors='coerce')
filtered_net_sales_ytd_ly['Quantity'] = pd.to_numeric(filtered_net_sales_ytd_ly['Quantity'], errors='coerce')

# Prepare the result DataFrame
result = pd.DataFrame(columns=['Brand', 'TY Quantity', 'Budget Quantity', 'vs Budget', 'LY Quantity', 'vs LY'])

# Get unique brands from the net sales data
brands_ytd = filtered_net_sales_ytd['Brand'].unique()

# Calculate the total quantity for each brand for the specified range and last year
for brand in brands_ytd:
    total_quantity_ty = 0
    total_quantity_budget = 0
    total_quantity_ly = 0

    # Calculate quantity for this year YTD
    brand_net_sales_ytd = filtered_net_sales_ytd[filtered_net_sales_ytd['Brand'] == brand]
    total_quantity_ty = brand_net_sales_ytd['Quantity'].sum()

    # Calculate quantity for last year YTD
    brand_net_sales_ytd_ly = filtered_net_sales_ytd_ly[filtered_net_sales_ytd_ly['Brand'] == brand]
    total_quantity_ly = brand_net_sales_ytd_ly['Quantity'].sum()
    total_quantity_budget = total_quantity_ly  # Using last year's quantity as the budget

    # Calculate vs Budget as percentage
    vs_budget = 0
    if total_quantity_budget > 0:
        vs_budget = ((total_quantity_ty - total_quantity_budget) / total_quantity_budget) * 100

    # Calculate vs LY as percentage
    vs_ly = 0
    if total_quantity_ly > 0:
        vs_ly = ((total_quantity_ty - total_quantity_ly) / total_quantity_ly) * 100

    # Append result
    result = pd.concat([
        result,
        pd.DataFrame({
            'Brand': [brand],
            'TY Quantity': [total_quantity_ty],
            'Budget Quantity': [total_quantity_budget],
            'vs Budget': [vs_budget],
            'LY Quantity': [total_quantity_ly],
            'vs LY': [vs_ly]
        })
    ], ignore_index=True)

# Replace brand codes with brand names
result = result.merge(bcodes, left_on='Brand', right_on='Code')
result = result.drop(columns=['Brand', 'Code'])
result = result.rename(columns={'Name': 'Brand'})

# Save the result to a CSV file and Excel file
result.to_csv('C:/Users/Aliya/Downloads/yearly_net_quantity.csv', index=False)
result.to_excel('C:/Users/Aliya/Downloads/yearly_net_quantity.xlsx', index=False)

# Print the result DataFrame
print("Result DataFrame:\n", result)
