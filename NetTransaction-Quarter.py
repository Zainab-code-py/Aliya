import pandas as pd
from datetime import datetime

# Load the data with specified dtypes to avoid mixed type warning
net_sales = pd.read_csv('C:/Users/Aliya/Downloads/NetSales_modified.csv')
bcodes = pd.read_csv('C:/Users/Aliya/Downloads/BCodes.csv')
currencies = pd.read_excel('C:/Users/Aliya/Downloads/Currencies.xlsx')

# Define the date range for the fiscal quarter (2nd April to 30th June) for this year and last year
specified_start = datetime(2024, 4, 2)
specified_end = datetime(2024, 6, 30)
specified_start_ly = datetime(2023, 4, 2)
specified_end_ly = datetime(2023, 6, 30)

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

# Filter net sales data for the specified date range
filtered_net_sales_specified = net_sales[
    (net_sales['Finance_Date'] >= specified_start) &
    (net_sales['Finance_Date'] <= specified_end)
].copy()

filtered_net_sales_specified_ly = net_sales[
    (net_sales['Finance_Date'] >= specified_start_ly) &
    (net_sales['Finance_Date'] <= specified_end_ly)
].copy()

# Prepare the result DataFrame
result = pd.DataFrame(columns=['Brand', 'TY Transactions', 'Budget Transactions', 'vs Budget', 'LY Transactions', 'vs LY'])

# Get unique brands from the net sales data
brands_specified = filtered_net_sales_specified['Brand'].unique()

# Calculate the total transactions for each brand for the specified range and last year
for brand in brands_specified:
    total_transactions_ty = 0
    total_transactions_budget = 0
    total_transactions_ly = 0

    # Calculate transactions for the specified period
    brand_net_sales_specified = filtered_net_sales_specified[filtered_net_sales_specified['Brand'] == brand]
    total_transactions_ty = brand_net_sales_specified['Order No.'].nunique()

    # Calculate transactions for the same period last year
    brand_net_sales_specified_ly = filtered_net_sales_specified_ly[filtered_net_sales_specified_ly['Brand'] == brand]
    total_transactions_ly = brand_net_sales_specified_ly['Order No.'].nunique()
    total_transactions_budget = total_transactions_ly  # Using last year's transactions as the budget

    # Calculate vs Budget as percentage
    vs_budget = 0
    if total_transactions_budget > 0:
        vs_budget = ((total_transactions_ty - total_transactions_budget) / total_transactions_budget) * 100

    # Calculate vs LY as percentage
    vs_ly = 0
    if total_transactions_ly > 0:
        vs_ly = ((total_transactions_ty - total_transactions_ly) / total_transactions_ly) * 100

    # Append result
    result = pd.concat([
        result,
        pd.DataFrame({
            'Brand': [brand],
            'TY Transactions': [total_transactions_ty],
            'Budget Transactions': [total_transactions_budget],
            'vs Budget': [vs_budget],
            'LY Transactions': [total_transactions_ly],
            'vs LY': [vs_ly]
        })
    ], ignore_index=True)

# Replace brand codes with brand names
result = result.merge(bcodes, left_on='Brand', right_on='Code')
result = result.drop(columns=['Brand', 'Code'])
result = result.rename(columns={'Name': 'Brand'})

# Save the result to a CSV file and Excel file
result.to_csv('quarterly_net_transactions.csv', index=False)

# Print the result DataFrame
print("Result DataFrame:\n", result)
