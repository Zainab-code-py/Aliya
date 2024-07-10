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

# Convert Amount to appropriate data type
filtered_net_sales_specified['Amount Inc VAT'] = pd.to_numeric(filtered_net_sales_specified['Amount Inc VAT'], errors='coerce')
filtered_net_sales_specified_ly['Amount Inc VAT'] = pd.to_numeric(filtered_net_sales_specified_ly['Amount Inc VAT'], errors='coerce')

# Function to calculate the amount in USD for a given sales data, conversion column prefix
def calculate_amount_usd(sales_data, conversion_prefix):
    total_amount_usd = 0

    for _, row in sales_data.iterrows():
        order_date = row['Finance_Date']
        amount = row['Amount Inc VAT']
        country_code = row['Country Code']

        # Determine the conversion rate column based on the month of the order date
        month = order_date.strftime('%B')
        conversion_rate = currencies.loc[currencies['Country_Code'] == country_code, f'{month}-{conversion_prefix}'].values

        if conversion_rate is not None and len(conversion_rate) > 0 and conversion_rate[0] > 0:
            amount_usd = amount / conversion_rate[0]
        else:
            amount_usd = 0

        total_amount_usd += amount_usd

    return total_amount_usd

# Prepare the result DataFrame
result = pd.DataFrame(columns=['Brand', 'TY', 'Budget', 'vs Budget', 'LY', 'vs LY'])

# Get unique brands from the net sales data
brands_specified = filtered_net_sales_specified['Brand'].unique()

# Calculate the total amount for each brand for the specified range and last year
for brand in brands_specified:
    total_net_sales_usd = 0
    total_budget_usd = 0
    total_amount_usd_specified_ly = 0

    # Calculate net sales for the specified period
    brand_net_sales_specified = filtered_net_sales_specified[filtered_net_sales_specified['Brand'] == brand]
    total_net_sales_usd = calculate_amount_usd(brand_net_sales_specified, '2024')

    # Calculate net sales for the same period last year
    brand_net_sales_specified_ly = filtered_net_sales_specified_ly[filtered_net_sales_specified_ly['Brand'] == brand]
    total_amount_usd_specified_ly = calculate_amount_usd(brand_net_sales_specified_ly, '2023')
    total_budget_usd = total_amount_usd_specified_ly  # Using last year's sales as the budget

    # Calculate vs Budget as percentage
    vs_budget = 0
    if total_budget_usd > 0:
        vs_budget = ((total_net_sales_usd - total_budget_usd) / total_budget_usd) * 100

    # Calculate vs LY as percentage
    vs_ly = 0
    if total_amount_usd_specified_ly > 0:
        vs_ly = ((total_net_sales_usd - total_amount_usd_specified_ly) / total_amount_usd_specified_ly) * 100

    # Append result
    result = pd.concat([
        result,
        pd.DataFrame({
            'Brand': [brand],
            'TY': [total_net_sales_usd],
            'Budget': [total_budget_usd],
            'vs Budget': [vs_budget],
            'LY': [total_amount_usd_specified_ly],
            'vs LY': [vs_ly]
        })
    ], ignore_index=True)

# Replace brand codes with brand names
result = result.merge(bcodes, left_on='Brand', right_on='Code')
result = result.drop(columns=['Brand', 'Code'])
result = result.rename(columns={'Name': 'Brand'})

# Save the result to a CSV file and Excel file
result.to_csv('quarterly_net_sales.csv', index=False)

# Print the result DataFrame
print("Result DataFrame:\n", result)