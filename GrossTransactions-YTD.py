import pandas as pd
from datetime import datetime, timedelta

# Load the data with specified dtypes to avoid mixed type warning
gross_sales = pd.read_csv('C:/Users/Aliya/Downloads/GrossSales_transformed (1).csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
bcodes = pd.read_csv('C:/Users/Aliya/Downloads/BCodes.csv')
currencies = pd.read_excel('C:/Users/Aliya/Downloads/Currencies.xlsx')

# Define the date range for the fiscal quarter (2nd April to 30th June) for this year and last year
specified_start = datetime(2024, 4, 2)
specified_end = datetime(2024, 6, 30)
specified_start_ly = specified_start - timedelta(days=365)
specified_end_ly = specified_end - timedelta(days=365)
specified_ly_days = specified_start_ly + timedelta(days=80)

# Convert Order Date in gross_sales to datetime
gross_sales['Order Date'] = pd.to_datetime(gross_sales['Order Date'], errors='coerce', dayfirst=True)

# Filter gross sales data for the specified date range for TY and the first 80 days for LY
filtered_gross_sales_specified = gross_sales[
    (gross_sales['Order Date'] >= specified_start) &
    (gross_sales['Order Date'] <= specified_end)
].copy()

filtered_gross_sales_specified_ly = gross_sales[
    (gross_sales['Order Date'] >= specified_start_ly) &
    (gross_sales['Order Date'] <= specified_ly_days)
].copy()

# Print unique brands before the merge
print("Unique Brands before merge (this year):", filtered_gross_sales_specified['Brand'].unique())
print("Unique Brands before merge (last year):", filtered_gross_sales_specified_ly['Brand'].unique())

# Merge with bcodes to get brand names
filtered_gross_sales_specified = filtered_gross_sales_specified.merge(bcodes, left_on='Brand', right_on='Code', how='left')
filtered_gross_sales_specified_ly = filtered_gross_sales_specified_ly.merge(bcodes, left_on='Brand', right_on='Code', how='left')

# Handle missing brand names after the merge
filtered_gross_sales_specified['Name'] = filtered_gross_sales_specified['Name'].fillna(filtered_gross_sales_specified['Brand'])
filtered_gross_sales_specified_ly['Name'] = filtered_gross_sales_specified_ly['Name'].fillna(filtered_gross_sales_specified_ly['Brand'])

# Print unique brands after the merge
print("Unique Brands after merge (this year):", filtered_gross_sales_specified['Name'].unique())
print("Unique Brands after merge (last year):", filtered_gross_sales_specified_ly['Name'].unique())

# Prepare the result DataFrame
result = pd.DataFrame(columns=['Brand', 'TY', 'LY', 'vsLY'])

# Get unique brands from the gross sales data
brands_specified = filtered_gross_sales_specified['Brand'].unique()

# Calculate the total transactions for each brand for the specified range and last year
for brand in brands_specified:
    total_transactions_specified = 0
    total_transactions_specified_ly = 0

    # Calculate transactions for this year
    brand_gross_sales_specified = filtered_gross_sales_specified[filtered_gross_sales_specified['Brand'] == brand]
    total_transactions_specified = brand_gross_sales_specified['Order No.'].nunique()

    # Calculate transactions for last year
    if brand in filtered_gross_sales_specified_ly['Brand'].unique():
        brand_gross_sales_specified_ly = filtered_gross_sales_specified_ly[filtered_gross_sales_specified_ly['Brand'] == brand]
        total_transactions_specified_ly = brand_gross_sales_specified_ly['Order No.'].nunique()

    # Calculate vs LY as percentage
    vs_ly = 0
    if total_transactions_specified_ly > 0:
        vs_ly = ((total_transactions_specified - total_transactions_specified_ly) / total_transactions_specified_ly) * 100

    # Append result
    result = pd.concat([
        result,
        pd.DataFrame({
            'Brand': [brand_gross_sales_specified.iloc[0]['Name']],  # Use brand name from the mapping
            'TY': [total_transactions_specified],
            'LY': [total_transactions_specified_ly],
            'vsLY': [f"{vs_ly:.2f}%"]
        })
    ], ignore_index=True)

# Save the result to a CSV file and Excel file
result.to_csv('C:/Users/Aliya/Downloads/yearly_gross_transactions.csv', index=False)
result.to_excel('C:/Users/Aliya/Downloads/yearly_gross_transactions.xlsx', index=False)

print("Gross transactions data has been successfully saved.")
