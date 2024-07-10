import pandas as pd
from datetime import datetime

# Load the data with specified dtypes to avoid mixed type warning
net_sales = pd.read_csv('C:/Users/Aliya/Downloads/NetSales_modified.csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
currencies = pd.read_excel('C:/Users/Aliya/Downloads/Currencies.xlsx')
bcodes = pd.read_csv('C:/Users/Aliya/Downloads/BCodes.csv')

# Define the date range for YTD (1st January to 30th June) for this year and last year
start_date_ytd = datetime(datetime.now().year, 1, 1)
end_date_ytd = datetime(datetime.now().year, 6, 30)
start_date_ytd_ly = datetime(datetime.now().year - 1, 1, 1)
end_date_ytd_ly = datetime(datetime.now().year - 1, 6, 30)

# Function to filter sales data based on date range
def filter_sales_by_date(net_sales, start_date, end_date):
    net_sales['Finance_Date'] = pd.to_datetime(net_sales['Finance_Date'], errors='coerce')
    return net_sales[(net_sales['Finance_Date'] >= start_date) & (net_sales['Finance_Date'] <= end_date)]

# Filter sales data for the specified date ranges
filtered_sales_ytd = filter_sales_by_date(net_sales, start_date_ytd, end_date_ytd)
filtered_sales_ytd_ly = filter_sales_by_date(net_sales, start_date_ytd_ly, end_date_ytd_ly)

# Convert Amount and currency rates to appropriate data types
filtered_sales_ytd['Amount Inc VAT'] = pd.to_numeric(filtered_sales_ytd['Amount Inc VAT'], errors='coerce')
filtered_sales_ytd_ly['Amount Inc VAT'] = pd.to_numeric(filtered_sales_ytd_ly['Amount Inc VAT'], errors='coerce')
currency_columns = [col for col in currencies.columns if col not in ['Country_Code', 'Country_Name']]
for col in currency_columns:
    currencies[col] = pd.to_numeric(currencies[col], errors='coerce')

# Remove negative values
filtered_sales_ytd = filtered_sales_ytd[filtered_sales_ytd['Amount Inc VAT'] >= 0]
filtered_sales_ytd_ly = filtered_sales_ytd_ly[filtered_sales_ytd_ly['Amount Inc VAT'] >= 0]

# Prepare the result DataFrame
result = pd.DataFrame(columns=['Brand', 'TY', 'LY', 'vs LY'])

# Get unique brands from TY and LY sales data
brands_ytd = filtered_sales_ytd['Brand'].unique()
brands_ytd_ly = filtered_sales_ytd_ly['Brand'].unique()
all_brands = set(brands_ytd).union(set(brands_ytd_ly))

# Calculate the total amount for each brand for TY and LY
for brand in all_brands:
    total_amount_usd_ty = 0
    total_amount_usd_ly = 0

    # Calculate TY sales
    if brand in brands_ytd:
        brand_sales_ty = filtered_sales_ytd[filtered_sales_ytd['Brand'] == brand]
        country_codes_ty = brand_sales_ty['Country Code'].unique()

        for country_code in country_codes_ty:
            country_sales_ty = brand_sales_ty[brand_sales_ty['Country Code'] == country_code]
            total_amount_ty = country_sales_ty['Amount Inc VAT'].sum()

            conversion_rate_ty = currencies.loc[currencies['Country_Code'] == country_code, 'June-2024'].values
            if len(conversion_rate_ty) > 0 and conversion_rate_ty[0] > 0:
                amount_usd_ty = total_amount_ty / conversion_rate_ty[0]
            else:
                amount_usd_ty = 0

            total_amount_usd_ty += amount_usd_ty

    # Calculate LY sales
    if brand in brands_ytd_ly:
        brand_sales_ly = filtered_sales_ytd_ly[filtered_sales_ytd_ly['Brand'] == brand]
        country_codes_ly = brand_sales_ly['Country Code'].unique()

        for country_code in country_codes_ly:
            country_sales_ly = brand_sales_ly[brand_sales_ly['Country Code'] == country_code]
            total_amount_ly = country_sales_ly['Amount Inc VAT'].sum()

            conversion_rate_ly = currencies.loc[currencies['Country_Code'] == country_code, 'June-2023'].values
            if len(conversion_rate_ly) > 0 and conversion_rate_ly[0] > 0:
                amount_usd_ly = total_amount_ly / conversion_rate_ly[0]
            else:
                amount_usd_ly = 0

            total_amount_usd_ly += amount_usd_ly

    # Calculate vs LY
    if total_amount_usd_ly > 0:
        vs_ly = ((total_amount_usd_ty - total_amount_usd_ly) / total_amount_usd_ly) * 100
    else:
        vs_ly = float('inf') if total_amount_usd_ty > 0 else 0

    # Append result
    result = pd.concat([
        result,
        pd.DataFrame({
            'Brand': [brand],
            'TY': [total_amount_usd_ty],
            'LY': [total_amount_usd_ly],
            'vs LY': [f"{vs_ly:.2f}%"]
        })
    ], ignore_index=True)

# Replace brand codes with brand names
result = result.merge(bcodes, left_on='Brand', right_on='Code')
result = result.drop(columns=['Brand', 'Code'])
result = result.rename(columns={'Name': 'Brand'})

# Reorder columns to place Brand first
result = result[['Brand', 'TY', 'LY', 'vs LY']]

# Print the result DataFrame
print("Result DataFrame:\n", result)

# Save the result to a CSV file and Excel file
result.to_csv('C:/Users/Aliya/Downloads/yearly_net_sales_ytd.csv', index=False)
result.to_excel('C:/Users/Aliya/Downloads/yearly_net_sales_ytd.xlsx', sheet_name='Net Sales', index=False)

print("Net sales data has been successfully saved.")
