import pandas as pd
from datetime import datetime, timedelta
import numpy as np

# Load the data with specified dtypes to avoid mixed type warning
offline_sales = pd.read_csv('C:/Users/Aliya/Downloads/Filtered_OfflineSales.csv', dtype={'Country': str, 'Brand': str}, low_memory=False)
currencies = pd.read_excel('C:/Users/Aliya/Downloads/Currencies.xlsx')
countries = pd.read_csv('C:/Users/Aliya/Downloads/Countries.csv')
calendar = pd.read_csv('C:/Users/Aliya/Downloads/Calendarx.csv')

# Define the date range for 2nd April to 30th June for 2024 and the same period in 2023
start_date_2024 = datetime(2024, 4, 2)
end_date_2024 = datetime(2024, 6, 30)
start_date_2023 = datetime(2023, 4, 2)
end_date_2023 = datetime(2023, 6, 30)

# Convert Transaction Date to datetime and filter sales data for TW and LW
offline_sales['Transaction Date'] = pd.to_datetime(offline_sales['Transaction Date'], format='%d-%b-%y', errors='coerce')

# Remove commas and convert Revenue to numeric, handling errors
offline_sales['Revenue'] = offline_sales['Revenue'].str.replace(',', '')
offline_sales['Revenue'] = pd.to_numeric(offline_sales['Revenue'], errors='coerce')

# Filter sales data for the specified date ranges
filtered_sales_ty = offline_sales[
    (offline_sales['Transaction Date'] >= start_date_2024) &
    (offline_sales['Transaction Date'] <= end_date_2024)
    ].copy()

filtered_sales_ly = offline_sales[
    (offline_sales['Transaction Date'] >= start_date_2023) &
    (offline_sales['Transaction Date'] <= end_date_2023)
    ].copy()

# Prepare the result DataFrame
result_offline = pd.DataFrame(columns=['Brand', 'TY', 'LY', 'vs LY'])

# Get unique brands from TY and LY sales data
brands_ty = filtered_sales_ty['Brand'].unique()
brands_ly = filtered_sales_ly['Brand'].unique()
all_brands = set(brands_ty).union(set(brands_ly))

# Helper function to get the month column name in the format "Month-Year"
def get_month_column_name(date):
    return date.strftime('%B-%Y')  # Using full month name

# Calculate the total amount for each brand for TY and LY
for brand in all_brands:
    total_amount_usd_ty = 0
    total_amount_usd_ly = 0

    # Calculate TY sales
    brand_sales_ty = filtered_sales_ty[filtered_sales_ty['Brand'] == brand]
    country_names_ty = brand_sales_ty['Country'].unique()

    for country_name in country_names_ty:
        country_sales_ty = brand_sales_ty[brand_sales_ty['Country'] == country_name]

        # Split the sales data based on the months in the date range
        sales_by_month = country_sales_ty.groupby(country_sales_ty['Transaction Date'].dt.to_period('M'))

        for period, sales in sales_by_month:
            total_revenue_ty = sales['Revenue'].sum()

            # Find the country code for the country
            country_code = countries.loc[countries['Country Name'] == country_name, 'Country Code'].values
            if len(country_code) == 0:
                continue
            country_code = country_code[0]

            # Adjust period format to match currency DataFrame column names
            period_str = get_month_column_name(period.start_time)
            if period_str not in currencies.columns:
                continue

            conversion_rate_ty = currencies.loc[currencies['Country_Code'] == country_code, period_str].values
            if len(conversion_rate_ty) > 0 and conversion_rate_ty[0] > 0:
                amount_usd_ty = total_revenue_ty / conversion_rate_ty[0]
            else:
                amount_usd_ty = 0

            total_amount_usd_ty += amount_usd_ty

    # Calculate LY sales
    brand_sales_ly = filtered_sales_ly[filtered_sales_ly['Brand'] == brand]
    country_names_ly = brand_sales_ly['Country'].unique()

    for country_name in country_names_ly:
        country_sales_ly = brand_sales_ly[brand_sales_ly['Country'] == country_name]

        # Split the sales data based on the months in the date range
        sales_by_month = country_sales_ly.groupby(country_sales_ly['Transaction Date'].dt.to_period('M'))

        for period, sales in sales_by_month:
            total_revenue_ly = sales['Revenue'].sum()

            # Find the country code for the country
            country_code = countries.loc[countries['Country Name'] == country_name, 'Country Code'].values
            if len(country_code) == 0:
                continue
            country_code = country_code[0]

            # Adjust period format to match currency DataFrame column names
            period_str = get_month_column_name(period.start_time)
            if period_str not in currencies.columns:
                continue

            conversion_rate_ly = currencies.loc[currencies['Country_Code'] == country_code, period_str].values
            if len(conversion_rate_ly) > 0 and conversion_rate_ly[0] > 0:
                amount_usd_ly = total_revenue_ly / conversion_rate_ly[0]
            else:
                amount_usd_ly = 0

            total_amount_usd_ly += amount_usd_ly

    # Calculate vs LY
    vs_ly = (total_amount_usd_ty - total_amount_usd_ly) / total_amount_usd_ly * 100 if total_amount_usd_ly != 0 else 0

    # Round the TY and LY values to whole numbers
    total_amount_usd_ty = np.ceil(total_amount_usd_ty)
    total_amount_usd_ly = np.ceil(total_amount_usd_ly)

    # Round vs LY to 2 decimal places and add a percent sign
    vs_ly = f"{round(vs_ly, 2)}%"

    # Append result
    result_offline = pd.concat([
        result_offline,
        pd.DataFrame({'Brand': [brand], 'TY': [total_amount_usd_ty], 'LY': [total_amount_usd_ly], 'vs LY': [vs_ly]})
    ], ignore_index=True)

# Save the result to a CSV file and Excel sheet
result_offline.to_csv('C:/Users/Aliya/Downloads/Offline_YTD.csv', index=False)
result_offline.to_excel('C:/Users/Aliya/Downloads/Offline_YTD.xlsx', index=False)

# Debug information
print("Filtered Sales Data TY:\n", filtered_sales_ty)
print("Result Offline Sales:\n", result_offline)
