import pandas as pd

# Load the data
calendar = pd.read_csv('C:/Users/Aliya/Downloads/Calendarx_22_converted_offline_sales.csv', low_memory=False)
offline_sales = pd.read_csv('C:/Users/Aliya/Downloads/Customer Data Store.csv')
currencies= pd.read_excel('C:/Users/Aliya/Downloads/Currencies.xlsx')

# Remove commas and convert the 'Revenue' column to integers
offline_sales['Revenue'] = offline_sales['Revenue'].str.replace(',', '').astype(int)

# Ensure the 'Sort' column is in MM/DD/YYYY format
calendar['Sort'] = pd.to_datetime(calendar['Sort'], errors='coerce').dt.strftime('%d/%m/%Y')

# Ensure the 'Transaction Date' column is in MM/DD/YYYY format
offline_sales['Transaction Date'] = pd.to_datetime(offline_sales['Transaction Date'], errors='coerce').dt.strftime('%d/%m/%Y')

# Define the filtering function for the calendar data
def filter_calendar(calendar, start_date, end_date, fiscal_quarter, fiscal_year):
    filtered_calendar = calendar[
        (calendar['Sort'] >= start_date) &
        (calendar['Sort'] <= end_date) &
        (calendar['Fiscal Qtr'] == fiscal_quarter) &
        (calendar['Trading Year'] == fiscal_year)
    ]
    return filtered_calendar

# Define the start and end dates as strings in MM/DD/YYYY format
start_date_2024 = '04/02/2024'
end_date_2024 = '06/30/2024'
start_date_2023 = '04/02/2023'
end_date_2023 = '06/30/2023'

# Filter calendar data for the specified date range, fiscal quarter, and fiscal year
filtered_calendar_2024 = filter_calendar(calendar, start_date_2024, end_date_2024, 'Q2', 'FY 24-25')
filtered_calendar_2023 = filter_calendar(calendar, start_date_2023, end_date_2023, 'Q2', 'FY 23-24')

# Function to filter sales data based on calendar dates
def filter_sales(offline_sales, filtered_calendar):
    if not filtered_calendar.empty:
        filtered_dates = filtered_calendar['Sort'].tolist()
        filtered_sales = offline_sales[offline_sales['Transaction Date'].isin(filtered_dates)]
        print(filtered_sales)
        return filtered_sales
    else:
        return pd.DataFrame(columns=offline_sales.columns)

# Filter offline sales data for the specified date range
filtered_sales_2024 = filter_sales(offline_sales, filtered_calendar_2024)
filtered_sales_2023 = filter_sales(offline_sales, filtered_calendar_2023)

# Prepare the result DataFrame
result = pd.DataFrame(columns=['Brand', 'TY', 'LY', 'vs LY'])

# Get unique brands from both TY and LY sales data
brands_ty = filtered_sales_2024['Brand'].unique() if not filtered_sales_2024.empty else []
brands_ly = filtered_sales_2023['Brand'].unique() if not filtered_sales_2023.empty else []
all_brands = set(brands_ty).union(set(brands_ly))

# Placeholder for currencies data
# Assuming currencies is already loaded, convert currency rates to numeric
# currencies['June-2024'] = pd.to_numeric(currencies['June-2024'], errors='coerce')
# currencies['June-2023'] = pd.to_numeric(currencies['June-2023'], errors='coerce')

# Calculate the total amount for each brand for TY and LY
for brand in all_brands:
    total_amount_usd_ty = 0
    total_amount_usd_ly = 0

    # Calculate TY sales
    if brand in brands_ty:
        brand_sales_ty = filtered_sales_2024[filtered_sales_2024['Brand'] == brand]
        country_codes_ty = brand_sales_ty['Country Code'].unique()

        for country_code in country_codes_ty:
            country_sales_ty = brand_sales_ty[brand_sales_ty['Country Code'] == country_code]
            total_amount_ty = country_sales_ty['Revenue'].sum()
            #Placeholder for conversion rate calculation
            conversion_rate_ty = currencies.loc[currencies['Country_Code'] == country_code, 'June-2024'].values
            if len(conversion_rate_ty) > 0 and conversion_rate_ty > 0:
                amount_usd_ty = total_amount_ty / conversion_rate_ty

            else:
                amount_usd_ty = 0
            amount_usd_ty = total_amount_ty  # Placeholder
            total_amount_usd_ty += amount_usd_ty

    # Calculate LY sales
    if brand in brands_ly:
        brand_sales_ly = filtered_sales_2023[filtered_sales_2023['Brand'] == brand]
        country_codes_ly = brand_sales_ly['Country Code'].unique()

        for country_code in country_codes_ly:
            country_sales_ly = brand_sales_ly[brand_sales_ly['Country Code'] == country_code]
            total_amount_ly = country_sales_ly['Revenue'].sum()
            # Placeholder for conversion rate calculation
            # conversion_rate_ly = currencies.loc[currencies['Country_Code'] == country_code, 'June-2023'].values
            # if len(conversion_rate_ly) > 0 and conversion_rate_ly[0] > 0:
            #     amount_usd_ly = total_amount_ly / conversion_rate_ly[0]
            # else:
            #     amount_usd_ly = 0
            amount_usd_ly = total_amount_ly  # Placeholder
            total_amount_usd_ly += amount_usd_ly

    # Calculate vs LY
    if total_amount_usd_ly > 0:
        vs_ly = ((total_amount_usd_ty - total_amount_usd_ly) / total_amount_usd_ly) * 100
    else:
        vs_ly = float('inf') if total_amount_usd_ty > 0 else 0

    result = pd.concat([
        result,
        pd.DataFrame({
            'Brand': [brand],
            'TY': [total_amount_usd_ty],
            'LY': [total_amount_usd_ly],
            'vs LY': [vs_ly]
        })
    ], ignore_index=True)

# Save the result to a CSV file
result_path = 'C:/Users/Aliya/Downloads/Offline-QTD.csv'
result.to_csv(result_path, index=False)

print(result_path)
