import pandas as pd
from datetime import datetime, timedelta

# Load the uploaded files
countries_df = pd.read_csv('C:/Users/Aliya/Downloads/Countries.csv')
combined_data_df = pd.read_csv('C:/Users/Aliya/Downloads/All_Store_Transactions_Data.csv')
currencies_df = pd.read_excel('C:/Users/Aliya/Downloads/Currencies.xlsx')
calendar_df = pd.read_csv('C:/Users/Aliya/Downloads/Calendarx.csv')  # Assuming the Calendar.csv file

# Rename 'Country_Code' to 'Country Code' in currencies_df for consistency
currencies_df.rename(columns={'Country_Code': 'Country Code'}, inplace=True)

# Convert 'Transaction Date' to datetime format
combined_data_df['Transaction Date'] = pd.to_datetime(combined_data_df['Transaction Date'], format='%d-%b-%Y', errors='coerce')

# Convert 'Order Date' to datetime format
calendar_df['Order Date'] = pd.to_datetime(calendar_df['Order Date'], format='%d/%m/%Y')

# Define the date range for Q2
start_date = datetime(datetime.today().year, 4, 2)
end_date = datetime(datetime.today().year, 6, 26)

# Filter dates for this year (TY) in the specified range
filtered_data_ty = combined_data_df[(combined_data_df['Transaction Date'] >= start_date) &
                                    (combined_data_df['Transaction Date'] <= end_date)]

# Filter dates for last year (LY) in the same range
start_date_ly = start_date - timedelta(days=365)
end_date_ly = end_date - timedelta(days=365)

filtered_data_ly = combined_data_df[(combined_data_df['Transaction Date'] >= start_date_ly) &
                                    (combined_data_df['Transaction Date'] <= end_date_ly)]

# Merge filtered data with countries to get country code
filtered_data_ty = filtered_data_ty.merge(countries_df, left_on='Country', right_on='Country Name', how='left')
filtered_data_ly = filtered_data_ly.merge(countries_df, left_on='Country', right_on='Country Name', how='left')

# Rename columns properly
filtered_data_ty.rename(columns={'Country Code_y': 'Country Code'}, inplace=True)
filtered_data_ly.rename(columns={'Country Code_y': 'Country Code'}, inplace=True)

# Create a function to get conversion rate based on the transaction date
def get_conversion_rate(row, currency_df, column_prefix):
    month_year = row['Transaction Date'].strftime('%B-%Y')
    rate = currency_df.loc[currency_df['Country Code'] == row['Country Code'], month_year]
    return rate.values[0] if not rate.empty else None

# Apply the function to get conversion rates for each transaction
filtered_data_ty['Conversion Rate'] = filtered_data_ty.apply(get_conversion_rate, axis=1, args=(currencies_df, 'TY'))
filtered_data_ly['Conversion Rate'] = filtered_data_ly.apply(get_conversion_rate, axis=1, args=(currencies_df, 'LY'))

# Convert 'Revenue' column to numeric type
filtered_data_ty['Revenue'] = pd.to_numeric(filtered_data_ty['Revenue'], errors='coerce')
filtered_data_ly['Revenue'] = pd.to_numeric(filtered_data_ly['Revenue'], errors='coerce')

# Debugging: Check for null or missing values
print("TY missing values:\n", filtered_data_ty.isnull().sum())
print("LY missing values:\n", filtered_data_ly.isnull().sum())

# Calculate revenue in USD
filtered_data_ty['Revenue USD'] = filtered_data_ty['Revenue'] / filtered_data_ty['Conversion Rate']
filtered_data_ly['Revenue USD'] = filtered_data_ly['Revenue'] / filtered_data_ly['Conversion Rate']

# Debugging: Check for null or missing values in 'Revenue USD'
print("TY Revenue USD missing values:\n", filtered_data_ty['Revenue USD'].isnull().sum())
print("LY Revenue USD missing values:\n", filtered_data_ly['Revenue USD'].isnull().sum())

# Aggregate revenue by brand
revenue_ty = filtered_data_ty.groupby('Brand')['Revenue USD'].sum().reset_index()
revenue_ly = filtered_data_ly.groupby('Brand')['Revenue USD'].sum().reset_index()

# Debugging: Print aggregated revenues
print("TY revenue by brand:\n", revenue_ty)
print("LY revenue by brand:\n", revenue_ly)

# Merge TY and LY dataframes
revenue = revenue_ty.merge(revenue_ly, on='Brand', suffixes=('_TY', '_LY'), how='outer').fillna(0)

# Calculate vs LY
revenue['vs LY'] = revenue['Revenue USD_TY'] - revenue['Revenue USD_LY']

# Rename columns
revenue.rename(columns={'Revenue USD_LY': 'LY', 'Revenue USD_TY': 'TY'}, inplace=True)

# Save the result to a CSV file
output_path = 'C:/Users/Aliya/Downloads/Q2_Revenue_Comparison.csv'
revenue.to_csv(output_path, index=False)

print(f"Revenue comparison for Q2 saved to CSV at {output_path}.")


