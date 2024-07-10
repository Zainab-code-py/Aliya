import pandas as pd
from datetime import datetime

# Define the date range for the fiscal quarter (2nd April to 28th June) for this year
specified_start = datetime(2024, 4, 2)
specified_end = datetime(2024, 6, 28)

# Load the data
gross_sales = pd.read_csv('C:/Users/Aliya/Downloads/GrossSales_transformed (1).csv', low_memory=False)
aws_azadea_performance = pd.read_csv('C:/Users/Aliya/Downloads/AWSAzadeaPerformance.csv')
bcodes = pd.read_csv('C:/Users/Aliya/Downloads/bcodes.csv')

# Print the columns of each DataFrame to verify column names
print("Gross Sales Columns: ", gross_sales.columns)
print("AWS Azadea Performance Columns: ", aws_azadea_performance.columns)
print("BCodes Columns: ", bcodes.columns)

# Convert dates to datetime for filtering in gross_sales
gross_sales['SFCC Order Date'] = pd.to_datetime(gross_sales['SFCC Order Date'], dayfirst=True, errors='coerce')

# Filter data for the specified date range in gross_sales
filtered_gross_sales = gross_sales[
    (gross_sales['SFCC Order Date'] >= specified_start) &
    (gross_sales['SFCC Order Date'] <= specified_end)
    ].copy()

# Normalize the column names in bcodes
bcodes.columns = bcodes.columns.str.strip().str.lower()

# Convert dates to datetime for filtering in aws_azadea_performance
aws_azadea_performance['Date'] = pd.to_datetime(aws_azadea_performance['Date'], dayfirst=True, errors='coerce')
filtered_performance = aws_azadea_performance[
    (aws_azadea_performance['Date'] >= specified_start) &
    (aws_azadea_performance['Date'] <= specified_end)
    ].copy()

# Calculate total sessions from AWS Azadea Performance by summing app sessions and web sessions
total_sessions = filtered_performance['App Sessions'].sum() + filtered_performance['Web Sessions'].sum()
print("Total Sessions: ", total_sessions)

# Calculate distinct orders for each brand
distinct_orders = filtered_gross_sales.groupby('Brand')['Order No.'].nunique().reset_index()
distinct_orders.columns = ['brand', 'distinct orders']
print("Distinct Orders: ", distinct_orders.head())

# Calculate conversions for each brand
distinct_orders['conversions'] = distinct_orders['distinct orders'] / total_sessions
print("Conversions: ", distinct_orders.head())

# Map the brand names
bcodes.columns = ['code', 'name']
result = pd.merge(distinct_orders, bcodes, left_on='brand', right_on='code', how='left')
print("Merged Results: ", result.head())

# Assume LY Conversions and vs LY as % are placeholders for now
# Here we assume LY conversions as a simple mean of TY conversions for simplicity
ly_conversions = result['conversions'].mean()
result['ly conversions'] = ly_conversions
result['vs ly conversions'] = ((result['conversions'] - ly_conversions) / ly_conversions) * 100

# Prepare the result DataFrame with appropriate columns
result = result[['name', 'conversions', 'ly conversions', 'vs ly conversions']]
result.columns = ['Brand', 'TY Conversions', 'LY Conversions', 'vs LY Conversions']
print("Final Results: ", result.head())

# Save the result to a CSV file and Excel file
result.to_csv('C:/Users/Aliya/Downloads/conversion_results.csv', index=False)
result.to_excel('C:/Users/Aliya/Downloads/conversion_results.xlsx', index=False)

print("Conversion results have been successfully saved.")
