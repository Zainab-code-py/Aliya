import pandas as pd
from datetime import datetime, timedelta

# Load the data with specified dtypes to avoid mixed type warning
views_data = pd.read_csv('C:/Users/Aliya/Downloads/sku_azadea.csv', dtype={'Item ID': str, 'Items viewed': int}, low_memory=False)
bcodes = pd.read_csv('C:/Users/Aliya/Downloads/BCodes.csv')
calendar = pd.read_csv('C:/Users/Aliya/Downloads/Calendarx.csv')

# Define the start and end dates for the quarter
quarter_start = datetime(2024, 4, 2)
quarter_end = datetime(2024, 6, 30)

# Convert Date to datetime and filter views data for the quarter
views_data['Date'] = pd.to_datetime(views_data['Date'], errors='coerce')
filtered_views_quarter = views_data[
    (views_data['Date'] >= quarter_start) &
    (views_data['Date'] <= quarter_end)
    ].copy()

# Prepare the result DataFrame
result = pd.DataFrame(columns=['Brand', 'Total Views'])

# Get unique brands from the quarter views data
brands_quarter = filtered_views_quarter['Item ID'].apply(lambda x: x.split('_')[0]).unique()

# Calculate the total number of items viewed for each brand for the quarter
for brand in brands_quarter:
    total_views_quarter = 0

    # Calculate quarter total views
    if brand in brands_quarter:
        brand_views_quarter = filtered_views_quarter[filtered_views_quarter['Item ID'].apply(lambda x: x.split('_')[0]) == brand]
        total_views_quarter = brand_views_quarter['Items viewed'].sum()

    # Append result
    result = pd.concat([
        result,
        pd.DataFrame({
            'Brand': [brand],
            'Total Views': [total_views_quarter]
        })
    ], ignore_index=True)

# Replace brand codes with brand names
result = result.merge(bcodes, left_on='Brand', right_on='Code')
result = result.drop(columns=['Brand', 'Code'])
result = result.rename(columns={'Name': 'Brand'})

# Reorder columns to place Brand first
result = result[['Brand', 'Total Views']]

# Print the result DataFrame
print(result)

# Save the result to the WTD.xlsx file
with pd.ExcelWriter('C:/Users/Aliya/Downloads/WTD.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    result.to_excel(writer, sheet_name='Total Views Q2 2024', index=False)
