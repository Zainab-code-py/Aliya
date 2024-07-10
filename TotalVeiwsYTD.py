import pandas as pd
from datetime import datetime, timedelta

# Load the data with specified dtypes to avoid mixed type warning
views_data = pd.read_csv('C:/Users/Aliya/Downloads/sku_azadea.csv', dtype={'Item ID': str, 'Items viewed': int}, low_memory=False)
bcodes = pd.read_csv('C:/Users/Aliya/Downloads/BCodes.csv')
calendar = pd.read_csv('C:/Users/Aliya/Downloads/Calendarx.csv')

# Define the start and end dates for YTD (This Year and Last Year)
today = datetime.today()
start_of_year = datetime(today.year, 1, 1)
start_of_last_year = datetime(today.year - 1, 1, 1)
end_of_last_year = datetime(today.year - 1, today.month, today.day)

# Convert Date to datetime and filter views data for YTD and Last YTD
views_data['Date'] = pd.to_datetime(views_data['Date'], errors='coerce')
filtered_views_ytd = views_data[
    (views_data['Date'] >= start_of_year) &
    (views_data['Date'] <= today)
    ].copy()
filtered_views_last_ytd = views_data[
    (views_data['Date'] >= start_of_last_year) &
    (views_data['Date'] <= end_of_last_year)
    ].copy()

# Ensure 'Item ID' is treated as a string
filtered_views_ytd['Item ID'] = filtered_views_ytd['Item ID'].astype(str)
filtered_views_last_ytd['Item ID'] = filtered_views_last_ytd['Item ID'].astype(str)

# Prepare the result DataFrame
result = pd.DataFrame(columns=['Brand', 'YTD', 'Last YTD', 'vs Last YTD'])

# Get unique brands from YTD and Last YTD views data
brands_ytd = filtered_views_ytd['Item ID'].apply(lambda x: x.split('_')[0]).unique()
brands_last_ytd = filtered_views_last_ytd['Item ID'].apply(lambda x: x.split('_')[0]).unique()
all_brands = set(brands_ytd).union(set(brands_last_ytd))

# Calculate the total number of items viewed for each brand for YTD and Last YTD
for brand in all_brands:
    total_views_ytd = 0
    total_views_last_ytd = 0

    # Calculate YTD total views
    if brand in brands_ytd:
        brand_views_ytd = filtered_views_ytd[filtered_views_ytd['Item ID'].apply(lambda x: x.split('_')[0]) == brand]
        total_views_ytd = brand_views_ytd['Items viewed'].sum()

    # Calculate Last YTD total views
    if brand in brands_last_ytd:
        brand_views_last_ytd = filtered_views_last_ytd[filtered_views_last_ytd['Item ID'].apply(lambda x: x.split('_')[0]) == brand]
        total_views_last_ytd = brand_views_last_ytd['Items viewed'].sum()

    # Calculate vs Last YTD
    if total_views_last_ytd > 0:
        vs_last_ytd = ((total_views_ytd - total_views_last_ytd) / total_views_last_ytd) * 100
    else:
        vs_last_ytd = float('inf') if total_views_ytd > 0 else 0

    # Append result
    result = pd.concat([
        result,
        pd.DataFrame({
            'Brand': [brand],
            'YTD': [total_views_ytd],
            'Last YTD': [total_views_last_ytd],
            'vs Last YTD': [f"{round(vs_last_ytd, 2)}%"]  # Round vs Last YTD to 2 decimal places and add %
        })
    ], ignore_index=True)

# Replace brand codes with brand names
result = result.merge(bcodes, left_on='Brand', right_on='Code')
result = result.drop(columns=['Brand', 'Code'])
result = result.rename(columns={'Name': 'Brand'})

# Reorder columns to place Brand first
result = result[['Brand', 'YTD', 'Last YTD', 'vs Last YTD']]

# Print the result DataFrame
print(result)

# Save the result to the WTD.xlsx file
with pd.ExcelWriter('C:/Users/Aliya/Downloads/WTD.xlsx', mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    result.to_excel(writer, sheet_name='Total Views YTD', index=False)
