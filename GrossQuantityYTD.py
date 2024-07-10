import pandas as pd
from datetime import datetime

# Load the data with specified dtypes to avoid mixed type warning
gross_sales = pd.read_csv('C:/Users/Aliya/Downloads/GrossSales_transformed (1).csv', dtype={'Country Code': str, 'Brand': str}, low_memory=False)
currencies = pd.read_excel('C:/Users/Aliya/Downloads/Currencies.xlsx')
bcodes = pd.read_csv('C:/Users/Aliya/Downloads/BCodes.csv')

# Define the date range for 2nd April to 30th June for 2024 and the same period in 2023
start_date_2024 = datetime(2024, 4, 2)
end_date_2024 = datetime(2024, 6, 30)
start_date_2023 = datetime(2023, 4, 2)
end_date_2023 = datetime(2023, 6, 30)

# Function to filter sales data based on the SFCC Order Date column
def filter_sales(gross_sales, start_date, end_date):
    gross_sales['SFCC Order Date'] = pd.to_datetime(gross_sales['SFCC Order Date'], format='%d/%m/%Y', errors='coerce')
    return gross_sales[(gross_sales['SFCC Order Date'] >= start_date) & (gross_sales['SFCC Order Date'] <= end_date)]

# Filter offline sales data for the specified date ranges
filtered_sales_2024 = filter_sales(gross_sales, start_date_2024, end_date_2024)
filtered_sales_2023 = filter_sales(gross_sales, start_date_2023, end_date_2023)

# Print filtered sales data for TY and LY
print("Filtered Sales Data for 2nd April - 30th Jun 2024 (TY):\n", filtered_sales_2024)
print("Filtered Sales Data for 2nd April - 30th Jun 2023 (LY):\n", filtered_sales_2023)

# Remove negative values
filtered_sales_2024 = filtered_sales_2024[filtered_sales_2024['Quantity'] >= 0]
filtered_sales_2023 = filtered_sales_2023[filtered_sales_2023['Quantity'] >= 0]

# Prepare the result DataFrame
result = pd.DataFrame(columns=['Brand', 'TY Quantity', 'LY Quantity', 'vs LY'])

# Get unique brands from TY and LY sales data
brands_ty = filtered_sales_2024['Brand'].unique()
brands_ly = filtered_sales_2023['Brand'].unique()
all_brands = set(brands_ty).union(set(brands_ly))

print("Unique Brands for TY:\n", brands_ty)
print("Unique Brands for LY:\n", brands_ly)
print("All Unique Brands:\n", all_brands)

# Calculate the total quantity for each brand for TY and LY
for brand in all_brands:
    total_quantity_ty = 0
    total_quantity_ly = 0

    # Calculate TY quantity
    if brand in brands_ty:
        print(f"Processing TY for brand: {brand}")
        brand_sales_ty = filtered_sales_2024[filtered_sales_2024['Brand'] == brand]
        total_quantity_ty = brand_sales_ty['Quantity'].sum()
        print(f"Total quantity TY for brand {brand}: {total_quantity_ty}")

    # Calculate LY quantity
    if brand in brands_ly:
        print(f"Processing LY for brand: {brand}")
        brand_sales_ly = filtered_sales_2023[filtered_sales_2023['Brand'] == brand]
        total_quantity_ly = brand_sales_ly['Quantity'].sum()
        print(f"Total quantity LY for brand {brand}: {total_quantity_ly}")

    # Calculate vs LY
    if total_quantity_ly > 0:
        vs_ly = ((total_quantity_ty - total_quantity_ly) / total_quantity_ly) * 100
    else:
        vs_ly = float('inf') if total_quantity_ty > 0 else 0

    print(f"Total Quantity for TY for brand {brand}: {total_quantity_ty}")
    print(f"Total Quantity for LY for brand {brand}: {total_quantity_ly}")
    print(f"vs LY for brand {brand}: {vs_ly}")

    # Append result
    result = pd.concat([
        result,
        pd.DataFrame({
            'Brand': [brand],
            'TY Quantity': [total_quantity_ty],
            'LY Quantity': [total_quantity_ly],
            'vs LY': [f"{vs_ly:.2f}%"]
        })
    ], ignore_index=True)

# Replace brand codes with brand names
result = result.merge(bcodes, left_on='Brand', right_on='Code')
result = result.drop(columns=['Brand', 'Code'])
result = result.rename(columns={'Name': 'Brand'})

# Reorder columns to place Brand first
result = result[['Brand', 'TY Quantity', 'LY Quantity', 'vs LY']]

# Print the result DataFrame
print("Result DataFrame:\n", result)

# Save the result to a CSV file and Excel file
result.to_csv('C:/Users/Aliya/Downloads/yearly_gross_quantity.csv', index=False)
