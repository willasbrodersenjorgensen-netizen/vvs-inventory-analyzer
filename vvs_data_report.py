import pandas as pd

# 1. Loading the Excel file
try:
    df = pd.read_excel("vvs_inventory.xlsx")
except FileNotFoundError:
    print("Error: 'vvs_inventory.xlsx' not found. Please run your generator script first.")
    exit()

# 2. Cleaning the data
df = df.drop_duplicates()

print("--- INITIAL DATA PREVIEW ---")
print(df.head(), "\n")

# 3. What Materials are running low
low_stock = df[df['Stock_Qty'] < 5]
print("--- LOW QUANTITY STOCK ITEMS ---")
print(low_stock[['Material_Name', 'Stock_Qty']], "\n")


# 4. Calculating Revenue for each item and total revenue for all items
df['Revenue'] = df['Stock_Qty'] * df['Price']
total_revenue = df['Revenue'].sum()

# 5. What has high Price
high_value = df[df['Revenue'] > df['Revenue'].mean()]
print("HIGH VALUE STOCK ITEMS")
print(high_value[['Material_Name', 'Stock_Qty', 'Price', 'Revenue']], "\n")

# Summary for items and revenue
print("SUMMARY")
print(df[['Material_Name', 'Stock_Qty', 'Price', 'Revenue']].head())
print(f"\nWhole Warehouse Revenue (Total Value): {total_revenue:.2f}")

df.to_excel('vvs_report.xlsx', index=False)
print("Report saved to vvs_report.xlsx")