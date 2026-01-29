import pandas as pd

df = pd.read_csv("data/sales_cleaned.csv", parse_dates=["Order Date"])

# ---- KPI Calculations ----
total_sales = df["Sales"].sum()
total_orders = df["Order ID"].nunique()
average_order_value = round(total_sales / total_orders, 2)

best_product = df.groupby("Product")["Sales"].sum().idxmax()
best_region = df.groupby("Region")["Sales"].sum().idxmax()

# ---- Print KPI Cards ----
print("\n===== KPI Dashboard =====")
print(f"💰 Total Revenue: ₹{total_sales:,.0f}")
print(f"📦 Total Orders: {total_orders:,}")
print(f"📊 Average Order Value (AOV): ₹{average_order_value:,.2f}")
print(f"🏆 Best-Selling Product: {best_product}")
print(f"🌍 Top-Performing Region: {best_region}")
print("=========================\n")
