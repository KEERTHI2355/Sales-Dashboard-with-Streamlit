import pandas as pd
import numpy as np

# Set seed for reproducibility
np.random.seed(42)

# Number of rows
n = 5000

# Generate columns
dates = pd.date_range(start="2023-01-01", end="2023-12-31", periods=n)
products = ["Laptop", "Headphones", "Keyboard", "Mouse", "Monitor", "Webcam"]
regions = ["North", "South", "East", "West"]
categories = ["Electronics", "Accessories"]

df = pd.DataFrame({
    "Order ID": np.arange(10001, 10001 + n),
    "Order Date": dates,
    "Product": np.random.choice(products, n),
    "Category": np.random.choice(categories, n),
    "Region": np.random.choice(regions, n),
    "Quantity": np.random.randint(1, 6, n),
    "Price": np.random.randint(1000, 60000, n)
})

# Calculate Sales
df["Sales"] = df["Quantity"] * df["Price"]

# Save dataset
df.to_csv("sales_dataset.csv", index=False)

print("Dataset created successfully! File saved as: sales_dataset.csv")
