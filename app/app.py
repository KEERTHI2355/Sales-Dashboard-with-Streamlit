import streamlit as st
import pandas as pd
import numpy as np
import io
from pathlib import Path

# CONFIG
st.set_page_config(
    page_title="Sales Insights Dashboard",
    layout="wide"
)

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_PATH = BASE_DIR / "data" / "sales_cleaned.csv"


# LOAD DATA
@st.cache_data
def load_data(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path, parse_dates=["Order Date"])
    if "order_month" not in df.columns:
        df["order_month"] = df["Order Date"].dt.to_period("M").astype(str) # type: ignore
    return df

df = load_data(DATA_PATH)


# EXCEL EXPORT FUNCTION
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="raw_data", index=False)

        df.groupby("Product")["Sales"].sum().reset_index() \
            .to_excel(writer, sheet_name="product_sales", index=False)

        df.groupby("Category")["Sales"].sum().reset_index() \
            .to_excel(writer, sheet_name="category_sales", index=False)

        df.groupby("Region")["Sales"].sum().reset_index() \
            .to_excel(writer, sheet_name="region_sales", index=False)

    return buffer.getvalue()


# SIDEBAR FILTERS
st.sidebar.header("Filters")

min_date = df["Order Date"].min().date()
max_date = df["Order Date"].max().date()

date_range = st.sidebar.date_input(
    "Order Date Range",
    [min_date, max_date],
    min_value=min_date,
    max_value=max_date
)

regions = st.sidebar.multiselect(
    "Region",
    options=sorted(df["Region"].unique()),
    default=list(df["Region"].unique())
)

categories = st.sidebar.multiselect(
    "Category",
    options=sorted(df["Category"].unique()),
    default=list(df["Category"].unique())
)


# APPLY FILTERS
start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1]) # type: ignore

mask = (
    (df["Order Date"] >= start) &
    (df["Order Date"] <= end) &
    (df["Region"].isin(regions)) &
    (df["Category"].isin(categories))
)

df_f = df.loc[mask].copy()


# KPI CALCULATIONS
total_revenue = df_f["Sales"].sum()
total_orders = df_f["Order ID"].nunique()
avg_order_value = total_revenue / total_orders if total_orders > 0 else 0

monthly_sales = df_f.groupby("order_month")["Sales"].sum().sort_index()

mom_growth = np.nan
if len(monthly_sales) >= 2 and monthly_sales.iloc[-2] != 0:
    mom_growth = (
        (monthly_sales.iloc[-1] - monthly_sales.iloc[-2]) /
        monthly_sales.iloc[-2] * 100
    )


# MAIN UI
st.title("📊 Sales Insights Dashboard")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Revenue", f"₹{total_revenue:,.0f}")
c2.metric("Total Orders", f"{total_orders:,}")
c3.metric("Average Order Value", f"₹{avg_order_value:,.0f}")
c4.metric(
    "MoM Revenue Growth",
    f"{mom_growth:.2f}%" if not np.isnan(mom_growth) else "N/A"
)

st.divider()


# TOP PRODUCTS
st.subheader("Top Products by Sales")

top_products = (
    df_f.groupby("Product")["Sales"]
    .sum()
    .sort_values(ascending=False)
    .head(5)
    .reset_index()
)

st.dataframe(
    top_products.assign(
        Sales=top_products["Sales"].apply(lambda x: f"₹{x:,.0f}")
    ),
    use_container_width=True
)


# CHARTS
st.subheader("Monthly Revenue Trend")
st.line_chart(monthly_sales)

st.subheader("Sales by Region")
region_sales = df_f.groupby("Region")["Sales"].sum()
st.bar_chart(region_sales)


# EXPORT SECTION
st.sidebar.markdown("---")
st.sidebar.subheader("Export")

excel_data = to_excel_bytes(df_f)

st.sidebar.download_button(
    label="Download Excel Report",
    data=excel_data,
    file_name="sales_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
