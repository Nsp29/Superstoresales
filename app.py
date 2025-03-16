import streamlit as st
import pandas as pd
import plotly.express as px

# Load Data
file_path = "Sample - Superstore.xlsx"
xls = pd.ExcelFile(file_path)

# Read sheets
orders_df = pd.read_excel(xls, sheet_name="Orders")
returns_df = pd.read_excel(xls, sheet_name="Returns")

# Merge Returns with Orders
orders_df["Returned"] = orders_df["Order ID"].isin(returns_df["Order ID"])

# Convert Dates
orders_df["Order Date"] = pd.to_datetime(orders_df["Order Date"])
orders_df["Ship Date"] = pd.to_datetime(orders_df["Ship Date"])
orders_df["Delivery Days"] = (orders_df["Ship Date"] - orders_df["Order Date"]).dt.days

# Sidebar Filters
st.sidebar.header("Filters")
category = st.sidebar.multiselect("Select Category", orders_df["Category"].unique(), default=orders_df["Category"].unique())
region = st.sidebar.multiselect("Select Region", orders_df["Region"].unique(), default=orders_df["Region"].unique())
customer_segment = st.sidebar.multiselect("Select Customer Segment", orders_df["Segment"].unique(), default=orders_df["Segment"].unique())

# Filter Data
filtered_df = orders_df[
    (orders_df["Category"].isin(category)) & 
    (orders_df["Region"].isin(region)) & 
    (orders_df["Segment"].isin(customer_segment))
]

# KPI Metrics
total_sales = filtered_df["Sales"].sum()
total_profit = filtered_df["Profit"].sum()
total_orders = filtered_df["Order ID"].nunique()
return_rate = (filtered_df["Returned"].sum() / total_orders) * 100 if total_orders > 0 else 0

# Display KPIs
st.title("📊 Superstore Business Dashboard")
st.metric("Total Sales", f"${total_sales:,.2f}")
st.metric("Total Profit", f"${total_profit:,.2f}")
st.metric("Total Orders", total_orders)
st.metric("Return Rate", f"{return_rate:.2f}%")

# Sales & Profit Over Time
st.subheader("📈 Sales & Profit Over Time")
fig = px.line(filtered_df, x="Order Date", y=["Sales", "Profit"], title="Sales & Profit Trends", labels={"value": "Amount", "variable": "Metric"})
st.plotly_chart(fig)

# Sales by Region
st.subheader("🌎 Sales & Profit by Region")
fig = px.bar(filtered_df.groupby("Region")[["Sales", "Profit"]].sum().reset_index(),
             x="Region", y=["Sales", "Profit"], barmode="group",
             title="Regional Sales & Profit")
st.plotly_chart(fig)

# Discount Impact on Profit
st.subheader("💰 Discount vs. Profitability")
fig = px.scatter(filtered_df, x="Discount", y="Profit", color="Category", 
                 title="Impact of Discounts on Profitability")
st.plotly_chart(fig)

# Top 10 Products by Sales & Profit
st.subheader("🏆 Top 10 Products by Sales")
top_products = filtered_df.groupby("Product Name")[["Sales", "Profit"]].sum().reset_index().nlargest(10, "Sales")
fig = px.bar(top_products, x="Sales", y="Product Name", orientation="h", title="Top 10 Products by Sales")
st.plotly_chart(fig)

st.subheader("⚠️ Top 10 Least Profitable Products")
bottom_products = filtered_df.groupby("Product Name")[["Sales", "Profit"]].sum().reset_index().nsmallest(10, "Profit")
fig = px.bar(bottom_products, x="Profit", y="Product Name", orientation="h", title="Top 10 Least Profitable Products")
st.plotly_chart(fig)

# Shipping Performance
st.subheader("🚚 Shipping Performance")
fig = px.box(filtered_df, x="Ship Mode", y="Delivery Days", title="Delivery Time by Shipping Mode")
st.plotly_chart(fig)

# Return Analysis
st.subheader("🔄 Returns by Category")
returns_data = filtered_df[filtered_df["Returned"]].groupby("Category")["Returned"].count().reset_index()
fig = px.pie(returns_data, names="Category", values="Returned", title="Return Rate by Category")
st.plotly_chart(fig)

# Footer
st.markdown("---")
st.caption("📊 Dashboard built with Streamlit & Plotly | Data: Superstore Dataset")
