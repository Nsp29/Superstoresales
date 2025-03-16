import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# Set page config for wide layout
st.set_page_config(page_title="SuperStore KPI Dashboard", layout="wide")

# ---- Load Data ----
@st.cache_data
def load_data():
    # Adjust the path if needed, e.g. "data/Sample - Superstore.xlsx"
    df = pd.read_excel("Sample - Superstore.xlsx", engine="openpyxl")
    # Convert Order Date to datetime if not already
    if not pd.api.types.is_datetime64_any_dtype(df["Order Date"]):
        df["Order Date"] = pd.to_datetime(df["Order Date"])
    return df

df_original = load_data()

# Sub-Category Filter
all_subcats = sorted(df_filtered_category["Sub-Category"].dropna().unique())
selected_subcat = st.sidebar.selectbox("Select Sub-Category", options=["All"] + all_subcats)

# Final filter by Sub-Category
df = df_filtered_category.copy()
if selected_subcat != "All":
    df = df[df["Sub-Category"] == selected_subcat]

# ---- Sidebar Date Range (From and To) ----
if df.empty:
    # If there's no data after filters, default to overall min/max
    min_date = df_original["Order Date"].min()
    max_date = df_original["Order Date"].max()
else:
    min_date = df["Order Date"].min()
    max_date = df["Order Date"].max()

from_date = st.sidebar.date_input(
    "From Date", value=min_date, min_value=min_date, max_value=max_date
)
to_date = st.sidebar.date_input(
    "To Date", value=max_date, min_value=min_date, max_value=max_date
)

# Ensure from_date <= to_date
if from_date > to_date:
    st.sidebar.error("From Date must be earlier than To Date.")

# Apply date range filter
df = df[
    (df["Order Date"] >= pd.to_datetime(from_date))
    & (df["Order Date"] <= pd.to_datetime(to_date))
]
# ---- Page Title ----
st.title("SuperStore KPI Dashboard")

# ---- Custom CSS for KPI Tiles ----
st.markdown(
    """
    <style>
    .kpi-box {
        background-color: #ffffff;
        border: 2px solid #EAEAEA;
        border-radius: 8px;
        padding: 18px;
        margin: 8px;
        text-align: center;
    }
    .kpi-title {
        font-weight: 500;
        color: #333333;
        font-size: 12px;
        margin-bottom: 8px;
    }
    .kpi-value {
        font-weight: 700;
        font-size: 24px;
        color: #1E90FF;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ---- KPI Calculation ----
if df.empty:
    total_sales = 0
    total_quantity = 0
    total_profit = 0
    margin_rate = 0
else:
    total_sales = df["Sales"].sum()
    total_quantity = df["Quantity"].sum()
    total_profit = df["Profit"].sum()
    margin_rate = (total_profit / total_sales) if total_sales != 0 else 0

# ---- KPI Display (Rectangles) ----
kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
with kpi_col1:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Sales</div>
            <div class='kpi-value'>${total_sales:,.2f}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col2:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Quantity Sold</div>
            <div class='kpi-value'>{total_quantity:,.0f}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
# Determine Profitability Color (Green for profit, Red for loss)
profit_color = "green" if total_profit >= 0 else "red"
with kpi_col3:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Profit</div>
            <div class='kpi-value' style='color:{profit_color};'>${total_profit:,.2f}</div>
        </div>
        """,
        unsafe_allow_html=True
    )
with kpi_col4:
    st.markdown(
        f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Margin Rate</div>
            <div class='kpi-value'>{(margin_rate * 100):,.2f}%</div>
        </div>
        """,
        unsafe_allow_html=True
    )

# ---- KPI Selection & Visualizations ----
st.subheader("Visualize KPI Across Time & Top Products")

# Ensure `selected_kpi` is defined even if data is empty
kpi_options = ["Sales", "Quantity", "Profit", "Margin Rate"]
selected_kpi = st.radio("Select KPI to display:", options=kpi_options, horizontal=True)

if df.empty:
    st.warning("No data available for the selected filters and date range.")
else:
    # ---- Prepare Data for Charts ----
    daily_grouped = df.groupby("Order Date").agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum"
    }).reset_index()
    daily_grouped["Margin Rate"] = daily_grouped["Profit"] / daily_grouped["Sales"].replace(0, 1)

    product_grouped = df.groupby("Product Name").agg({
        "Sales": "sum",
        "Quantity": "sum",
        "Profit": "sum"
    }).reset_index()
    product_grouped["Margin Rate"] = product_grouped["Profit"] / product_grouped["Sales"].replace(0, 1)
    product_grouped.sort_values(by=selected_kpi, ascending=False, inplace=True)
    top_10 = product_grouped.head(10)

    # Create a mapping of full state names to abbreviations
state_abbreviation_map = {
    'Alabama': 'AL', 'Alaska': 'AK', 'Arizona': 'AZ', 'Arkansas': 'AR', 'California': 'CA',
    'Colorado': 'CO', 'Connecticut': 'CT', 'Delaware': 'DE', 'District of Columbia': 'DC', 'Florida': 'FL',
    'Georgia': 'GA', 'Hawaii': 'HI', 'Idaho': 'ID', 'Illinois': 'IL', 'Indiana': 'IN', 'Iowa': 'IA',
    'Kansas': 'KS', 'Kentucky': 'KY', 'Louisiana': 'LA', 'Maine': 'ME', 'Maryland': 'MD', 'Massachusetts': 'MA',
    'Michigan': 'MI', 'Minnesota': 'MN', 'Mississippi': 'MS', 'Missouri': 'MO', 'Montana': 'MT',
    'Nebraska': 'NE', 'Nevada': 'NV', 'New Hampshire': 'NH', 'New Jersey': 'NJ', 'New Mexico': 'NM',
    'New York': 'NY', 'North Carolina': 'NC', 'North Dakota': 'ND', 'Ohio': 'OH', 'Oklahoma': 'OK',
    'Oregon': 'OR', 'Pennsylvania': 'PA', 'Rhode Island': 'RI', 'South Carolina': 'SC', 'South Dakota': 'SD',
    'Tennessee': 'TN', 'Texas': 'TX', 'Utah': 'UT', 'Vermont': 'VT', 'Virginia': 'VA', 'Washington': 'WA',
    'West Virginia': 'WV', 'Wisconsin': 'WI', 'Wyoming': 'WY'
}

# Map full state names to abbreviations
df["State"] = df["State"].map(state_abbreviation_map)
# ---- Three-column Layout ----
col1, col2, col3 = st.columns(3)
with col1:
# Time Series Line Chart
    fig_line = px.line(
        daily_grouped,
        x="Order Date",
        y=selected_kpi,
        title=f"{selected_kpi} Over Time",
        labels={"Order Date": "Date", selected_kpi: selected_kpi},
        template="plotly_white",
        )
    fig_line.update_layout(height=400)
    st.plotly_chart(fig_line, use_container_width=True)

with col2:
# Top 10 Products Bar Chart
    fig_bar = px.bar(
        top_10,
        x=selected_kpi,
        y="Product Name",
        title=f"Top 10 Products by {selected_kpi}",
        labels={selected_kpi: selected_kpi, "Product Name": "Product"},
        color=selected_kpi,
        color_continuous_scale="Blues",
        template="plotly_white",
        )
    fig_bar.update_layout(
        height=400,
        yaxis={"categoryorder": "total ascending"}
        )
    st.plotly_chart(fig_bar, use_container_width=True)

with col3:

    # Aggregate profit by city
    city_profit = df.groupby("City").agg({"Profit": "sum"}).reset_index()

    # Sort by profit and take the top 10
    top_10_cities = city_profit.sort_values(by="Profit", ascending=False).head(10)

    # Create a horizontal bar chart for top cities
    fig_city_bar = px.bar(
        top_10_cities,
        x="Profit",
        y="City",
        title="Top 10 Profitable Cities",
        labels={"Profit": "Total Profit ($)", "City": "City"},
        color="Profit",
        color_continuous_scale="Blues",
        template="plotly_white",
    )
    # Improve layout
    fig_city_bar.update_layout(
        height=400,
        yaxis=dict(categoryorder="total ascending"),  # Order by profit
    )

    # Display the chart
    st.plotly_chart(fig_city_bar, use_container_width=True)
