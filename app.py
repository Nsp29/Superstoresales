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

# ---- Sidebar Filters ----
st.sidebar.title("Filters")

# Region Filter
all_regions = sorted(df_original["Region"].dropna().unique())
selected_region = st.sidebar.selectbox("Select Region", options=["All"] + all_regions)

# Filter data by Region
if selected_region != "All":
    df_filtered_region = df_original[df_original["Region"] == selected_region]
else:
    df_filtered_region = df_original

# State Filter
all_states = sorted(df_filtered_region["State"].dropna().unique())
selected_state = st.sidebar.selectbox("Select State", options=["All"] + all_states)

# Filter data by State
if selected_state != "All":
    df_filtered_state = df_filtered_region[df_filtered_region["State"] == selected_state]
else:
    df_filtered_state = df_filtered_region

# Category Filter
all_categories = sorted(df_filtered_state["Category"].dropna().unique())
selected_category = st.sidebar.selectbox("Select Category", options=["All"] + all_categories)

# Filter data by Category
if selected_category != "All":
    df_filtered_category = df_filtered_state[df_filtered_state["Category"] == selected_category]
else:
    df_filtered_category = df_filtered_state

# Sub-Category Filter
all_subcats = sorted(df_filtered_category["Sub-Category"].dropna().unique())
selected_subcat = st.sidebar.selectbox("Select Sub-Category", options=["All"] + all_subcats)

# Final filter by Sub-Category
df = df_filtered_category.copy()
if selected_subcat != "All":
    df = df[df["Sub-Category"] == selected_subcat]

# ---- Sidebar Time Period Selection ----
st.sidebar.subheader("Select Time Period")
time_period_options = ["1M", "6M", "12M", "YTD"]
selected_time_period = st.sidebar.selectbox("Select Period", time_period_options, key="time_period_filter")

# ---- Determine Date Range Based on Selected Period ----
today = pd.to_datetime("today")

if selected_time_period == "1M":
    from_date = today - pd.DateOffset(months=1)
elif selected_time_period == "6M":
    from_date = today - pd.DateOffset(months=6)
elif selected_time_period == "12M":
    from_date = today - pd.DateOffset(years=1)
elif selected_time_period == "YTD":  # Year-To-Date
    from_date = pd.Timestamp(year=today.year, month=1, day=1)  # Start of the year

to_date = today  # Always set `to_date` to today

# Apply date range filter
df = df_original[
    (df_original["Order Date"] >= from_date) &
    (df_original["Order Date"] <= to_date)
]
# ---- Compute KPI values at Start and End Dates ----
df_start = df[df["Order Date"] == df["Order Date"].min()]  # Data at the start date
df_end = df[df["Order Date"] == df["Order Date"].max()]  # Data at the end date

# Calculate values at the start and end of the selected period
start_sales = df_start["Sales"].sum()
end_sales = df_end["Sales"].sum()

start_quantity = df_start["Quantity"].sum()
end_quantity = df_end["Quantity"].sum()

start_profit = df_start["Profit"].sum()
end_profit = df_end["Profit"].sum()

start_margin = (df_start["Profit"].sum() / df_start["Sales"].sum()) * 100 if df_start["Sales"].sum() != 0 else 0
end_margin = (df_end["Profit"].sum() / df_end["Sales"].sum()) * 100 if df_end["Sales"].sum() != 0 else 0

# ---- Compute Percentage Growth from Start to End Date ----
sales_growth = ((end_sales - start_sales) / start_sales) * 100 if start_sales != 0 else 0
quantity_growth = ((end_quantity - start_quantity) / start_quantity) * 100 if start_quantity != 0 else 0
profit_growth = ((end_profit - start_profit) / start_profit) * 100 if start_profit != 0 else 0
margin_growth = ((end_margin - start_margin) / start_margin) * 100 if start_margin != 0 else 0

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


# Function to format KPI with growth percentage and arrow
def format_kpi(value, growth):
    arrow = "▲" if growth > 0 else "▼"
    color = "green" if growth > 0 else "red"
    return f"""
        <div class='kpi-value' style='color:#1E90FF; font-size:24px; font-weight:700;'>${value:,.2f}</div>
        <div style='color:{color}; font-size:14px; font-weight:600;'> {arrow} {abs(growth):.1f}%</div>
    """
# ---- KPI Display (Rectangles) ----
kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)

with kpi_col1:
    st.markdown(f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Total Sales</div>
            {format_kpi(end_sales, sales_growth)}
        </div>
    """, unsafe_allow_html=True)

with kpi_col2:
    st.markdown(f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Quantity Sold</div>
            {format_kpi(end_quantity, quantity_growth)}
        </div>
    """, unsafe_allow_html=True)

with kpi_col3:
    st.markdown(f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Total Profit</div>
            {format_kpi(end_profit, profit_growth)}
        </div>
    """, unsafe_allow_html=True)

with kpi_col4:
    st.markdown(f"""
        <div class='kpi-box'>
            <div class='kpi-title'>Profit Margin (%)</div>
            {format_kpi(end_margin, margin_growth)}
        </div>
    """, unsafe_allow_html=True)

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
