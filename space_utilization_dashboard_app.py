import io
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# ------------------------------------------------------------
# Page setup
# ------------------------------------------------------------
st.set_page_config(
    page_title="Plant Space Utilization Dashboard",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .main {background-color: #f7f9fc;}
        .block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
        .hero-card {
            background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 52%, #2563eb 100%);
            padding: 24px 28px;
            border-radius: 22px;
            color: white;
            box-shadow: 0 10px 28px rgba(15, 23, 42, 0.25);
            margin-bottom: 18px;
        }
        .hero-title {font-size: 34px; font-weight: 800; margin: 0;}
        .hero-subtitle {font-size: 15px; opacity: 0.92; margin-top: 8px;}
        .metric-card {
            background: white;
            padding: 18px 20px;
            border-radius: 18px;
            border: 1px solid #e5e7eb;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.07);
        }
        .metric-label {font-size: 13px; color: #64748b; font-weight: 700;}
        .metric-value {font-size: 25px; color: #0f172a; font-weight: 850; margin-top: 4px;}
        .section-title {font-size: 21px; font-weight: 800; color: #0f172a; margin-top: 14px;}
        .small-note {font-size: 12px; color: #64748b;}
        div[data-testid="stDataFrame"] {background: white; border-radius: 16px;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ------------------------------------------------------------
# Helper functions
# ------------------------------------------------------------
def excel_engine_for_file(file_name: str):
    suffix = Path(file_name).suffix.lower()
    if suffix == ".xlsb":
        return "pyxlsb"
    if suffix in [".xlsx", ".xlsm"]:
        return "openpyxl"
    return None


@st.cache_data(show_spinner=False)
def get_sheet_names(uploaded_file_bytes: bytes, file_name: str):
    engine = excel_engine_for_file(file_name)
    bio = io.BytesIO(uploaded_file_bytes)
    xl = pd.ExcelFile(bio, engine=engine)
    return xl.sheet_names


@st.cache_data(show_spinner=False)
def load_excel_data(uploaded_file_bytes: bytes, file_name: str, sheet_name: str):
    engine = excel_engine_for_file(file_name)
    bio = io.BytesIO(uploaded_file_bytes)
    df = pd.read_excel(bio, sheet_name=sheet_name, engine=engine)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def format_number(value, decimals=0):
    try:
        if pd.isna(value):
            return "0"
        return f"{value:,.{decimals}f}"
    except Exception:
        return str(value)


def prepare_monthly_space(df, month_col, week_col, plant_col, area_col):
    work = df.copy()

    work[month_col] = pd.to_numeric(work[month_col], errors="coerce")
    work[week_col] = pd.to_numeric(work[week_col], errors="coerce")
    work[area_col] = pd.to_numeric(work[area_col], errors="coerce").fillna(0)
    work[plant_col] = work[plant_col].astype(str).str.strip()

    work = work.dropna(subset=[month_col, week_col])
    work = work[work[plant_col].notna() & (work[plant_col] != "") & (work[plant_col].str.lower() != "nan")]

    month_names = {
        1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
        7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec",
    }

    weekly = (
        work.groupby([month_col, week_col, plant_col], dropna=False)[area_col]
        .sum()
        .reset_index()
        .rename(columns={area_col: "Weekly Utilized Space"})
    )

    monthly = (
        weekly.groupby([month_col, plant_col], dropna=False)
        .agg(
            Total_Utilized_Space=("Weekly Utilized Space", "sum"),
            No_of_Weeks=(week_col, "nunique"),
            Highest_Week_Space=("Weekly Utilized Space", "max"),
            Lowest_Week_Space=("Weekly Utilized Space", "min"),
        )
        .reset_index()
    )
    monthly["Average_Monthly_Space"] = monthly["Total_Utilized_Space"] / monthly["No_of_Weeks"].replace(0, pd.NA)
    monthly["Month_Name"] = monthly[month_col].astype(int).map(month_names).fillna(monthly[month_col].astype(str))

    total_monthly = (
        weekly.groupby([month_col, week_col], dropna=False)["Weekly Utilized Space"]
        .sum()
        .reset_index()
        .groupby(month_col, dropna=False)
        .agg(
            Total_Utilized_Space=("Weekly Utilized Space", "sum"),
            No_of_Weeks=(week_col, "nunique"),
            Highest_Week_Space=("Weekly Utilized Space", "max"),
            Lowest_Week_Space=("Weekly Utilized Space", "min"),
        )
        .reset_index()
    )
    total_monthly["Average_Monthly_Space"] = total_monthly["Total_Utilized_Space"] / total_monthly["No_of_Weeks"].replace(0, pd.NA)
    total_monthly["Month_Name"] = total_monthly[month_col].astype(int).map(month_names).fillna(total_monthly[month_col].astype(str))

    return work, weekly, monthly, total_monthly


def to_excel_bytes(monthly_df, weekly_df, filtered_raw_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        monthly_df.to_excel(writer, index=False, sheet_name="Monthly Plant Average")
        weekly_df.to_excel(writer, index=False, sheet_name="Weekly Plant Summary")
        filtered_raw_df.to_excel(writer, index=False, sheet_name="Filtered Raw Data")

        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "font_color": "white", "bg_color": "#1e3a8a", "border": 1})
        num_fmt = workbook.add_format({"num_format": "#,##0.00", "border": 1})
        text_fmt = workbook.add_format({"border": 1})

        for sheet_name in ["Monthly Plant Average", "Weekly Plant Summary", "Filtered Raw Data"]:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.set_row(0, 22, header_fmt)
            worksheet.set_column(0, 12, 18, text_fmt)
            worksheet.set_column(2, 8, 20, num_fmt)
    output.seek(0)
    return output.getvalue()


# ------------------------------------------------------------
# Header
# ------------------------------------------------------------
st.markdown(
    """
    <div class="hero-card">
        <div class="hero-title">🏭 Plant Space Utilization Control Tower</div>
        <div class="hero-subtitle">
            Upload your weekly stock occupancy file and view month-wise average utilized space by plant.
            Formula used: <b>Monthly Average = Sum of weekly utilized space / Number of weeks available in that month</b>.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# ------------------------------------------------------------
# Sidebar upload and mapping
# ------------------------------------------------------------
with st.sidebar:
    st.header("📁 Upload File")
    uploaded_file = st.file_uploader(
        "Upload Stock Occupancy Excel file",
        type=["xlsb", "xlsx", "xlsm"],
        help="Your file should contain the DATA sheet with Month, Week, Plant, and volumetric area columns.",
    )

    st.markdown("---")
    st.caption("Default mapping expected from your file:")
    st.caption("Month = Column B | Week = Column C | Plant = Plant | Area = Column AP / volumetric area")

if uploaded_file is None:
    st.info("Please upload your Stock Occupancy file to start the dashboard.")
    st.stop()

uploaded_bytes = uploaded_file.getvalue()

try:
    sheets = get_sheet_names(uploaded_bytes, uploaded_file.name)
except Exception as e:
    st.error(f"Unable to read this Excel file. Error: {e}")
    st.stop()

with st.sidebar:
    sheet_name = st.selectbox("Select data sheet", sheets, index=sheets.index("DATA") if "DATA" in sheets else 0)

try:
    df = load_excel_data(uploaded_bytes, uploaded_file.name, sheet_name)
except Exception as e:
    st.error(f"Unable to load sheet '{sheet_name}'. Error: {e}")
    st.stop()

if df.empty:
    st.warning("The selected sheet is blank.")
    st.stop()

columns = list(df.columns)

def default_col(name, fallback_index=None):
    for c in columns:
        if c.lower() == name.lower():
            return c
    if fallback_index is not None and fallback_index < len(columns):
        return columns[fallback_index]
    return columns[0]

with st.sidebar:
    st.header("🧭 Column Mapping")
    month_col = st.selectbox("Month column", columns, index=columns.index(default_col("Month", 1)))
    week_col = st.selectbox("Week column", columns, index=columns.index(default_col("Week", 2)))
    plant_col = st.selectbox("Plant column", columns, index=columns.index(default_col("Plant", 6)))
    area_col = st.selectbox("Volumetric area column", columns, index=columns.index(default_col("volumetric area", 41 if len(columns) > 41 else None)))

try:
    clean_df, weekly_df, monthly_df, total_monthly_df = prepare_monthly_space(df, month_col, week_col, plant_col, area_col)
except Exception as e:
    st.error(f"Unable to calculate dashboard. Please check column mapping. Error: {e}")
    st.stop()

# ------------------------------------------------------------
# Filters
# ------------------------------------------------------------
st.markdown('<div class="section-title">🔎 Filters</div>', unsafe_allow_html=True)
f1, f2, f3 = st.columns([1.1, 1.5, 1])

all_months = sorted(monthly_df[month_col].dropna().unique())
all_plants = sorted(monthly_df[plant_col].dropna().astype(str).unique())

with f1:
    selected_months = st.multiselect("Month", all_months, default=all_months)
with f2:
    selected_plants = st.multiselect("Plant", all_plants, default=all_plants)
with f3:
    chart_mode = st.radio("Chart value", ["Average Monthly Space", "Total Utilized Space"], horizontal=False)

filtered_monthly = monthly_df[
    monthly_df[month_col].isin(selected_months) & monthly_df[plant_col].astype(str).isin(selected_plants)
].copy()
filtered_weekly = weekly_df[
    weekly_df[month_col].isin(selected_months) & weekly_df[plant_col].astype(str).isin(selected_plants)
].copy()
filtered_raw = clean_df[
    clean_df[month_col].isin(selected_months) & clean_df[plant_col].astype(str).isin(selected_plants)
].copy()

if filtered_monthly.empty:
    st.warning("No records found for the selected filters.")
    st.stop()

# ------------------------------------------------------------
# KPI cards
# ------------------------------------------------------------
total_area = filtered_monthly["Total_Utilized_Space"].sum()
avg_area = filtered_monthly["Average_Monthly_Space"].mean()
plant_count = filtered_monthly[plant_col].nunique()
week_count = filtered_weekly[week_col].nunique()
record_count = len(filtered_raw)

k1, k2, k3, k4, k5 = st.columns(5)
metric_data = [
    ("Total Utilized Space", format_number(total_area, 2)),
    ("Avg. Monthly Space", format_number(avg_area, 2)),
    ("Plants", format_number(plant_count, 0)),
    ("Weeks", format_number(week_count, 0)),
    ("Records", format_number(record_count, 0)),
]
for col, (label, value) in zip([k1, k2, k3, k4, k5], metric_data):
    with col:
        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value">{value}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

# ------------------------------------------------------------
# Charts
# ------------------------------------------------------------
st.markdown('<div class="section-title">📊 Monthly Plant-wise Space Analysis</div>', unsafe_allow_html=True)

chart_col = "Average_Monthly_Space" if chart_mode == "Average Monthly Space" else "Total_Utilized_Space"
y_title = "Average Monthly Space" if chart_col == "Average_Monthly_Space" else "Total Utilized Space"

c1, c2 = st.columns([1.35, 1])
with c1:
    fig_bar = px.bar(
        filtered_monthly.sort_values([month_col, plant_col]),
        x="Month_Name",
        y=chart_col,
        color=plant_col,
        barmode="group",
        text_auto=".2s",
        title=f"Plant-wise {y_title} by Month",
        labels={"Month_Name": "Month", chart_col: y_title, plant_col: "Plant"},
    )
    fig_bar.update_layout(height=470, title_font_size=18, legend_title_text="Plant", margin=dict(l=10, r=10, t=55, b=10))
    st.plotly_chart(fig_bar, use_container_width=True)

with c2:
    plant_rank = (
        filtered_monthly.groupby(plant_col)["Average_Monthly_Space"]
        .mean()
        .reset_index()
        .sort_values("Average_Monthly_Space", ascending=False)
    )
    fig_rank = px.bar(
        plant_rank,
        x="Average_Monthly_Space",
        y=plant_col,
        orientation="h",
        text_auto=".2s",
        title="Average Space Ranking by Plant",
        labels={"Average_Monthly_Space": "Avg. Space", plant_col: "Plant"},
    )
    fig_rank.update_layout(height=470, yaxis={"categoryorder": "total ascending"}, margin=dict(l=10, r=10, t=55, b=10))
    st.plotly_chart(fig_rank, use_container_width=True)

c3, c4 = st.columns([1.2, 1])
with c3:
    trend_df = (
        filtered_monthly.groupby([month_col, "Month_Name"], as_index=False)["Average_Monthly_Space"]
        .sum()
        .sort_values(month_col)
    )
    fig_line = px.line(
        trend_df,
        x="Month_Name",
        y="Average_Monthly_Space",
        markers=True,
        title="Total Average Space Trend Month-wise",
        labels={"Month_Name": "Month", "Average_Monthly_Space": "Average Space"},
    )
    fig_line.update_traces(line_width=4, marker_size=9)
    fig_line.update_layout(height=420, margin=dict(l=10, r=10, t=55, b=10))
    st.plotly_chart(fig_line, use_container_width=True)

with c4:
    latest_month = max(selected_months) if selected_months else filtered_monthly[month_col].max()
    latest_df = filtered_monthly[filtered_monthly[month_col] == latest_month]
    donut_df = latest_df.groupby(plant_col, as_index=False)["Average_Monthly_Space"].sum()
    fig_donut = px.pie(
        donut_df,
        names=plant_col,
        values="Average_Monthly_Space",
        hole=0.45,
        title=f"Plant Share in Month {latest_month}",
    )
    fig_donut.update_layout(height=420, margin=dict(l=10, r=10, t=55, b=10))
    st.plotly_chart(fig_donut, use_container_width=True)

# ------------------------------------------------------------
# Detailed tables and export
# ------------------------------------------------------------
st.markdown('<div class="section-title">📋 Monthly Average Calculation Table</div>', unsafe_allow_html=True)
show_cols = [month_col, "Month_Name", plant_col, "Total_Utilized_Space", "No_of_Weeks", "Average_Monthly_Space", "Highest_Week_Space", "Lowest_Week_Space"]
view_df = filtered_monthly[show_cols].sort_values([month_col, plant_col]).copy()
st.dataframe(
    view_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Total_Utilized_Space": st.column_config.NumberColumn("Total Utilized Space", format="%.2f"),
        "Average_Monthly_Space": st.column_config.NumberColumn("Average Monthly Space", format="%.2f"),
        "Highest_Week_Space": st.column_config.NumberColumn("Highest Week Space", format="%.2f"),
        "Lowest_Week_Space": st.column_config.NumberColumn("Lowest Week Space", format="%.2f"),
    },
)

with st.expander("View weekly plant-wise summary"):
    st.dataframe(filtered_weekly.sort_values([month_col, week_col, plant_col]), use_container_width=True, hide_index=True)

export_bytes = to_excel_bytes(view_df, filtered_weekly, filtered_raw)
st.download_button(
    label="⬇️ Download Filtered Dashboard Data in Excel",
    data=export_bytes,
    file_name="Plant_Space_Utilization_Dashboard_Output.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.markdown(
    """
    <div class="small-note">
        Calculation logic: first the app sums volumetric area at Month + Week + Plant level, then calculates Month + Plant average as monthly total divided by the number of distinct weeks in that month.
    </div>
    """,
    unsafe_allow_html=True,
)
