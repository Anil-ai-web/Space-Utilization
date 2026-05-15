import io
from pathlib import Path

import pandas as pd
import plotly.express as px
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

        .hero-title {
            font-size: 34px;
            font-weight: 800;
            margin: 0;
        }

        .hero-subtitle {
            font-size: 15px;
            opacity: 0.92;
            margin-top: 8px;
            line-height: 1.6;
        }

        .metric-card {
            background: white;
            padding: 18px 20px;
            border-radius: 18px;
            border: 1px solid #e5e7eb;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.07);
            min-height: 105px;
        }

        .metric-label {
            font-size: 13px;
            color: #64748b;
            font-weight: 700;
        }

        .metric-value {
            font-size: 24px;
            color: #0f172a;
            font-weight: 850;
            margin-top: 6px;
        }

        .section-title {
            font-size: 21px;
            font-weight: 800;
            color: #0f172a;
            margin-top: 14px;
            margin-bottom: 8px;
        }

        .small-note {
            font-size: 12px;
            color: #64748b;
            margin-top: 8px;
        }

        .formula-box {
            background: #eff6ff;
            border: 1px solid #bfdbfe;
            border-radius: 16px;
            padding: 14px 18px;
            color: #0f172a;
            margin-bottom: 16px;
            font-size: 14px;
            line-height: 1.7;
        }

        div[data-testid="stDataFrame"] {
            background: white;
            border-radius: 16px;
        }
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

    df = pd.read_excel(
        bio,
        sheet_name=sheet_name,
        engine=engine,
    )

    df.columns = [str(c).strip() for c in df.columns]
    return df


def format_number(value, decimals=0):
    try:
        if pd.isna(value):
            return "0"
        return f"{value:,.{decimals}f}"
    except Exception:
        return str(value)


def find_default_column(columns, possible_names=None, fallback_index=None):
    possible_names = possible_names or []

    lower_map = {str(c).strip().lower(): c for c in columns}

    for name in possible_names:
        if str(name).strip().lower() in lower_map:
            return lower_map[str(name).strip().lower()]

    if fallback_index is not None and fallback_index < len(columns):
        return columns[fallback_index]

    return columns[0]


def safe_to_datetime(series):
    """
    Converts normal Excel dates, text dates, and Excel serial dates into pandas datetime.
    """
    dt = pd.to_datetime(series, errors="coerce", dayfirst=True)

    if dt.notna().sum() < max(1, len(series) * 0.30):
        numeric_series = pd.to_numeric(series, errors="coerce")
        serial_dt = pd.to_datetime(
            numeric_series,
            errors="coerce",
            unit="D",
            origin="1899-12-30",
        )

        if serial_dt.notna().sum() > dt.notna().sum():
            dt = serial_dt

    return dt


def prepare_monthly_space(
    df,
    month_col,
    date_col,
    plant_col,
    area_col,
    stacking_height,
):
    work = df.copy()

    work[month_col] = pd.to_numeric(work[month_col], errors="coerce")
    work[date_col] = safe_to_datetime(work[date_col])
    work[area_col] = pd.to_numeric(work[area_col], errors="coerce").fillna(0)
    work[plant_col] = work[plant_col].astype(str).str.strip()

    work = work.dropna(subset=[month_col, date_col])

    work = work[
        work[plant_col].notna()
        & (work[plant_col] != "")
        & (work[plant_col].str.lower() != "nan")
    ]

    work["Date_Only"] = work[date_col].dt.date

    month_names = {
        1: "Jan",
        2: "Feb",
        3: "Mar",
        4: "Apr",
        5: "May",
        6: "Jun",
        7: "Jul",
        8: "Aug",
        9: "Sep",
        10: "Oct",
        11: "Nov",
        12: "Dec",
    }

    # --------------------------------------------------------
    # Step 1:
    # Date-wise total volumetric area for each Month + Date + Plant
    # --------------------------------------------------------
    daily = (
        work.groupby([month_col, "Date_Only", plant_col], dropna=False)[area_col]
        .sum()
        .reset_index()
        .rename(columns={area_col: "Daily_Total_Volumetric_Area"})
    )

    # --------------------------------------------------------
    # Step 2:
    # Divide total volumetric area by stacking height.
    # Default stacking height = 5 feet.
    # --------------------------------------------------------
    daily["Daily_Utilized_Floor_Space"] = (
        daily["Daily_Total_Volumetric_Area"] / stacking_height
    )

    # --------------------------------------------------------
    # Step 3:
    # Monthly plant average before aisle space.
    # Base Monthly Average Space =
    # Sum of daily utilized floor space / Number of dates
    # --------------------------------------------------------
    monthly = (
        daily.groupby([month_col, plant_col], dropna=False)
        .agg(
            Total_Volumetric_Area=("Daily_Total_Volumetric_Area", "sum"),
            Total_Utilized_Floor_Space_Before_Aisle=(
                "Daily_Utilized_Floor_Space",
                "sum",
            ),
            No_of_Dates=("Date_Only", "nunique"),
            Highest_Date_Space_Before_Aisle=("Daily_Utilized_Floor_Space", "max"),
            Lowest_Date_Space_Before_Aisle=("Daily_Utilized_Floor_Space", "min"),
        )
        .reset_index()
    )

    monthly["Base_Average_Monthly_Space"] = (
        monthly["Total_Utilized_Floor_Space_Before_Aisle"]
        / monthly["No_of_Dates"].replace(0, pd.NA)
    )

    # --------------------------------------------------------
    # Step 4:
    # Add aisle space:
    # Plant I070 = 25%
    # All other plants = 35%
    # --------------------------------------------------------
    monthly["Aisle_Percentage"] = monthly[plant_col].astype(str).str.upper().eq("I070")
    monthly["Aisle_Percentage"] = monthly["Aisle_Percentage"].map(
        {
            True: 0.25,
            False: 0.35,
        }
    )

    monthly["Aisle_Percentage_Display"] = monthly["Aisle_Percentage"] * 100

    monthly["Aisle_Space"] = (
        monthly["Base_Average_Monthly_Space"] * monthly["Aisle_Percentage"]
    )

    monthly["Average_Monthly_Space_With_Aisle"] = (
        monthly["Base_Average_Monthly_Space"] + monthly["Aisle_Space"]
    )

    monthly["Total_Utilized_Floor_Space_With_Aisle"] = (
        monthly["Average_Monthly_Space_With_Aisle"] * monthly["No_of_Dates"]
    )

    monthly["Highest_Date_Space_With_Aisle"] = (
        monthly["Highest_Date_Space_Before_Aisle"]
        * (1 + monthly["Aisle_Percentage"])
    )

    monthly["Lowest_Date_Space_With_Aisle"] = (
        monthly["Lowest_Date_Space_Before_Aisle"]
        * (1 + monthly["Aisle_Percentage"])
    )

    monthly["Month_Name"] = (
        monthly[month_col]
        .astype(int)
        .map(month_names)
        .fillna(monthly[month_col].astype(str))
    )

    return work, daily, monthly


def to_excel_bytes(monthly_df, daily_df, filtered_raw_df):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        monthly_df.to_excel(writer, index=False, sheet_name="Monthly Plant Average")
        daily_df.to_excel(writer, index=False, sheet_name="Daily Plant Summary")
        filtered_raw_df.to_excel(writer, index=False, sheet_name="Filtered Raw Data")

        workbook = writer.book

        header_fmt = workbook.add_format(
            {
                "bold": True,
                "font_color": "white",
                "bg_color": "#1e3a8a",
                "border": 1,
            }
        )

        num_fmt = workbook.add_format(
            {
                "num_format": "#,##0.00",
                "border": 1,
            }
        )

        text_fmt = workbook.add_format(
            {
                "border": 1,
            }
        )

        for sheet_name in [
            "Monthly Plant Average",
            "Daily Plant Summary",
            "Filtered Raw Data",
        ]:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.set_row(0, 22, header_fmt)
            worksheet.set_column(0, 18, 22, text_fmt)
            worksheet.set_column(3, 16, 24, num_fmt)

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
            Upload your stock occupancy file and view month-wise average utilized floor space by plant.
            This dashboard calculates monthly average based on actual stock dates and adds plant-wise aisle space.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="formula-box">
        <b>Final Calculation Logic:</b><br>
        1. Daily Total Volumetric Area = Sum of Volumetric Area column, usually Column AP, by Plant + Month + Date<br>
        2. Daily Utilized Floor Space = Daily Total Volumetric Area ÷ Stacking Height<br>
        3. Base Monthly Average Space = Sum of Daily Utilized Floor Space ÷ Number of Dates in that month<br>
        4. Aisle Space Addition = I070: 25%, All other plants: 35%<br>
        5. Final Monthly Average Space = Base Monthly Average Space + Aisle Space
    </div>
    """,
    unsafe_allow_html=True,
)


# ------------------------------------------------------------
# Sidebar upload
# ------------------------------------------------------------
with st.sidebar:
    st.header("📁 Upload File")

    uploaded_file = st.file_uploader(
        "Upload Stock Occupancy Excel file",
        type=["xlsb", "xlsx", "xlsm"],
        help="Your file should contain DATA sheet with Month, Date, Plant, and Volumetric Area columns.",
    )

    st.markdown("---")

    st.header("📐 Stacking Height")

    stacking_height = st.number_input(
        "Stacking height in feet",
        min_value=1.0,
        max_value=50.0,
        value=5.0,
        step=0.5,
        help="Default is 5 feet. Total volumetric area will be divided by this height.",
    )

    st.markdown("---")
    st.header("🚧 Aisle Space Rule")
    st.caption("Plant I070 = 25%")
    st.caption("All other plants = 35%")

    st.markdown("---")
    st.caption("Default expected mapping:")
    st.caption("Month = Column B")
    st.caption("Date = Select actual stock/date column")
    st.caption("Volumetric Area = Column AP")
    st.caption("Plant = Select from dropdown")


if uploaded_file is None:
    st.info("Please upload your Stock Occupancy file to start the dashboard.")
    st.stop()


uploaded_bytes = uploaded_file.getvalue()


# ------------------------------------------------------------
# Load sheet
# ------------------------------------------------------------
try:
    sheets = get_sheet_names(uploaded_bytes, uploaded_file.name)
except Exception as e:
    st.error(f"Unable to read this Excel file. Error: {e}")
    st.stop()


with st.sidebar:
    default_sheet_index = sheets.index("DATA") if "DATA" in sheets else 0

    sheet_name = st.selectbox(
        "Select data sheet",
        sheets,
        index=default_sheet_index,
    )


try:
    df = load_excel_data(uploaded_bytes, uploaded_file.name, sheet_name)
except Exception as e:
    st.error(f"Unable to load sheet '{sheet_name}'. Error: {e}")
    st.stop()


if df.empty:
    st.warning("The selected sheet is blank.")
    st.stop()


columns = list(df.columns)


# ------------------------------------------------------------
# Column mapping
# ------------------------------------------------------------
default_month_col = find_default_column(
    columns,
    possible_names=[
        "Month",
        "month",
        "MONTH",
    ],
    fallback_index=1,  # Column B
)

default_date_col = find_default_column(
    columns,
    possible_names=[
        "Date",
        "date",
        "DATE",
        "Stock Date",
        "stock date",
        "Stock_Date",
        "Created on",
        "created on",
        "Posting Date",
        "posting date",
        "Document Date",
        "document date",
        "As on Date",
        "as on date",
    ],
    fallback_index=0,
)

default_area_col = find_default_column(
    columns,
    possible_names=[
        "volumetric area",
        "Volumetric Area",
        "Volumetric_Area",
        "VOL Area",
        "VOL AREA",
        "AP",
    ],
    fallback_index=41 if len(columns) > 41 else None,  # Column AP
)

default_plant_col = find_default_column(
    columns,
    possible_names=[
        "Plant",
        "plant",
        "PLANT",
        "Location",
        "location",
        "Godown",
        "Warehouse",
        "warehouse",
        "Plant Name",
        "plant name",
    ],
    fallback_index=6 if len(columns) > 6 else None,
)


with st.sidebar:
    st.header("🧭 Column Mapping")

    month_col = st.selectbox(
        "Month column",
        columns,
        index=columns.index(default_month_col),
    )

    date_col = st.selectbox(
        "Date column",
        columns,
        index=columns.index(default_date_col),
        help="Select the actual stock date column. Monthly average will divide by number of unique dates.",
    )

    plant_col = st.selectbox(
        "Plant column",
        columns,
        index=columns.index(default_plant_col),
    )

    area_col = st.selectbox(
        "Volumetric area column",
        columns,
        index=columns.index(default_area_col),
    )


# ------------------------------------------------------------
# Prepare calculation
# ------------------------------------------------------------
try:
    clean_df, daily_df, monthly_df = prepare_monthly_space(
        df=df,
        month_col=month_col,
        date_col=date_col,
        plant_col=plant_col,
        area_col=area_col,
        stacking_height=stacking_height,
    )
except Exception as e:
    st.error(f"Unable to calculate dashboard. Please check column mapping. Error: {e}")
    st.stop()


if daily_df.empty or monthly_df.empty:
    st.warning("No valid data found after applying Month, Date, Plant, and Volumetric Area logic.")
    st.stop()


# ------------------------------------------------------------
# Filters
# ------------------------------------------------------------
st.markdown('<div class="section-title">🔎 Filters</div>', unsafe_allow_html=True)

f1, f2, f3 = st.columns([1.1, 1.5, 1])

all_months = sorted(monthly_df[month_col].dropna().unique())
all_plants = sorted(monthly_df[plant_col].dropna().astype(str).unique())

with f1:
    selected_months = st.multiselect(
        "Month",
        all_months,
        default=all_months,
    )

with f2:
    selected_plants = st.multiselect(
        "Plant",
        all_plants,
        default=all_plants,
    )

with f3:
    chart_mode = st.radio(
        "Chart value",
        [
            "Final Monthly Space With Aisle",
            "Base Monthly Space Before Aisle",
            "Aisle Space",
            "Total Volumetric Area",
        ],
        horizontal=False,
    )


filtered_monthly = monthly_df[
    monthly_df[month_col].isin(selected_months)
    & monthly_df[plant_col].astype(str).isin(selected_plants)
].copy()

filtered_daily = daily_df[
    daily_df[month_col].isin(selected_months)
    & daily_df[plant_col].astype(str).isin(selected_plants)
].copy()

filtered_raw = clean_df[
    clean_df[month_col].isin(selected_months)
    & clean_df[plant_col].astype(str).isin(selected_plants)
].copy()


if filtered_monthly.empty:
    st.warning("No records found for the selected filters.")
    st.stop()


# ------------------------------------------------------------
# KPI cards
# ------------------------------------------------------------
total_volumetric_area = filtered_monthly["Total_Volumetric_Area"].sum()
total_floor_before_aisle = filtered_monthly[
    "Total_Utilized_Floor_Space_Before_Aisle"
].sum()
avg_base_space = filtered_monthly["Base_Average_Monthly_Space"].mean()
avg_final_space = filtered_monthly["Average_Monthly_Space_With_Aisle"].mean()
total_aisle_space = filtered_monthly["Aisle_Space"].sum()
plant_count = filtered_monthly[plant_col].nunique()
date_count = filtered_daily["Date_Only"].nunique()
record_count = len(filtered_raw)

k1, k2, k3, k4, k5, k6 = st.columns(6)

metric_data = [
    ("Total Volumetric Area", format_number(total_volumetric_area, 2)),
    ("Floor Space Before Aisle", format_number(total_floor_before_aisle, 2)),
    ("Avg. Before Aisle", format_number(avg_base_space, 2)),
    ("Avg. With Aisle", format_number(avg_final_space, 2)),
    ("Aisle Space Added", format_number(total_aisle_space, 2)),
    ("Dates", format_number(date_count, 0)),
]

for col, (label, value) in zip([k1, k2, k3, k4, k5, k6], metric_data):
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

k7, k8 = st.columns(2)

with k7:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">Plants</div>
            <div class="metric-value">{format_number(plant_count, 0)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

with k8:
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">Records</div>
            <div class="metric-value">{format_number(record_count, 0)}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ------------------------------------------------------------
# Chart column selection
# ------------------------------------------------------------
if chart_mode == "Final Monthly Space With Aisle":
    chart_col = "Average_Monthly_Space_With_Aisle"
    y_title = "Final Monthly Space With Aisle"
elif chart_mode == "Base Monthly Space Before Aisle":
    chart_col = "Base_Average_Monthly_Space"
    y_title = "Base Monthly Space Before Aisle"
elif chart_mode == "Aisle Space":
    chart_col = "Aisle_Space"
    y_title = "Aisle Space Added"
else:
    chart_col = "Total_Volumetric_Area"
    y_title = "Total Volumetric Area"


# ------------------------------------------------------------
# Charts
# ------------------------------------------------------------
st.markdown(
    '<div class="section-title">📊 Monthly Plant-wise Space Analysis</div>',
    unsafe_allow_html=True,
)

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
        labels={
            "Month_Name": "Month",
            chart_col: y_title,
            plant_col: "Plant",
        },
    )

    fig_bar.update_layout(
        height=470,
        title_font_size=18,
        legend_title_text="Plant",
        margin=dict(l=10, r=10, t=55, b=10),
    )

    st.plotly_chart(fig_bar, use_container_width=True)


with c2:
    plant_rank = (
        filtered_monthly.groupby(plant_col)["Average_Monthly_Space_With_Aisle"]
        .mean()
        .reset_index()
        .sort_values("Average_Monthly_Space_With_Aisle", ascending=False)
    )

    fig_rank = px.bar(
        plant_rank,
        x="Average_Monthly_Space_With_Aisle",
        y=plant_col,
        orientation="h",
        text_auto=".2s",
        title="Final Average Space Ranking by Plant",
        labels={
            "Average_Monthly_Space_With_Aisle": "Avg. Space With Aisle",
            plant_col: "Plant",
        },
    )

    fig_rank.update_layout(
        height=470,
        yaxis={"categoryorder": "total ascending"},
        margin=dict(l=10, r=10, t=55, b=10),
    )

    st.plotly_chart(fig_rank, use_container_width=True)


c3, c4 = st.columns([1.2, 1])

with c3:
    trend_df = (
        filtered_monthly.groupby([month_col, "Month_Name"], as_index=False)[
            "Average_Monthly_Space_With_Aisle"
        ]
        .sum()
        .sort_values(month_col)
    )

    fig_line = px.line(
        trend_df,
        x="Month_Name",
        y="Average_Monthly_Space_With_Aisle",
        markers=True,
        title="Final Average Space Trend Month-wise",
        labels={
            "Month_Name": "Month",
            "Average_Monthly_Space_With_Aisle": "Final Average Space With Aisle",
        },
    )

    fig_line.update_traces(line_width=4, marker_size=9)

    fig_line.update_layout(
        height=420,
        margin=dict(l=10, r=10, t=55, b=10),
    )

    st.plotly_chart(fig_line, use_container_width=True)


with c4:
    latest_month = max(selected_months) if selected_months else filtered_monthly[month_col].max()
    latest_df = filtered_monthly[filtered_monthly[month_col] == latest_month]

    donut_df = latest_df.groupby(plant_col, as_index=False)[
        "Average_Monthly_Space_With_Aisle"
    ].sum()

    fig_donut = px.pie(
        donut_df,
        names=plant_col,
        values="Average_Monthly_Space_With_Aisle",
        hole=0.45,
        title=f"Plant Share in Month {latest_month} After Aisle",
    )

    fig_donut.update_layout(
        height=420,
        margin=dict(l=10, r=10, t=55, b=10),
    )

    st.plotly_chart(fig_donut, use_container_width=True)


# ------------------------------------------------------------
# Daily trend chart
# ------------------------------------------------------------
st.markdown(
    '<div class="section-title">📈 Daily Floor Space Movement Before Aisle</div>',
    unsafe_allow_html=True,
)

daily_trend = filtered_daily.copy()
daily_trend["Date_Label"] = pd.to_datetime(daily_trend["Date_Only"]).dt.strftime("%d-%b-%Y")

fig_daily = px.line(
    daily_trend.sort_values(["Date_Only", plant_col]),
    x="Date_Label",
    y="Daily_Utilized_Floor_Space",
    color=plant_col,
    markers=True,
    title="Daily Utilized Floor Space by Plant Before Aisle",
    labels={
        "Date_Label": "Date",
        "Daily_Utilized_Floor_Space": "Daily Floor Space Before Aisle",
        plant_col: "Plant",
    },
)

fig_daily.update_layout(
    height=430,
    margin=dict(l=10, r=10, t=55, b=10),
)

st.plotly_chart(fig_daily, use_container_width=True)


# ------------------------------------------------------------
# Detailed table
# ------------------------------------------------------------
st.markdown(
    '<div class="section-title">📋 Monthly Average Calculation Table</div>',
    unsafe_allow_html=True,
)

show_cols = [
    month_col,
    "Month_Name",
    plant_col,
    "Total_Volumetric_Area",
    "Total_Utilized_Floor_Space_Before_Aisle",
    "No_of_Dates",
    "Base_Average_Monthly_Space",
    "Aisle_Percentage_Display",
    "Aisle_Space",
    "Average_Monthly_Space_With_Aisle",
    "Total_Utilized_Floor_Space_With_Aisle",
    "Highest_Date_Space_With_Aisle",
    "Lowest_Date_Space_With_Aisle",
]

view_df = filtered_monthly[show_cols].sort_values([month_col, plant_col]).copy()

st.dataframe(
    view_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Total_Volumetric_Area": st.column_config.NumberColumn(
            "Total Volumetric Area",
            format="%.2f",
        ),
        "Total_Utilized_Floor_Space_Before_Aisle": st.column_config.NumberColumn(
            "Total Floor Space Before Aisle",
            format="%.2f",
        ),
        "No_of_Dates": st.column_config.NumberColumn(
            "No. of Dates",
            format="%d",
        ),
        "Base_Average_Monthly_Space": st.column_config.NumberColumn(
            "Base Average Monthly Space",
            format="%.2f",
        ),
        "Aisle_Percentage_Display": st.column_config.NumberColumn(
            "Aisle %",
            format="%.2f",
        ),
        "Aisle_Space": st.column_config.NumberColumn(
            "Aisle Space Added",
            format="%.2f",
        ),
        "Average_Monthly_Space_With_Aisle": st.column_config.NumberColumn(
            "Final Average Monthly Space With Aisle",
            format="%.2f",
        ),
        "Total_Utilized_Floor_Space_With_Aisle": st.column_config.NumberColumn(
            "Total Floor Space With Aisle",
            format="%.2f",
        ),
        "Highest_Date_Space_With_Aisle": st.column_config.NumberColumn(
            "Highest Date Space With Aisle",
            format="%.2f",
        ),
        "Lowest_Date_Space_With_Aisle": st.column_config.NumberColumn(
            "Lowest Date Space With Aisle",
            format="%.2f",
        ),
    },
)


# ------------------------------------------------------------
# Daily summary and raw data
# ------------------------------------------------------------
with st.expander("View daily plant-wise summary before aisle"):
    st.dataframe(
        filtered_daily.sort_values([month_col, "Date_Only", plant_col]),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Daily_Total_Volumetric_Area": st.column_config.NumberColumn(
                "Daily Total Volumetric Area",
                format="%.2f",
            ),
            "Daily_Utilized_Floor_Space": st.column_config.NumberColumn(
                "Daily Utilized Floor Space Before Aisle",
                format="%.2f",
            ),
        },
    )


with st.expander("View filtered raw data"):
    st.dataframe(
        filtered_raw,
        use_container_width=True,
        hide_index=True,
    )


# ------------------------------------------------------------
# Download output
# ------------------------------------------------------------
export_bytes = to_excel_bytes(
    monthly_df=view_df,
    daily_df=filtered_daily,
    filtered_raw_df=filtered_raw,
)

st.download_button(
    label="⬇️ Download Filtered Dashboard Data in Excel",
    data=export_bytes,
    file_name="Plant_Space_Utilization_Dashboard_Output.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


st.markdown(
    f"""
    <div class="small-note">
        Calculation logic used in this dashboard:<br>
        1. The app first sums volumetric area at <b>Month + Date + Plant</b> level.<br>
        2. Then it divides the total volumetric area by <b>{stacking_height} feet stacking height</b>
        to calculate utilized floor space.<br>
        3. Base Month + Plant average is calculated as total utilized floor space divided by
        the <b>number of unique dates</b> available in that month.<br>
        4. Then aisle space is added: <b>I070 = 25%</b>, and <b>all other plants = 35%</b>.<br>
        5. Final Monthly Average Space = Base Average Monthly Space + Aisle Space.
    </div>
    """,
    unsafe_allow_html=True,
)
