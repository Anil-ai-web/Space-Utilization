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
        }

        .metric-label {
            font-size: 13px;
            color: #64748b;
            font-weight: 700;
        }

        .metric-value {
            font-size: 25px;
            color: #0f172a;
            font-weight: 850;
            margin-top: 4px;
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


def prepare_monthly_space(
    df,
    month_col,
    week_col,
    plant_col,
    area_col,
    stacking_height,
):
    work = df.copy()

    work[month_col] = pd.to_numeric(work[month_col], errors="coerce")
    work[week_col] = pd.to_numeric(work[week_col], errors="coerce")
    work[area_col] = pd.to_numeric(work[area_col], errors="coerce").fillna(0)
    work[plant_col] = work[plant_col].astype(str).str.strip()

    work = work.dropna(subset=[month_col, week_col])

    work = work[
        work[plant_col].notna()
        & (work[plant_col] != "")
        & (work[plant_col].str.lower() != "nan")
    ]

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
    # Calculate week-wise total volumetric area for each:
    # Month + Week + Plant
    # --------------------------------------------------------
    weekly = (
        work.groupby([month_col, week_col, plant_col], dropna=False)[area_col]
        .sum()
        .reset_index()
        .rename(columns={area_col: "Weekly_Total_Volumetric_Area"})
    )

    # --------------------------------------------------------
    # Step 2:
    # Divide total volumetric area by stacking height.
    # Default stacking height = 5 feet.
    # --------------------------------------------------------
    weekly["Weekly_Utilized_Floor_Space"] = (
        weekly["Weekly_Total_Volumetric_Area"] / stacking_height
    )

    # --------------------------------------------------------
    # Step 3:
    # Monthly plant average:
    # Monthly Average Space Before Aisle =
    # Sum of weekly utilized floor space / Number of weeks
    # --------------------------------------------------------
    monthly = (
        weekly.groupby([month_col, plant_col], dropna=False)
        .agg(
            Total_Volumetric_Area=("Weekly_Total_Volumetric_Area", "sum"),
            Total_Utilized_Floor_Space=("Weekly_Utilized_Floor_Space", "sum"),
            No_of_Weeks=(week_col, "nunique"),
            Highest_Week_Space=("Weekly_Utilized_Floor_Space", "max"),
            Lowest_Week_Space=("Weekly_Utilized_Floor_Space", "min"),
        )
        .reset_index()
    )

    monthly["Average_Monthly_Space_Before_Aisle"] = (
        monthly["Total_Utilized_Floor_Space"]
        / monthly["No_of_Weeks"].replace(0, pd.NA)
    )

    # --------------------------------------------------------
    # Step 4:
    # Add additional aisle space:
    # Plant I070 = 25%
    # All other plants = 35%
    # --------------------------------------------------------
    monthly["Aisle_Percentage"] = (
        monthly[plant_col]
        .astype(str)
        .str.upper()
        .str.strip()
        .apply(lambda x: 0.25 if x == "I070" else 0.35)
    )

    monthly["Aisle_Space"] = (
        monthly["Average_Monthly_Space_Before_Aisle"]
        * monthly["Aisle_Percentage"]
    )

    monthly["Average_Monthly_Space"] = (
        monthly["Average_Monthly_Space_Before_Aisle"]
        + monthly["Aisle_Space"]
    )

    monthly["Month_Name"] = (
        monthly[month_col]
        .astype(int)
        .map(month_names)
        .fillna(monthly[month_col].astype(str))
    )

    return work, weekly, monthly


def to_excel_bytes(monthly_df, weekly_df, filtered_raw_df):
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        monthly_df.to_excel(writer, index=False, sheet_name="Monthly Plant Average")
        weekly_df.to_excel(writer, index=False, sheet_name="Weekly Plant Summary")
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

        percent_fmt = workbook.add_format(
            {
                "num_format": "0%",
                "border": 1,
            }
        )

        for sheet_name in [
            "Monthly Plant Average",
            "Weekly Plant Summary",
            "Filtered Raw Data",
        ]:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.set_row(0, 22, header_fmt)
            worksheet.set_column(0, 20, 20, text_fmt)
            worksheet.set_column(3, 15, 22, num_fmt)

            if sheet_name == "Monthly Plant Average":
                # Aisle percentage column usually comes around column H/I depending on mapping
                worksheet.set_column(7, 7, 15, percent_fmt)

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
            Upload your weekly stock occupancy file and view month-wise average utilized floor space by plant.
            This dashboard considers stacking height and additional aisle space while calculating practical floor space utilization.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="formula-box">
        <b>Calculation Logic:</b><br>
        1. Weekly Total Volumetric Area = Sum of Volumetric Area column, usually Column AP<br>
        2. Weekly Utilized Floor Space = Weekly Total Volumetric Area ÷ Stacking Height<br>
        3. Monthly Average Space Before Aisle = Sum of Weekly Utilized Floor Space ÷ Number of weeks in that month<br>
        4. Additional Aisle Space = 25% for Plant I070 and 35% for all other plants<br>
        5. Final Average Monthly Space = Monthly Average Space Before Aisle + Additional Aisle Space<br>
        <b>Default stacking height used:</b> 5 feet
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
        help="Your file should contain DATA sheet with Month, Week, Plant, and Volumetric Area columns.",
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

    st.header("🚚 Aisle Space Logic")
    st.info(
        "Plant I070 = 25% additional aisle space\n\n"
        "All other plants = 35% additional aisle space"
    )

    st.markdown("---")
    st.caption("Default expected mapping:")
    st.caption("Month = Column B")
    st.caption("Week = Column C")
    st.caption("Volumetric Area = Column AP")
    st.caption("Plant = select from dropdown")


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
    possible_names=["Month", "month"],
    fallback_index=1,  # Column B
)

default_week_col = find_default_column(
    columns,
    possible_names=["Week", "week"],
    fallback_index=2,  # Column C
)

default_area_col = find_default_column(
    columns,
    possible_names=[
        "volumetric area",
        "Volumetric Area",
        "Volumetric_Area",
        "AP",
    ],
    fallback_index=41 if len(columns) > 41 else None,  # Column AP
)

default_plant_col = find_default_column(
    columns,
    possible_names=[
        "Plant",
        "plant",
        "Location",
        "location",
        "Godown",
        "Warehouse",
        "warehouse",
        "Plant Name",
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

    week_col = st.selectbox(
        "Week column",
        columns,
        index=columns.index(default_week_col),
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
    clean_df, weekly_df, monthly_df = prepare_monthly_space(
        df=df,
        month_col=month_col,
        week_col=week_col,
        plant_col=plant_col,
        area_col=area_col,
        stacking_height=stacking_height,
    )
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
            "Final Average Monthly Space",
            "Average Monthly Space Before Aisle",
            "Additional Aisle Space",
            "Total Utilized Floor Space",
            "Total Volumetric Area",
        ],
        horizontal=False,
    )


filtered_monthly = monthly_df[
    monthly_df[month_col].isin(selected_months)
    & monthly_df[plant_col].astype(str).isin(selected_plants)
].copy()

filtered_weekly = weekly_df[
    weekly_df[month_col].isin(selected_months)
    & weekly_df[plant_col].astype(str).isin(selected_plants)
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
total_floor_space = filtered_monthly["Total_Utilized_Floor_Space"].sum()
avg_space_before_aisle = filtered_monthly["Average_Monthly_Space_Before_Aisle"].mean()
total_aisle_space = filtered_monthly["Aisle_Space"].sum()
final_avg_space = filtered_monthly["Average_Monthly_Space"].mean()
plant_count = filtered_monthly[plant_col].nunique()
record_count = len(filtered_raw)

k1, k2, k3, k4, k5, k6 = st.columns(6)

metric_data = [
    ("Total Volumetric Area", format_number(total_volumetric_area, 2)),
    ("Total Floor Space / 5 Ft", format_number(total_floor_space, 2)),
    ("Avg. Space Before Aisle", format_number(avg_space_before_aisle, 2)),
    ("Total Aisle Space", format_number(total_aisle_space, 2)),
    ("Final Avg. Space", format_number(final_avg_space, 2)),
    ("Plants", format_number(plant_count, 0)),
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


# ------------------------------------------------------------
# Chart column selection
# ------------------------------------------------------------
 if_chart_label = ""

if chart_mode == "Final Average Monthly Space":
    chart_col = "Average_Monthly_Space"
    y_title = "Final Average Monthly Space"
elif chart_mode == "Average Monthly Space Before Aisle":
    chart_col = "Average_Monthly_Space_Before_Aisle"
    y_title = "Average Monthly Space Before Aisle"
elif chart_mode == "Additional Aisle Space":
    chart_col = "Aisle_Space"
    y_title = "Additional Aisle Space"
elif chart_mode == "Total Utilized Floor Space":
    chart_col = "Total_Utilized_Floor_Space"
    y_title = "Total Utilized Floor Space"
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
        title="Final Average Monthly Floor Space Ranking by Plant",
        labels={
            "Average_Monthly_Space": "Final Avg. Monthly Floor Space",
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
            "Average_Monthly_Space"
        ]
        .sum()
        .sort_values(month_col)
    )

    fig_line = px.line(
        trend_df,
        x="Month_Name",
        y="Average_Monthly_Space",
        markers=True,
        title="Final Average Floor Space Trend Month-wise",
        labels={
            "Month_Name": "Month",
            "Average_Monthly_Space": "Final Average Floor Space",
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
        "Average_Monthly_Space"
    ].sum()

    fig_donut = px.pie(
        donut_df,
        names=plant_col,
        values="Average_Monthly_Space",
        hole=0.45,
        title=f"Plant Share in Month {latest_month} After Aisle Space",
    )

    fig_donut.update_layout(
        height=420,
        margin=dict(l=10, r=10, t=55, b=10),
    )

    st.plotly_chart(fig_donut, use_container_width=True)


# ------------------------------------------------------------
# Weekly trend chart
# ------------------------------------------------------------
st.markdown(
    '<div class="section-title">📈 Weekly Floor Space Movement</div>',
    unsafe_allow_html=True,
)

weekly_trend = filtered_weekly.copy()
weekly_trend["Month_Week"] = (
    "M"
    + weekly_trend[month_col].astype(int).astype(str)
    + " - W"
    + weekly_trend[week_col].astype(int).astype(str)
)

fig_weekly = px.line(
    weekly_trend.sort_values([month_col, week_col, plant_col]),
    x="Month_Week",
    y="Weekly_Utilized_Floor_Space",
    color=plant_col,
    markers=True,
    title="Weekly Utilized Floor Space by Plant Before Aisle Loading",
    labels={
        "Month_Week": "Month - Week",
        "Weekly_Utilized_Floor_Space": "Weekly Floor Space Before Aisle",
        plant_col: "Plant",
    },
)

fig_weekly.update_layout(
    height=430,
    margin=dict(l=10, r=10, t=55, b=10),
)

st.plotly_chart(fig_weekly, use_container_width=True)


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
    "Total_Utilized_Floor_Space",
    "No_of_Weeks",
    "Average_Monthly_Space_Before_Aisle",
    "Aisle_Percentage",
    "Aisle_Space",
    "Average_Monthly_Space",
    "Highest_Week_Space",
    "Lowest_Week_Space",
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
        "Total_Utilized_Floor_Space": st.column_config.NumberColumn(
            "Total Utilized Floor Space",
            format="%.2f",
        ),
        "Average_Monthly_Space_Before_Aisle": st.column_config.NumberColumn(
            "Average Monthly Space Before Aisle",
            format="%.2f",
        ),
        "Aisle_Percentage": st.column_config.NumberColumn(
            "Aisle %",
            format="%.0f%%",
        ),
        "Aisle_Space": st.column_config.NumberColumn(
            "Additional Aisle Space",
            format="%.2f",
        ),
        "Average_Monthly_Space": st.column_config.NumberColumn(
            "Final Average Monthly Space",
            format="%.2f",
        ),
        "Highest_Week_Space": st.column_config.NumberColumn(
            "Highest Week Space",
            format="%.2f",
        ),
        "Lowest_Week_Space": st.column_config.NumberColumn(
            "Lowest Week Space",
            format="%.2f",
        ),
    },
)


with st.expander("View weekly plant-wise summary"):
    st.dataframe(
        filtered_weekly.sort_values([month_col, week_col, plant_col]),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Weekly_Total_Volumetric_Area": st.column_config.NumberColumn(
                "Weekly Total Volumetric Area",
                format="%.2f",
            ),
            "Weekly_Utilized_Floor_Space": st.column_config.NumberColumn(
                "Weekly Utilized Floor Space Before Aisle",
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
    weekly_df=filtered_weekly,
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
        1. First, the app sums volumetric area at Month + Week + Plant level.<br>
        2. Then, it divides the total volumetric area by <b>{stacking_height} feet stacking height</b>
        to calculate utilized floor space.<br>
        3. After that, Month + Plant average is calculated as total utilized floor space divided by
        the number of weeks available in that month.<br>
        4. Then, additional aisle space is added:
        <b>25% for Plant I070</b> and <b>35% for all other plants</b>.<br>
        5. Final Average Monthly Space = Average Monthly Space Before Aisle + Additional Aisle Space.
    </div>
    """,
    unsafe_allow_html=True,
)
