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


# ------------------------------------------------------------
# CSS
# ------------------------------------------------------------
st.markdown(
    """
    <style>
        .main {
            background-color: #f7f9fc;
        }

        .block-container {
            padding-top: 1.5rem;
            padding-bottom: 2rem;
        }

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

        .section-title {
            font-size: 21px;
            font-weight: 800;
            color: #0f172a;
            margin-top: 14px;
            margin-bottom: 8px;
        }

        .metric-card {
            background: white;
            padding: 16px 10px;
            border-radius: 18px;
            border: 1px solid #e5e7eb;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.07);
            min-height: 118px;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            width: 100%;
        }

        .metric-label {
            font-size: 13px;
            color: #64748b;
            font-weight: 800;
            margin-bottom: 8px;
            line-height: 1.25;
            text-align: center;
        }

        .metric-value {
            font-size: 20px;
            color: #0f172a;
            font-weight: 900;
            line-height: 1.2;
            text-align: center;
            white-space: nowrap;
        }

        .metric-unit {
            font-size: 12px;
            color: #475569;
            font-weight: 800;
            margin-top: 4px;
            text-align: center;
        }

        .small-note {
            font-size: 12px;
            color: #64748b;
            margin-top: 8px;
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
        key = str(name).strip().lower()
        if key in lower_map:
            return lower_map[key]

    if fallback_index is not None and fallback_index < len(columns):
        return columns[fallback_index]

    return columns[0]


def get_plant_display_name(plant_code):
    plant_map = {
        "I070": "DEL WH",
        "I030": "ZRK WH",
        "I080": "JAI WH",
        "I360": "HYD WH",
        "I290": "BNG WH",
        "I270": "MUM WH",
        "I190": "KOL WH",
        "I330": "CHN WH",
    }

    plant_code = str(plant_code).strip().upper()
    return plant_map.get(plant_code, plant_code)


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
    ].copy()

    work["Plant_Name"] = work[plant_col].apply(get_plant_display_name)

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
    # Weekly calculation
    # --------------------------------------------------------
    weekly = (
        work.groupby([month_col, week_col, plant_col, "Plant_Name"], dropna=False)[area_col]
        .sum()
        .reset_index()
        .rename(columns={area_col: "Weekly_Total_Volumetric_Area"})
    )

    weekly["Weekly_Utilized_Floor_Space"] = (
        weekly["Weekly_Total_Volumetric_Area"] / stacking_height
    )

    # Weekly aisle logic:
    # I070 / DEL WH = 25%
    # All other warehouses = 35%
    weekly["Aisle_Percentage"] = (
        weekly[plant_col]
        .astype(str)
        .str.upper()
        .str.strip()
        .apply(lambda x: 0.25 if x == "I070" else 0.35)
    )

    weekly["Weekly_Aisle_Space"] = (
        weekly["Weekly_Utilized_Floor_Space"] * weekly["Aisle_Percentage"]
    )

    weekly["Weekly_Final_Floor_Space"] = (
        weekly["Weekly_Utilized_Floor_Space"] + weekly["Weekly_Aisle_Space"]
    )

    # --------------------------------------------------------
    # Monthly calculation
    # --------------------------------------------------------
    monthly = (
        weekly.groupby([month_col, plant_col, "Plant_Name"], dropna=False)
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
                "num_format": "#,##0",
                "border": 1,
            }
        )

        percent_fmt = workbook.add_format(
            {
                "num_format": "0%",
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
            "Weekly Plant Summary",
            "Filtered Raw Data",
        ]:
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            worksheet.set_row(0, 22, header_fmt)
            worksheet.set_column(0, 25, 20, text_fmt)
            worksheet.set_column(3, 20, 22, num_fmt)

            if sheet_name == "Monthly Plant Average":
                worksheet.set_column(8, 8, 15, percent_fmt)

            if sheet_name == "Weekly Plant Summary":
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
            Upload your weekly stock occupancy file and view month-wise and week-wise utilized floor space by warehouse.
            This dashboard considers stacking height and additional aisle space while calculating practical floor space utilization in Sqft.
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
        3. Weekly Final Floor Space = Weekly Utilized Floor Space + Weekly Aisle Space<br>
        4. Monthly Average Space Before Aisle = Sum of Weekly Utilized Floor Space ÷ Number of weeks in that month<br>
        5. Additional Aisle Space = 25% for DEL WH / I070 and 35% for all other warehouses<br>
        6. Final Average Monthly Space = Monthly Average Space Before Aisle + Additional Aisle Space<br>
        <b>Measurement Unit:</b> Sqft
    </div>
    """,
    unsafe_allow_html=True,
)


# ------------------------------------------------------------
# Sidebar
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
    )

    st.markdown("---")

    st.header("🏭 Warehouse Mapping")
    st.caption("I070 = DEL WH")
    st.caption("I030 = ZRK WH")
    st.caption("I080 = JAI WH")
    st.caption("I360 = HYD WH")
    st.caption("I290 = BNG WH")
    st.caption("I270 = MUM WH")
    st.caption("I190 = KOL WH")
    st.caption("I330 = CHN WH")

    st.markdown("---")

    st.header("🚚 Aisle Space Logic")
    st.info(
        "DEL WH / I070 = 25% additional aisle space\n\n"
        "All other warehouses including CHN WH / I330 = 35% additional aisle space"
    )


if uploaded_file is None:
    st.info("Please upload your Stock Occupancy Excel file to start the dashboard.")
    st.stop()


uploaded_bytes = uploaded_file.getvalue()


# ------------------------------------------------------------
# Load workbook and sheet
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
    fallback_index=1,
)

default_week_col = find_default_column(
    columns,
    possible_names=["Week", "week"],
    fallback_index=2,
)

default_area_col = find_default_column(
    columns,
    possible_names=[
        "volumetric area",
        "Volumetric Area",
        "Volumetric_Area",
        "AP",
    ],
    fallback_index=41 if len(columns) > 41 else None,
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

f1, f2, f3 = st.columns([1, 2, 1])

all_months = sorted(monthly_df[month_col].dropna().unique())
all_plants = sorted(monthly_df["Plant_Name"].dropna().astype(str).unique())

with f1:
    selected_months = st.multiselect(
        "Select Month",
        all_months,
        default=all_months,
    )

with f2:
    selected_plants = st.multiselect(
        "Select Warehouse / Plant",
        all_plants,
        default=all_plants,
    )

with f3:
    chart_mode = st.radio(
        "Monthly Chart Value",
        [
            "Final Average Space",
            "Space Before Aisle",
            "Aisle Space",
            "Total Floor Space",
            "Total Volumetric Area",
        ],
        index=0,
    )


filtered_monthly = monthly_df[
    monthly_df[month_col].isin(selected_months)
    & monthly_df["Plant_Name"].astype(str).isin(selected_plants)
].copy()

filtered_weekly = weekly_df[
    weekly_df[month_col].isin(selected_months)
    & weekly_df["Plant_Name"].astype(str).isin(selected_plants)
].copy()

filtered_raw = clean_df[
    clean_df[month_col].isin(selected_months)
    & clean_df["Plant_Name"].astype(str).isin(selected_plants)
].copy()


if filtered_monthly.empty:
    st.warning("No records found for the selected Month / Warehouse filter.")
    st.stop()


# ------------------------------------------------------------
# KPI cards in one horizontal line
# ------------------------------------------------------------
plant_count = filtered_monthly["Plant_Name"].nunique()
month_count = filtered_monthly[month_col].nunique()
week_count = filtered_weekly[week_col].nunique()

total_floor_space = filtered_monthly["Total_Utilized_Floor_Space"].sum()
total_aisle_space = filtered_monthly["Aisle_Space"].sum()
final_avg_space = filtered_monthly["Average_Monthly_Space"].mean()

k1, k2, k3, k4, k5, k6 = st.columns(6)

metrics = [
    ("Months", format_number(month_count, 0), ""),
    ("Weeks", format_number(week_count, 0), ""),
    ("Warehouses", format_number(plant_count, 0), ""),
    ("Floor Space", format_number(total_floor_space, 0), "Sqft"),
    ("Aisle Space", format_number(total_aisle_space, 0), "Sqft"),
    ("Final Avg. Space", format_number(final_avg_space, 0), "Sqft"),
]

for col, (label, value, unit) in zip([k1, k2, k3, k4, k5, k6], metrics):
    with col:
        unit_html = f'<div class="metric-unit">{unit}</div>' if unit else ""

        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value">{value}</div>
                {unit_html}
            </div>
            """,
            unsafe_allow_html=True,
        )


# ------------------------------------------------------------
# Monthly chart value selection
# ------------------------------------------------------------
if chart_mode == "Final Average Space":
    chart_col = "Average_Monthly_Space"
    y_title = "Final Average Space Including Aisle Sqft"

elif chart_mode == "Space Before Aisle":
    chart_col = "Average_Monthly_Space_Before_Aisle"
    y_title = "Average Space Before Aisle Sqft"

elif chart_mode == "Aisle Space":
    chart_col = "Aisle_Space"
    y_title = "Aisle Space Sqft"

elif chart_mode == "Total Floor Space":
    chart_col = "Total_Utilized_Floor_Space"
    y_title = "Total Floor Space Sqft"

else:
    chart_col = "Total_Volumetric_Area"
    y_title = "Total Volumetric Area Sqft"


# ------------------------------------------------------------
# Monthly bar chart
# ------------------------------------------------------------
st.markdown(
    '<div class="section-title">📊 Monthly Warehouse-wise Space Analysis</div>',
    unsafe_allow_html=True,
)

fig_bar = px.bar(
    filtered_monthly.sort_values([month_col, "Plant_Name"]),
    x="Month_Name",
    y=chart_col,
    color="Plant_Name",
    barmode="group",
    text_auto=".0s",
    title=f"Warehouse-wise {y_title} by Month",
    labels={
        "Month_Name": "Month",
        chart_col: y_title,
        "Plant_Name": "Warehouse",
    },
)

fig_bar.update_layout(
    height=470,
    title_font_size=18,
    legend_title_text="Warehouse",
    margin=dict(l=10, r=10, t=55, b=10),
)

st.plotly_chart(fig_bar, use_container_width=True)


# ------------------------------------------------------------
# Weekly line chart including aisle space
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
    weekly_trend.sort_values([month_col, week_col, "Plant_Name"]),
    x="Month_Week",
    y="Weekly_Final_Floor_Space",
    color="Plant_Name",
    markers=True,
    title="Weekly Floor Space by Warehouse Including Aisle Space",
    labels={
        "Month_Week": "Month - Week",
        "Weekly_Final_Floor_Space": "Weekly Floor Space Including Aisle Sqft",
        "Plant_Name": "Warehouse",
    },
)

fig_weekly.update_layout(
    height=430,
    margin=dict(l=10, r=10, t=55, b=10),
    legend_title_text="Warehouse",
)

st.plotly_chart(fig_weekly, use_container_width=True)


# ------------------------------------------------------------
# Monthly table
# ------------------------------------------------------------
st.markdown(
    '<div class="section-title">📋 Monthly Average Calculation Table</div>',
    unsafe_allow_html=True,
)

show_cols = [
    month_col,
    "Month_Name",
    plant_col,
    "Plant_Name",
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

view_df = filtered_monthly[show_cols].sort_values([month_col, "Plant_Name"]).copy()

st.dataframe(
    view_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        plant_col: st.column_config.TextColumn("Plant Code"),
        "Plant_Name": st.column_config.TextColumn("Warehouse Name"),
        "Total_Volumetric_Area": st.column_config.NumberColumn(
            "Total Volumetric Area Sqft",
            format="%.0f",
        ),
        "Total_Utilized_Floor_Space": st.column_config.NumberColumn(
            "Total Floor Space Before Aisle Sqft",
            format="%.0f",
        ),
        "Average_Monthly_Space_Before_Aisle": st.column_config.NumberColumn(
            "Average Monthly Space Before Aisle Sqft",
            format="%.0f",
        ),
        "Aisle_Percentage": st.column_config.NumberColumn(
            "Aisle %",
            format="%.0f%%",
        ),
        "Aisle_Space": st.column_config.NumberColumn(
            "Additional Aisle Space Sqft",
            format="%.0f",
        ),
        "Average_Monthly_Space": st.column_config.NumberColumn(
            "Final Average Monthly Space Including Aisle Sqft",
            format="%.0f",
        ),
        "Highest_Week_Space": st.column_config.NumberColumn(
            "Highest Week Space Before Aisle Sqft",
            format="%.0f",
        ),
        "Lowest_Week_Space": st.column_config.NumberColumn(
            "Lowest Week Space Before Aisle Sqft",
            format="%.0f",
        ),
    },
)


# ------------------------------------------------------------
# Weekly table
# ------------------------------------------------------------
with st.expander("View weekly warehouse-wise summary"):
    st.dataframe(
        filtered_weekly.sort_values([month_col, week_col, "Plant_Name"]),
        use_container_width=True,
        hide_index=True,
        column_config={
            plant_col: st.column_config.TextColumn("Plant Code"),
            "Plant_Name": st.column_config.TextColumn("Warehouse Name"),
            "Weekly_Total_Volumetric_Area": st.column_config.NumberColumn(
                "Weekly Total Volumetric Area Sqft",
                format="%.0f",
            ),
            "Weekly_Utilized_Floor_Space": st.column_config.NumberColumn(
                "Weekly Floor Space Before Aisle Sqft",
                format="%.0f",
            ),
            "Aisle_Percentage": st.column_config.NumberColumn(
                "Aisle %",
                format="%.0f%%",
            ),
            "Weekly_Aisle_Space": st.column_config.NumberColumn(
                "Weekly Aisle Space Sqft",
                format="%.0f",
            ),
            "Weekly_Final_Floor_Space": st.column_config.NumberColumn(
                "Weekly Final Floor Space Including Aisle Sqft",
                format="%.0f",
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


# ------------------------------------------------------------
# Footer note
# ------------------------------------------------------------
st.markdown(
    f"""
    <div class="small-note">
        Calculation logic used in this dashboard:<br>
        1. First, the app sums volumetric area at Month + Week + Warehouse level.<br>
        2. Then, it divides the total volumetric area by <b>{stacking_height} feet stacking height</b>
        to calculate utilized floor space in <b>Sqft</b>.<br>
        3. Weekly chart shows floor space after adding aisle space.<br>
        4. Additional aisle space:
        <b>25% for DEL WH / I070</b> and <b>35% for all other warehouses including CHN WH / I330</b>.<br>
        5. Final Average Monthly Space = Average Monthly Space Before Aisle + Additional Aisle Space.
    </div>
    """,
    unsafe_allow_html=True,
)
