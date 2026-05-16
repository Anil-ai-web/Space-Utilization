import io
from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st

# ------------------------- PAGE SETUP -------------------------
st.set_page_config(page_title="Plant Space Utilization Dashboard",
                   page_icon="🏭", layout="wide",
                   initial_sidebar_state="expanded")

st.markdown("""
<style>
.metric-card {background:white;padding:18px 16px;border-radius:18px;border:1px solid #e5e7eb;
              box-shadow:0 6px 18px rgba(15,23,42,0.07);min-height:125px;display:flex;
              flex-direction:column;justify-content:center;}
.metric-label {font-size:14px;color:#64748b;font-weight:800;margin-bottom:10px;text-align:center;}
.metric-value {font-size:21px;color:#0f172a;font-weight:900;line-height:1.25;text-align:center;}
.metric-unit {font-size:13px;color:#475569;font-weight:800;margin-top:4px;text-align:center;}
.section-title {font-size:21px;font-weight:800;color:#0f172a;margin-top:14px;margin-bottom:8px;}
</style>
""", unsafe_allow_html=True)

# ------------------------- HELPERS -------------------------
def excel_engine_for_file(file_name: str):
    suffix = Path(file_name).suffix.lower()
    if suffix==".xlsb": return "pyxlsb"
    if suffix in [".xlsx",".xlsm"]: return "openpyxl"
    return None

@st.cache_data
def get_sheet_names(file_bytes: bytes, file_name: str):
    engine = excel_engine_for_file(file_name)
    return pd.ExcelFile(io.BytesIO(file_bytes), engine=engine).sheet_names

@st.cache_data
def load_excel_data(file_bytes: bytes, file_name: str, sheet_name: str):
    engine = excel_engine_for_file(file_name)
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, engine=engine)
    df.columns = [str(c).strip() for c in df.columns]
    return df

def format_number(val, decimals=0): return f"{val:,.{decimals}f}" if not pd.isna(val) else "0"

def get_plant_display_name(code):
    mapping = {"I070":"DEL WH","I030":"ZRK WH","I080":"JAI WH","I360":"HYD WH",
               "I290":"BNG WH","I270":"MUM WH","I190":"KOL WH","I330":"CHN WH"}
    return mapping.get(str(code).strip().upper(), str(code))

# ------------------------- CALCULATION -------------------------
def prepare_monthly_space(df, month_col, week_col, plant_col, area_col, stacking_height):
    work = df.copy()
    work[month_col]=pd.to_numeric(work[month_col],errors="coerce")
    work[week_col]=pd.to_numeric(work[week_col],errors="coerce")
    work[area_col]=pd.to_numeric(work[area_col],errors="coerce").fillna(0)
    work[plant_col]=work[plant_col].astype(str).str.strip()
    work=work.dropna(subset=[month_col,week_col])
    work=work[work[plant_col].notna()&(work[plant_col]!="")&(work[plant_col].str.lower()!="nan")]
    work["Plant_Name"]=work[plant_col].apply(get_plant_display_name)
    
    weekly = work.groupby([month_col, week_col, plant_col, "Plant_Name"], dropna=False)[area_col].sum().reset_index()
    weekly.rename(columns={area_col:"Weekly_Total_Volumetric_Area"}, inplace=True)
    weekly["Weekly_Utilized_Floor_Space"]=weekly["Weekly_Total_Volumetric_Area"]/stacking_height
    weekly["Aisle_Percentage"]=weekly[plant_col].astype(str).str.upper().str.strip().apply(lambda x:0.25 if x=="I070" else 0.35)
    weekly["Weekly_Aisle_Space"]=weekly["Weekly_Utilized_Floor_Space"]*weekly["Aisle_Percentage"]
    weekly["Weekly_Final_Floor_Space"]=weekly["Weekly_Utilized_Floor_Space"]+weekly["Weekly_Aisle_Space"]

    monthly = weekly.groupby([month_col,plant_col,"Plant_Name"],dropna=False).agg(
        Total_Volumetric_Area=("Weekly_Total_Volumetric_Area","sum"),
        Total_Utilized_Floor_Space=("Weekly_Utilized_Floor_Space","sum"),
        No_of_Weeks=(week_col,"nunique"),
        Highest_Week_Space=("Weekly_Utilized_Floor_Space","max"),
        Lowest_Week_Space=("Weekly_Utilized_Floor_Space","min")
    ).reset_index()
    monthly["Average_Monthly_Space_Before_Aisle"]=monthly["Total_Utilized_Floor_Space"]/monthly["No_of_Weeks"].replace(0,pd.NA)
    monthly["Aisle_Percentage"]=monthly[plant_col].astype(str).str.upper().str.strip().apply(lambda x:0.25 if x=="I070" else 0.35)
    monthly["Aisle_Space"]=monthly["Average_Monthly_Space_Before_Aisle"]*monthly["Aisle_Percentage"]
    monthly["Average_Monthly_Space"]=monthly["Average_Monthly_Space_Before_Aisle"]+monthly["Aisle_Space"]
    monthly["Month_Name"]=monthly[month_col].astype(int).map({1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",
                                                              6:"Jun",7:"Jul",8:"Aug",9:"Sep",
                                                              10:"Oct",11:"Nov",12:"Dec"}).fillna(monthly[month_col].astype(str))
    return work, weekly, monthly

# ------------------------- SIDEBAR -------------------------
with st.sidebar:
    st.header("Upload Stock File")
    uploaded_file=st.file_uploader("Excel file with weekly stock",type=["xlsb","xlsx","xlsm"])
    stacking_height=st.number_input("Stacking height (ft)",min_value=1.0,max_value=50.0,value=5.0,step=0.5)
    st.header("Warehouse mapping")
    st.caption("I070=DEL WH, I030=ZRK WH, I080=JAI WH, I360=HYD WH, I290=BNG WH, I270=MUM WH, I190=KOL WH, I330=CHN WH")

if uploaded_file is None: st.stop()
file_bytes=uploaded_file.getvalue()
sheets=get_sheet_names(file_bytes, uploaded_file.name)
sheet_name="DATA" if "DATA" in sheets else sheets[0]
df=load_excel_data(file_bytes, uploaded_file.name, sheet_name)

columns=list(df.columns)
month_col=columns[1]
week_col=columns[2]
area_col=columns[41] if len(columns)>41 else columns[-1]
plant_col=columns[6] if len(columns)>6 else columns[0]

clean_df, weekly_df, monthly_df = prepare_monthly_space(df, month_col, week_col, plant_col, area_col, stacking_height)

# ------------------------- FILTERS -------------------------
st.markdown('<div class="section-title">🔎 Filters</div>', unsafe_allow_html=True)
f1,f2 = st.columns([1,2])
all_months = sorted(monthly_df[month_col].dropna().unique())
all_plants = sorted(monthly_df["Plant_Name"].dropna().astype(str).unique())

with f1:
    selected_months=st.multiselect("Select Month",all_months,default=all_months)
with f2:
    selected_plants=st.multiselect("Select Warehouse / Plant",all_plants,default=all_plants)

filtered_monthly = monthly_df[monthly_df[month_col].isin(selected_months) & monthly_df["Plant_Name"].isin(selected_plants)]
filtered_weekly = weekly_df[weekly_df[month_col].isin(selected_months) & weekly_df["Plant_Name"].isin(selected_plants)]
filtered_raw = clean_df[clean_df[month_col].isin(selected_months) & clean_df["Plant_Name"].isin(selected_plants)]

if filtered_monthly.empty: st.warning("No records for selected Month/Warehouse filter."); st.stop()

# ------------------------- KPI CARDS -------------------------
plant_count = filtered_monthly["Plant_Name"].nunique()
month_count = filtered_monthly[month_col].nunique()
week_count = filtered_weekly[week_col].nunique()
total_floor_space = filtered_monthly["Total_Utilized_Floor_Space"].sum()
total_aisle_space = filtered_monthly["Aisle_Space"].sum()
final_avg_space = filtered_monthly["Average_Monthly_Space"].mean()

k1,k2,k3,k4,k5,k6 = st.columns(6)
metrics=[("Months", format_number(month_count,0),""),("Weeks", format_number(week_count,0),""),("Warehouses", format_number(plant_count,0),""),("Floor Space",format_number(total_floor_space,0),"Sqft"),("Aisle Space",format_number(total_aisle_space,0),"Sqft"),("Final Avg. Space",format_number(final_avg_space,0),"Sqft")]

for col,(label,value,unit) in zip([k1,k2,k3,k4,k5,k6],metrics):
    with col: unit_html=f'<div class="metric-unit">{unit}</div>' if unit else ""
    st.markdown(f"""<div class="metric-card"><div class="metric-label">{label}</div><div class="metric-value">{value}</div>{unit_html}</div>""",unsafe_allow_html=True)

# ------------------------- MONTHLY BAR CHART -------------------------
fig_bar=px.bar(filtered_monthly.sort_values([month_col,"Plant_Name"]),x="Month_Name",y="Average_Monthly_Space",color="Plant_Name",
               barmode="group",text_auto=".0s",title="Warehouse-wise Final Avg Space by Month Including Aisle",
               labels={"Month_Name":"Month","Average_Monthly_Space":"Final Avg Space Sqft","Plant_Name":"Warehouse"})
st.plotly_chart(fig_bar,use_container_width=True)

# ------------------------- WEEKLY LINE CHART -------------------------
filtered_weekly["Month_Week"]="M"+filtered_weekly[month_col].astype(int).astype(str)+" - W"+filtered_weekly[week_col].astype(int).astype(str)
fig_weekly=px.line(filtered_weekly.sort_values([month_col,week_col,"Plant_Name"]),x="Month_Week",y="Weekly_Final_Floor_Space",color="Plant_Name",
                   markers=True,title="Weekly Utilized Floor Space by Warehouse Including Aisle",
                   labels={"Month_Week":"Month-Week","Weekly_Final_Floor_Space":"Weekly Floor Space Including Aisle Sqft","Plant_Name":"Warehouse"})
st.plotly_chart(fig_weekly,use_container_width=True)
