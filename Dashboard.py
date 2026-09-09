# =========================================================
# Dashboard.py
# Advertising Campaign Performance Dashboard
# =========================================================

from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Advertising Campaign Performance Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# DATA FILE PATH
# =========================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "data" / "advertising_dataset.csv"


# =========================================================
# THEME COLORS
# =========================================================

COLOR_BG = "#F8FAFC"
COLOR_CARD = "#FFFFFF"
COLOR_BORDER = "#E2E8F0"
COLOR_TEXT = "#0F172A"
COLOR_MUTED = "#64748B"

COLOR_BLUE = "#2563EB"
COLOR_TEAL = "#14B8A6"
COLOR_LIGHT_TEAL = "#9BDDD7"
COLOR_AMBER = "#F59E0B"
COLOR_PURPLE = "#7C3AED"
COLOR_GREEN = "#22C55E"

PLATFORM_PALETTE = [
    "#2563EB", "#14B8A6", "#F59E0B", "#7C3AED", "#22C55E",
    "#EF4444", "#0EA5E9", "#A855F7", "#10B981", "#F97316",
]

TEAL_SCALE = [[0.0, "#CCFBF1"], [1.0, COLOR_TEAL]]
BLUE_SCALE = [[0.0, "#DBEAFE"], [1.0, COLOR_BLUE]]


# =========================================================
# CUSTOM CSS
# =========================================================

st.markdown(
    f"""
    <style>
        .stApp {{
            background-color: {COLOR_BG};
        }}

        .block-container {{
            padding-top: 2rem;
            padding-left: 2.2rem;
            padding-right: 2.2rem;
            padding-bottom: 3rem;
        }}

        section[data-testid="stSidebar"] {{
            background-color: #F1F5F9;
            border-right: 1px solid {COLOR_BORDER};
        }}

        h1, h2, h3, h4, p, div, span, label {{
            font-family: "Inter", "Segoe UI", sans-serif;
        }}

        .dashboard-header {{
            background: linear-gradient(135deg, #FFFFFF 0%, #EFF6FF 100%);
            border: 1px solid {COLOR_BORDER};
            border-radius: 24px;
            padding: 28px 32px;
            margin-bottom: 22px;
            box-shadow: 0px 8px 24px rgba(15, 23, 42, 0.05);
        }}

        .dashboard-title {{
            font-size: 31px;
            font-weight: 800;
            color: {COLOR_TEXT};
            margin-bottom: 6px;
            letter-spacing: -0.03em;
        }}

        .dashboard-subtitle {{
            font-size: 15px;
            color: {COLOR_MUTED};
        }}

        div[data-testid="stVerticalBlockBorderWrapper"] {{
            background-color: {COLOR_CARD} !important;
            border: 1px solid {COLOR_BORDER} !important;
            border-radius: 22px !important;
            box-shadow: 0px 8px 22px rgba(15, 23, 42, 0.045) !important;
            padding: 18px 20px 12px 20px !important;
            margin-bottom: 18px !important;
        }}

        .section-title {{
            font-size: 22px;
            font-weight: 750;
            color: {COLOR_TEXT};
            margin-top: 24px;
            margin-bottom: 4px;
            letter-spacing: -0.02em;
        }}

        .section-subtitle {{
            font-size: 14px;
            color: {COLOR_MUTED};
            margin-bottom: 16px;
        }}

        .metric-card {{
            background: {COLOR_CARD};
            border: 1px solid {COLOR_BORDER};
            border-radius: 20px;
            padding: 20px 22px;
            min-height: 138px;
            box-shadow: 0px 8px 22px rgba(15, 23, 42, 0.045);
        }}

        .metric-label {{
            color: {COLOR_MUTED};
            font-size: 13px;
            font-weight: 650;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            margin-bottom: 10px;
        }}

        .metric-value {{
            color: {COLOR_TEXT};
            font-size: 29px;
            font-weight: 800;
            letter-spacing: -0.03em;
            margin-bottom: 8px;
        }}

        .metric-desc {{
            color: {COLOR_MUTED};
            font-size: 12.5px;
            line-height: 1.45;
        }}

        .chart-title {{
            color: {COLOR_TEXT};
            font-size: 18px;
            font-weight: 750;
            margin-bottom: 3px;
            letter-spacing: -0.015em;
        }}

        .chart-subtitle {{
            color: {COLOR_MUTED};
            font-size: 13px;
            margin-bottom: 12px;
        }}

        .stSelectbox label, .stDateInput label {{
            color: {COLOR_MUTED} !important;
            font-size: 12px !important;
            font-weight: 700 !important;
            text-transform: uppercase;
            letter-spacing: 0.05em;
        }}

        div[data-baseweb="select"] > div {{
            border-radius: 12px;
            border-color: {COLOR_BORDER};
        }}

        div[data-testid="stDataFrame"] {{
            border-radius: 18px;
            overflow: hidden;
        }}
    </style>
    """,
    unsafe_allow_html=True
)


# =========================================================
# HELPER FUNCTIONS
# =========================================================

@st.cache_data(show_spinner=False)
def load_data(file_path):
    file_path = Path(file_path)

    if not file_path.exists():
        st.error(f"Dataset not found: {file_path}")
        st.stop()

    if file_path.suffix.lower() == ".csv":
        return pd.read_csv(file_path)

    if file_path.suffix.lower() in [".xlsx", ".xlsm", ".xls"]:
        return pd.read_excel(file_path)

    st.error("Unsupported file format. Please use CSV or Excel.")
    st.stop()


def standardize_columns(df):
    df = df.copy()

    column_map = {
        "campaign_id": "CampaignID",
        "campaign id": "CampaignID",
        "campaignid": "CampaignID",

        "campaign_name": "CampaignName",
        "campaign name": "CampaignName",
        "campaignname": "CampaignName",

        "company": "Company",

        "platform": "AdPlatform",
        "ad platform": "AdPlatform",
        "adplatform": "AdPlatform",

        "format": "AdFormat",
        "ad format": "AdFormat",
        "adformat": "AdFormat",

        "objective": "Objective",

        "gender": "TargetGender",
        "target gender": "TargetGender",
        "targetgender": "TargetGender",

        "age group": "TargetAgeGroup",
        "target age group": "TargetAgeGroup",
        "targetagegroup": "TargetAgeGroup",

        "interests": "TargetInterests",
        "target interests": "TargetInterests",
        "targetinterests": "TargetInterests",

        "cost": "Cost",
        "spend": "Cost",
        "amount spent": "Cost",

        "budget": "Budget",
        "impressions": "Impressions",
        "clicks": "Clicks",
        "ctr": "CTR",
        "cpm": "CPM",
        "cpc": "CPC",

        "start date": "StartDate",
        "startdate": "StartDate",
        "end date": "EndDate",
        "enddate": "EndDate",

        "month": "Month",
        "monthnum": "MonthNum",
        "month number": "MonthNum",
        "year": "Year",
    }

    rename_dict = {}

    for col in df.columns:
        clean_col = str(col).strip().lower()
        if clean_col in column_map:
            rename_dict[col] = column_map[clean_col]

    return df.rename(columns=rename_dict)


def parse_date_column(series):
    """
    Safely parse campaign dates for dashboard filtering and monthly charts.

    Priority:
    1. Malaysia-style dates such as DD/MM/YYYY or DD-MM-YYYY using dayfirst=True.
    2. Fallback parsing for ISO-style dates such as YYYY-MM-DD.
    3. Excel serial-number dates if the dataset stores dates as numeric values.
    """
    original = series.copy()

    parsed = pd.to_datetime(original, errors="coerce", dayfirst=True)

    missing_mask = parsed.isna() & original.notna()
    if missing_mask.any():
        fallback = pd.to_datetime(
            original[missing_mask],
            errors="coerce",
            dayfirst=False
        )
        parsed.loc[missing_mask] = fallback

    missing_mask = parsed.isna() & original.notna()
    if missing_mask.any():
        numeric_dates = pd.to_numeric(original[missing_mask], errors="coerce")
        excel_dates = pd.to_datetime(
            numeric_dates,
            errors="coerce",
            unit="D",
            origin="1899-12-30"
        )
        parsed.loc[missing_mask] = excel_dates

    return parsed


def clean_data(df):
    df = standardize_columns(df.copy())

    numeric_cols = ["Cost", "Budget", "Impressions", "Clicks", "CTR", "CPM", "CPC"]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace("RM", "", regex=False)
                .str.replace(",", "", regex=False)
                .str.replace("%", "", regex=False)
                .str.strip()
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")

    for col in ["StartDate", "EndDate"]:
        if col in df.columns:
            df[col] = parse_date_column(df[col])

    if "StartDate" in df.columns:
        df["Year"] = df["StartDate"].dt.year
        df["MonthNum"] = df["StartDate"].dt.month
        df["MonthPeriod"] = df["StartDate"].dt.to_period("M").astype(str)
        df.loc[df["StartDate"].isna(), "MonthPeriod"] = "Unknown"

    text_cols = [
        "CampaignID",
        "CampaignName",
        "Company",
        "AdPlatform",
        "AdFormat",
        "Objective",
        "TargetGender",
        "TargetAgeGroup",
        "TargetInterests",
    ]

    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].fillna("Unknown").astype(str).str.strip()
            df[col] = df[col].replace("", "Unknown")

    if "Cost" not in df.columns:
        df["Cost"] = 0

    if "Impressions" not in df.columns:
        df["Impressions"] = 0

    if "Clicks" not in df.columns:
        if "CTR" in df.columns:
            ctr_value = df["CTR"].copy()

            if ctr_value.dropna().median() > 1:
                ctr_value = ctr_value / 100

            df["Clicks"] = df["Impressions"] * ctr_value
        else:
            df["Clicks"] = 0

    df["Cost"] = df["Cost"].fillna(0)
    df["Impressions"] = df["Impressions"].fillna(0)
    df["Clicks"] = df["Clicks"].fillna(0)

    return df


def format_money(value):
    value = float(value) if pd.notna(value) else 0

    if abs(value) >= 1_000_000:
        return f"RM {value / 1_000_000:.2f}M"

    if abs(value) >= 1_000:
        return f"RM {value / 1_000:.1f}K"

    return f"RM {value:,.2f}"


def format_number(value):
    value = float(value) if pd.notna(value) else 0

    if abs(value) >= 1_000_000:
        return f"{value / 1_000_000:.2f}M"

    if abs(value) >= 1_000:
        return f"{value / 1_000:.1f}K"

    return f"{value:,.0f}"


def format_pct(value):
    value = float(value) if pd.notna(value) else 0
    return f"{value:.2f}%"


def truncate_text(text, max_len=45):
    text = str(text)
    return text if len(text) <= max_len else text[:max_len - 3] + "..."


def create_metric_card(label, value, description):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-desc">{description}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def section_header(title, subtitle):
    st.markdown(
        f"""
        <div class="section-title">{title}</div>
        <div class="section-subtitle">{subtitle}</div>
        """,
        unsafe_allow_html=True
    )


def chart_header(title, subtitle):
    st.markdown(
        f"""
        <div class="chart-title">{title}</div>
        <div class="chart-subtitle">{subtitle}</div>
        """,
        unsafe_allow_html=True
    )


def update_plot_layout(fig, height=390, showlegend=True):
    fig.update_layout(
        height=height,
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(
            family="Inter, Segoe UI, sans-serif",
            size=12,
            color=COLOR_TEXT
        ),
        margin=dict(l=20, r=20, t=20, b=20),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        showlegend=showlegend,
    )

    fig.update_xaxes(
        showgrid=False,
        zeroline=False,
        linecolor=COLOR_BORDER,
        tickfont=dict(color=COLOR_MUTED)
    )

    fig.update_yaxes(
        showgrid=True,
        gridcolor="#EEF2F7",
        zeroline=False,
        linecolor=COLOR_BORDER,
        tickfont=dict(color=COLOR_MUTED)
    )

    return fig


def get_monthly_data(df):
    df = df.copy()

    if "StartDate" in df.columns and df["StartDate"].notna().any():
        df["MonthPeriod"] = df["StartDate"].dt.to_period("M").astype(str)

    elif "Year" in df.columns and "MonthNum" in df.columns:
        df["MonthPeriod"] = (
            df["Year"].astype(int).astype(str)
            + "-"
            + df["MonthNum"].astype(int).astype(str).str.zfill(2)
        )

    elif "Month" in df.columns:
        df["MonthPeriod"] = df["Month"].astype(str)

    else:
        df["MonthPeriod"] = "Unknown"

    monthly = (
        df.groupby("MonthPeriod", as_index=False)
        .agg(
            Cost=("Cost", "sum"),
            Impressions=("Impressions", "sum"),
            Clicks=("Clicks", "sum")
        )
        .sort_values("MonthPeriod")
    )

    monthly["CPM"] = np.where(
        monthly["Impressions"] > 0,
        monthly["Cost"] / monthly["Impressions"] * 1000,
        0
    )

    return monthly


def create_group_summary(df, group_col):
    if group_col not in df.columns:
        return pd.DataFrame()

    grouped = (
        df.groupby(group_col, as_index=False)
        .agg(
            Cost=("Cost", "sum"),
            Impressions=("Impressions", "sum"),
            Clicks=("Clicks", "sum")
        )
    )

    grouped["CPM"] = np.where(
        grouped["Impressions"] > 0,
        grouped["Cost"] / grouped["Impressions"] * 1000,
        0
    )

    grouped["CTR"] = np.where(
        grouped["Impressions"] > 0,
        grouped["Clicks"] / grouped["Impressions"] * 100,
        0
    )

    return grouped


def split_interests(df):
    if "TargetInterests" not in df.columns:
        return pd.DataFrame()

    temp = df[["TargetInterests", "Cost", "Impressions", "Clicks"]].copy()

    temp["TargetInterests"] = (
        temp["TargetInterests"]
        .astype(str)
        .str.replace("|", ",", regex=False)
        .str.replace(";", ",", regex=False)
    )

    temp["Interest"] = temp["TargetInterests"].str.split(",")
    temp = temp.explode("Interest")
    temp["Interest"] = temp["Interest"].astype(str).str.strip()

    temp = temp[
        (temp["Interest"] != "")
        & (temp["Interest"].str.lower() != "nan")
        & (temp["Interest"].str.lower() != "unknown")
    ]

    if temp.empty:
        return pd.DataFrame()

    interest_summary = (
        temp.groupby("Interest", as_index=False)
        .agg(
            Cost=("Cost", "sum"),
            Impressions=("Impressions", "sum"),
            Clicks=("Clicks", "sum")
        )
    )

    interest_summary["CPM"] = np.where(
        interest_summary["Impressions"] > 0,
        interest_summary["Cost"] / interest_summary["Impressions"] * 1000,
        0
    )

    interest_summary["CTR"] = np.where(
        interest_summary["Impressions"] > 0,
        interest_summary["Clicks"] / interest_summary["Impressions"] * 100,
        0
    )

    return interest_summary.sort_values("Impressions", ascending=False).head(10)


def create_campaign_summary(df):
    group_cols = []

    for col in ["CampaignID", "CampaignName", "AdPlatform", "Objective", "AdFormat"]:
        if col in df.columns:
            group_cols.append(col)

    if not group_cols:
        df = df.copy()
        df["CampaignID"] = "Campaign"
        group_cols = ["CampaignID"]

    campaign = (
        df.groupby(group_cols, as_index=False)
        .agg(
            Spend=("Cost", "sum"),
            Impressions=("Impressions", "sum"),
            Clicks=("Clicks", "sum")
        )
    )

    campaign["CPM"] = np.where(
        campaign["Impressions"] > 0,
        campaign["Spend"] / campaign["Impressions"] * 1000,
        0
    )

    campaign["CTR"] = np.where(
        campaign["Impressions"] > 0,
        campaign["Clicks"] / campaign["Impressions"] * 100,
        0
    )

    if "CampaignName" in campaign.columns:
        campaign["CampaignLabel"] = np.where(
            campaign["CampaignName"].astype(str).str.len() > 0,
            campaign["CampaignName"].astype(str),
            campaign["CampaignID"].astype(str)
        )
    else:
        campaign["CampaignLabel"] = campaign["CampaignID"].astype(str)

    campaign["CampaignIDDisplay"] = campaign["CampaignID"].astype(str).apply(
        lambda x: truncate_text(x, 35)
    )

    return campaign


def assign_efficiency_status(row, median_cpm, median_impressions):
    if row["CPM"] <= median_cpm and row["Impressions"] >= median_impressions:
        return "Excellent"

    if row["CPM"] <= median_cpm:
        return "Good"

    return "Review"


# =========================================================
# LOAD DATA
# =========================================================

raw_df = load_data(DATA_FILE)
df = clean_data(raw_df)


# =========================================================
# HEADER
# =========================================================

st.markdown(
    """
    <div class="dashboard-header">
        <div class="dashboard-title">Advertising Campaign Performance Dashboard</div>
        <div class="dashboard-subtitle">
            Data-driven overview of campaign spending, impressions, audience targeting, and advertising efficiency.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)


# =========================================================
# FILTER SECTION
# =========================================================

filtered_df = df.copy()

with st.container(border=True):
    filter_cols = st.columns([1.2, 1.2, 1.4, 1.2, 1.6])

    with filter_cols[0]:
        if "Company" in df.columns:
            company_options = ["All"] + sorted(df["Company"].dropna().unique().tolist())
            selected_company = st.selectbox("Company", company_options)

            if selected_company != "All":
                filtered_df = filtered_df[filtered_df["Company"] == selected_company]
        else:
            st.selectbox("Company", ["All"], disabled=True)

    with filter_cols[1]:
        if "AdPlatform" in df.columns:
            platform_options = ["All"] + sorted(df["AdPlatform"].dropna().unique().tolist())
            selected_platform = st.selectbox("Platform", platform_options)

            if selected_platform != "All":
                filtered_df = filtered_df[filtered_df["AdPlatform"] == selected_platform]
        else:
            st.selectbox("Platform", ["All"], disabled=True)

    with filter_cols[2]:
        if "Objective" in df.columns:
            objective_options = ["All"] + sorted(df["Objective"].dropna().unique().tolist())
            selected_objective = st.selectbox("Objective", objective_options)

            if selected_objective != "All":
                filtered_df = filtered_df[filtered_df["Objective"] == selected_objective]
        else:
            st.selectbox("Objective", ["All"], disabled=True)

    with filter_cols[3]:
        if "AdFormat" in df.columns:
            format_options = ["All"] + sorted(df["AdFormat"].dropna().unique().tolist())
            selected_format = st.selectbox("Ad Format", format_options)

            if selected_format != "All":
                filtered_df = filtered_df[filtered_df["AdFormat"] == selected_format]
        else:
            st.selectbox("Ad Format", ["All"], disabled=True)

    with filter_cols[4]:
        if "StartDate" in df.columns and df["StartDate"].notna().any():
            # Use normalized datetime values to avoid date/time comparison issues.
            valid_start_dates = df["StartDate"].dropna().dt.normalize()

            min_date = valid_start_dates.min().date()
            max_date = valid_start_dates.max().date()

            selected_dates = st.date_input(
                "Date Range",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date
            )

            if isinstance(selected_dates, tuple) and len(selected_dates) == 2:
                start_date, end_date = selected_dates

                start_datetime = pd.to_datetime(start_date).normalize()
                end_datetime = pd.to_datetime(end_date).normalize()

                filtered_df = filtered_df.copy()
                filtered_df["StartDate_Filter"] = filtered_df["StartDate"].dt.normalize()

                filtered_df = filtered_df[
                    filtered_df["StartDate_Filter"].notna()
                    & (filtered_df["StartDate_Filter"] >= start_datetime)
                    & (filtered_df["StartDate_Filter"] <= end_datetime)
                ].drop(columns=["StartDate_Filter"])
        else:
            st.date_input("Date Range", disabled=True)


if filtered_df.empty:
    st.warning("No campaign records match the selected filters.")
    st.stop()


# =========================================================
# KPI CARDS
# =========================================================

total_campaigns = (
    filtered_df["CampaignID"].nunique()
    if "CampaignID" in filtered_df.columns
    else 0
)

total_ads = len(filtered_df)
total_spend = filtered_df["Cost"].sum()
total_impressions = filtered_df["Impressions"].sum()
total_clicks = filtered_df["Clicks"].sum()

average_cpm = total_spend / total_impressions * 1000 if total_impressions > 0 else 0
average_ctr = total_clicks / total_impressions * 100 if total_impressions > 0 else 0

kpi_cols_1 = st.columns(3)

with kpi_cols_1[0]:
    create_metric_card(
        "Total Campaigns",
        f"{total_campaigns:,}",
        "Number of unique advertising campaigns in the selected data."
    )

with kpi_cols_1[1]:
    create_metric_card(
        "Total Ads",
        f"{total_ads:,}",
        "Total number of advertisement records across selected campaigns."
    )

with kpi_cols_1[2]:
    create_metric_card(
        "Total Spend",
        format_money(total_spend),
        "Total advertising cost used across selected campaigns."
    )

st.markdown("<div style='height: 14px;'></div>", unsafe_allow_html=True)

kpi_cols_2 = st.columns(3)

with kpi_cols_2[0]:
    create_metric_card(
        "Total Impressions",
        format_number(total_impressions),
        "Total number of impressions generated by selected campaigns."
    )

with kpi_cols_2[1]:
    create_metric_card(
        "Average CPM",
        f"RM {average_cpm:,.2f}",
        "Cost required to generate one thousand impressions."
    )

with kpi_cols_2[2]:
    create_metric_card(
        "Average CTR",
        format_pct(average_ctr),
        "Percentage of impressions that resulted in clicks."
    )


# =========================================================
# PERFORMANCE OVERVIEW
# =========================================================

section_header(
    "Performance Overview",
    "Monthly view of advertising spending, impressions, and cost efficiency."
)

monthly_df = get_monthly_data(filtered_df)

with st.container(border=True):
    chart_header(
        "Monthly Spend and Impressions Trend",
        "Compares monthly advertising spending with impressions generated over time."
    )

    fig_monthly = make_subplots(specs=[[{"secondary_y": True}]])

    fig_monthly.add_trace(
        go.Bar(
            x=monthly_df["MonthPeriod"],
            y=monthly_df["Cost"],
            name="Total Spend",
            marker_color=COLOR_BLUE,
            hovertemplate="Month: %{x}<br>Spend: RM %{y:,.2f}<extra></extra>"
        ),
        secondary_y=False
    )

    fig_monthly.add_trace(
        go.Scatter(
            x=monthly_df["MonthPeriod"],
            y=monthly_df["Impressions"],
            name="Total Impressions",
            mode="lines+markers",
            line=dict(color=COLOR_TEAL, width=3),
            marker=dict(size=7),
            hovertemplate="Month: %{x}<br>Impressions: %{y:,.0f}<extra></extra>"
        ),
        secondary_y=True
    )

    fig_monthly.update_yaxes(title_text="Total Spend (RM)", secondary_y=False)
    fig_monthly.update_yaxes(title_text="Total Impressions", secondary_y=True)
    fig_monthly.update_xaxes(title_text="Month")

    fig_monthly = update_plot_layout(fig_monthly, height=430, showlegend=True)

    st.plotly_chart(fig_monthly, use_container_width=True)


with st.container(border=True):
    chart_header(
        "Monthly CPM Trend",
        "Shows whether monthly advertising cost efficiency improved or worsened."
    )

    fig_cpm_month = px.line(
        monthly_df,
        x="MonthPeriod",
        y="CPM",
        markers=True,
    )

    fig_cpm_month.update_traces(
        line=dict(color=COLOR_AMBER, width=3),
        marker=dict(size=8),
        hovertemplate="Month: %{x}<br>CPM: RM %{y:,.2f}<extra></extra>"
    )

    fig_cpm_month.update_xaxes(title_text="Month")
    fig_cpm_month.update_yaxes(title_text="CPM (RM)")

    fig_cpm_month = update_plot_layout(fig_cpm_month, height=350, showlegend=False)

    st.plotly_chart(fig_cpm_month, use_container_width=True)


# =========================================================
# PLATFORM AND FORMAT INSIGHTS
# =========================================================

section_header(
    "Platform and Format Insights",
    "Comparison of spending allocation, impressions performance, and cost efficiency across advertising platforms and ad formats."
)

platform_summary = create_group_summary(filtered_df, "AdPlatform")
format_summary = create_group_summary(filtered_df, "AdFormat")

row1_col1, row1_col2 = st.columns(2)

with row1_col1:
    with st.container(border=True):
        chart_header(
            "Spend by Platform",
            "Shows how advertising budget is distributed across platforms."
        )

        if not platform_summary.empty:
            fig_spend_platform = px.pie(
                platform_summary,
                names="AdPlatform",
                values="Cost",
                hole=0.58,
                color_discrete_sequence=PLATFORM_PALETTE
            )

            fig_spend_platform.update_traces(
                textposition="inside",
                textinfo="percent+label",
                hovertemplate=(
                    "Platform: %{label}<br>"
                    "Spend: RM %{value:,.2f}<br>"
                    "Share: %{percent}<extra></extra>"
                )
            )

            fig_spend_platform = update_plot_layout(
                fig_spend_platform,
                height=390,
                showlegend=True
            )

            st.plotly_chart(fig_spend_platform, use_container_width=True)
        else:
            st.info("AdPlatform column not available.")


with row1_col2:
    with st.container(border=True):
        chart_header(
            "Impressions by Platform",
            "Shows which platform generated the highest impressions."
        )

        if not platform_summary.empty:
            platform_impressions = platform_summary.sort_values(
                "Impressions",
                ascending=False
            )

            fig_imp_platform = px.bar(
                platform_impressions,
                x="AdPlatform",
                y="Impressions",
                color="AdPlatform",
                color_discrete_sequence=PLATFORM_PALETTE,
            )

            fig_imp_platform.update_traces(
                hovertemplate="Platform: %{x}<br>Impressions: %{y:,.0f}<extra></extra>"
            )

            fig_imp_platform.update_xaxes(title_text="Platform")
            fig_imp_platform.update_yaxes(title_text="Total Impressions")

            fig_imp_platform = update_plot_layout(
                fig_imp_platform,
                height=390,
                showlegend=False
            )

            st.plotly_chart(fig_imp_platform, use_container_width=True)
        else:
            st.info("AdPlatform column not available.")


row2_col1, row2_col2 = st.columns(2)

with row2_col1:
    with st.container(border=True):
        chart_header(
            "CPM by Platform",
            "Ranks platforms based on cost required to generate one thousand impressions."
        )

        if not platform_summary.empty:
            platform_cpm = platform_summary.sort_values("CPM", ascending=True)

            fig_cpm_platform = px.bar(
                platform_cpm,
                x="CPM",
                y="AdPlatform",
                orientation="h",
                color="CPM",
                color_continuous_scale=BLUE_SCALE,
            )

            fig_cpm_platform.update_traces(
                hovertemplate="Platform: %{y}<br>CPM: RM %{x:,.2f}<extra></extra>"
            )

            fig_cpm_platform.update_layout(coloraxis_showscale=False)
            fig_cpm_platform.update_xaxes(title_text="CPM (RM)")
            fig_cpm_platform.update_yaxes(title_text="Platform", autorange="reversed")

            fig_cpm_platform = update_plot_layout(
                fig_cpm_platform,
                height=390,
                showlegend=False
            )

            st.plotly_chart(fig_cpm_platform, use_container_width=True)
        else:
            st.info("AdPlatform column not available.")


with row2_col2:
    with st.container(border=True):
        chart_header(
            "Impressions by Ad Format",
            "Compares total impressions generated by each ad format."
        )

        if not format_summary.empty:
            format_plot = format_summary.sort_values(
                "Impressions",
                ascending=True
            ).copy()

            fig_format_lollipop = go.Figure()

            for _, row in format_plot.iterrows():
                fig_format_lollipop.add_trace(
                    go.Scatter(
                        x=[0, row["Impressions"]],
                        y=[row["AdFormat"], row["AdFormat"]],
                        mode="lines",
                        line=dict(color=COLOR_LIGHT_TEAL, width=4),
                        showlegend=False,
                        hoverinfo="skip"
                    )
                )

            fig_format_lollipop.add_trace(
                go.Scatter(
                    x=format_plot["Impressions"],
                    y=format_plot["AdFormat"],
                    mode="markers+text",
                    marker=dict(
                        size=16,
                        color=COLOR_TEAL,
                        line=dict(color="white", width=1.5)
                    ),
                    text=[f"{x:,.0f}" for x in format_plot["Impressions"]],
                    textposition="middle right",
                    customdata=np.stack(
                        [format_plot["CPM"], format_plot["CTR"]],
                        axis=-1
                    ),
                    hovertemplate=(
                        "Ad Format: %{y}<br>"
                        "Impressions: %{x:,.0f}<br>"
                        "CPM: RM %{customdata[0]:,.2f}<br>"
                        "CTR: %{customdata[1]:.2f}%<extra></extra>"
                    ),
                    showlegend=False
                )
            )

            format_order = format_plot["AdFormat"].tolist()

            fig_format_lollipop.update_xaxes(title_text="Total Impressions")

            fig_format_lollipop.update_yaxes(
                title_text="Ad Format",
                categoryorder="array",
                categoryarray=format_order
            )

            fig_format_lollipop = update_plot_layout(
                fig_format_lollipop,
                height=390,
                showlegend=False
            )

            st.plotly_chart(fig_format_lollipop, use_container_width=True)
        else:
            st.info("AdFormat column not available.")


# =========================================================
# AUDIENCE TARGETING INSIGHTS
# =========================================================

section_header(
    "Audience Targeting Insights",
    "Analysis of impressions distribution across gender, age group, and audience interests."
)

gender_summary = create_group_summary(filtered_df, "TargetGender")
age_summary = create_group_summary(filtered_df, "TargetAgeGroup")
interest_summary = split_interests(filtered_df)

aud_col1, aud_col2 = st.columns(2)

with aud_col1:
    with st.container(border=True):
        chart_header(
            "Impressions by Target Gender",
            "Shows impressions distribution across gender segments."
        )

        if not gender_summary.empty:
            fig_gender = px.pie(
                gender_summary,
                names="TargetGender",
                values="Impressions",
                hole=0.58,
                color_discrete_sequence=[
                    COLOR_BLUE,
                    COLOR_TEAL,
                    COLOR_PURPLE,
                    COLOR_AMBER,
                    COLOR_GREEN,
                ]
            )

            fig_gender.update_traces(
                textposition="inside",
                textinfo="percent+label",
                hovertemplate=(
                    "Gender: %{label}<br>"
                    "Impressions: %{value:,.0f}<br>"
                    "Share: %{percent}<extra></extra>"
                )
            )

            fig_gender = update_plot_layout(
                fig_gender,
                height=390,
                showlegend=True
            )

            st.plotly_chart(fig_gender, use_container_width=True)
        else:
            st.info("TargetGender column not available.")


with aud_col2:
    with st.container(border=True):
        chart_header(
            "Impressions by Target Age Group",
            "Shows which age group generated the highest impressions."
        )

        if not age_summary.empty:
            age_order = [
                "13-17",
                "18-24",
                "25-34",
                "35-44",
                "45-54",
                "55+",
                "All",
                "Unknown",
            ]

            age_summary["AgeOrder"] = age_summary["TargetAgeGroup"].apply(
                lambda x: age_order.index(x) if x in age_order else 999
            )

            age_summary = age_summary.sort_values(["AgeOrder", "TargetAgeGroup"])

            fig_age = px.bar(
                age_summary,
                x="TargetAgeGroup",
                y="Impressions",
                color_discrete_sequence=[COLOR_TEAL]
            )

            fig_age.update_traces(
                hovertemplate="Age Group: %{x}<br>Impressions: %{y:,.0f}<extra></extra>"
            )

            fig_age.update_xaxes(title_text="Age Group")
            fig_age.update_yaxes(title_text="Total Impressions")

            fig_age = update_plot_layout(
                fig_age,
                height=390,
                showlegend=False
            )

            st.plotly_chart(fig_age, use_container_width=True)
        else:
            st.info("TargetAgeGroup column not available.")


with st.container(border=True):
    chart_header(
        "Top 10 Target Interests by Impressions",
        "Shows which audience interests are associated with the highest impressions."
    )

    if not interest_summary.empty:
        interest_plot = interest_summary.sort_values("Impressions", ascending=True)

        fig_interest = px.bar(
            interest_plot,
            x="Impressions",
            y="Interest",
            orientation="h",
            color="Impressions",
            color_continuous_scale=TEAL_SCALE,
            custom_data=["CPM", "CTR"]
        )

        fig_interest.update_traces(
            hovertemplate=(
                "Interest: %{y}<br>"
                "Impressions: %{x:,.0f}<br>"
                "CPM: RM %{customdata[0]:,.2f}<br>"
                "CTR: %{customdata[1]:.2f}%<extra></extra>"
            )
        )

        fig_interest.update_layout(coloraxis_showscale=False)
        fig_interest.update_xaxes(title_text="Total Impressions")
        fig_interest.update_yaxes(title_text="Target Interest")

        fig_interest = update_plot_layout(
            fig_interest,
            height=450,
            showlegend=False
        )

        st.plotly_chart(fig_interest, use_container_width=True)
    else:
        st.info("TargetInterests column not available.")


# =========================================================
# CAMPAIGN EFFICIENCY
# =========================================================

section_header(
    "Campaign Efficiency",
    "Identifies high-impression campaigns and campaigns with strong cost efficiency."
)

campaign_summary = create_campaign_summary(filtered_df)


# ---------------------------
# Top 15 Campaigns by Impressions
# Fixed ranking order
# ---------------------------

with st.container(border=True):
    chart_header(
        "Top 15 Campaigns by Impressions",
        "Ranks the highest-impression campaigns within the selected filters."
    )

    top_campaigns = campaign_summary.sort_values(
        "Impressions",
        ascending=False
    ).head(15).copy()

    top_campaigns["Rank"] = range(1, len(top_campaigns) + 1)

    top_campaigns["CampaignIDRanked"] = (
        top_campaigns["Rank"].astype(str).str.zfill(2)
        + ". "
        + top_campaigns["CampaignID"].astype(str)
    )

    top_campaigns_plot = top_campaigns.sort_values(
        "Impressions",
        ascending=True
    ).copy()

    campaign_order = top_campaigns_plot["CampaignIDRanked"].tolist()

    color_col = "AdPlatform" if "AdPlatform" in top_campaigns_plot.columns else None

    top_campaigns_plot["HoverText"] = (
        "Rank: " + top_campaigns_plot["Rank"].astype(str)
        + "<br>Campaign ID: " + top_campaigns_plot["CampaignID"].astype(str)
        + "<br>Campaign Name: " + top_campaigns_plot["CampaignName"].astype(str)
        + "<br>Platform: " + top_campaigns_plot["AdPlatform"].astype(str)
        + "<br>Objective: " + top_campaigns_plot["Objective"].astype(str)
        + "<br>Ad Format: " + top_campaigns_plot["AdFormat"].astype(str)
        + "<br>Spend: RM " + top_campaigns_plot["Spend"].map(lambda x: f"{x:,.2f}")
        + "<br>Impressions: " + top_campaigns_plot["Impressions"].map(lambda x: f"{x:,.0f}")
        + "<br>CPM: RM " + top_campaigns_plot["CPM"].map(lambda x: f"{x:,.2f}")
        + "<br>CTR: " + top_campaigns_plot["CTR"].map(lambda x: f"{x:.2f}%")
    )

    fig_top_campaigns = px.bar(
        top_campaigns_plot,
        x="Impressions",
        y="CampaignIDRanked",
        orientation="h",
        color=color_col,
        color_discrete_sequence=PLATFORM_PALETTE,
        custom_data=["HoverText"],
    )

    fig_top_campaigns.update_traces(
        hovertemplate="%{customdata[0]}<extra></extra>"
    )

    fig_top_campaigns.update_xaxes(title_text="Total Impressions")

    fig_top_campaigns.update_yaxes(
        title_text="Campaign ID",
        categoryorder="array",
        categoryarray=campaign_order
    )

    fig_top_campaigns = update_plot_layout(
        fig_top_campaigns,
        height=520,
        showlegend=True
    )

    st.plotly_chart(fig_top_campaigns, use_container_width=True)


# ---------------------------
# Top Efficient Campaigns Table
# ---------------------------

with st.container(border=True):
    chart_header(
        "Top Efficient Campaigns",
        "Highlights campaigns with strong impressions and lower cost per thousand impressions."
    )

    efficient_campaigns = campaign_summary.copy()

    median_cpm = efficient_campaigns["CPM"].median() if not efficient_campaigns.empty else 0
    median_impressions = (
        efficient_campaigns["Impressions"].median()
        if not efficient_campaigns.empty
        else 0
    )

    efficient_campaigns["Efficiency Status"] = efficient_campaigns.apply(
        lambda row: assign_efficiency_status(row, median_cpm, median_impressions),
        axis=1
    )

    efficient_campaigns = efficient_campaigns.sort_values(
        by=["CPM", "Impressions"],
        ascending=[True, False]
    ).head(15)

    display_cols = []

    for col in [
        "CampaignID",
        "CampaignName",
        "AdPlatform",
        "Objective",
        "AdFormat",
        "Spend",
        "Impressions",
        "CPM",
        "CTR",
        "Efficiency Status",
    ]:
        if col in efficient_campaigns.columns:
            display_cols.append(col)

    table_df = efficient_campaigns[display_cols].copy()

    if "Spend" in table_df.columns:
        table_df["Spend"] = table_df["Spend"].apply(lambda x: f"RM {x:,.2f}")

    if "Impressions" in table_df.columns:
        table_df["Impressions"] = table_df["Impressions"].apply(lambda x: f"{x:,.0f}")

    if "CPM" in table_df.columns:
        table_df["CPM"] = table_df["CPM"].apply(lambda x: f"RM {x:,.2f}")

    if "CTR" in table_df.columns:
        table_df["CTR"] = table_df["CTR"].apply(lambda x: f"{x:.2f}%")

    st.dataframe(
        table_df,
        use_container_width=True,
        hide_index=True
    )


# =========================================================
# FOOTER
# =========================================================

st.markdown(
    f"""
    <div style="
        color:{COLOR_MUTED};
        font-size:12.5px;
        text-align:center;
        margin-top:26px;
        padding-bottom:10px;
    ">
        Advertising Campaign Performance Dashboard · Campaign spending, impressions, targeting, and efficiency analytics
    </div>
    """,
    unsafe_allow_html=True
)