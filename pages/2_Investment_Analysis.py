
# =========================================================
# Investment_Analysis.py
# Historical Reinvestment & Lapse Pattern Analysis
# =========================================================

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from pathlib import Path


# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Investment Analysis",
    page_icon="💼",
    layout="wide",
)


# =========================================================
# THEME COLORS
# =========================================================
PRIMARY = "#315EFB"
NAVY = "#0F172A"
TEXT = "#111827"
MUTED = "#64748B"
BORDER = "#E5E7EB"
LIGHT_BG = "#F8FAFC"

GREEN = "#10B981"
BLUE = "#315EFB"
AMBER = "#F59E0B"
ORANGE = "#F97316"
RED = "#EF4444"
GRAY = "#94A3B8"

PATTERN_COLORS = {
    "Growing Investor": GREEN,
    "Stable Investor": BLUE,
    "Declining Investor": ORANGE,
    "New Client": GRAY,
    "Returned After Lapse": AMBER,
    "Currently Lapsed": RED,
    "Returned + Currently Lapsed": RED,
    "No Lapse Pattern": BLUE,
}


# =========================================================
# CSS
# =========================================================
st.markdown(
    """
    <style>
    .block-container {
        padding-top: 3.25rem;
        padding-bottom: 2rem;
        padding-left: 2rem;
        padding-right: 2rem;
        max-width: 100%;
    }

    .page-title {
        margin-top: 0.25rem;
        font-size: 2.2rem;
        font-weight: 800;
        color: #0F172A;
        margin-bottom: 0.15rem;
        line-height: 1.1;
    }

    .page-subtitle {
        font-size: 0.95rem;
        color: #64748B;
        margin-bottom: 1.35rem;
    }

    .section-title {
        font-size: 1.28rem;
        font-weight: 800;
        color: #0F172A;
        margin-top: 0.8rem;
        margin-bottom: 0.1rem;
    }

    .section-subtitle {
        font-size: 0.88rem;
        color: #64748B;
        margin-bottom: 0.9rem;
    }

    .metric-card {
        background: #FFFFFF;
        border: 1px solid #E5E7EB;
        border-radius: 16px;
        padding: 14px 16px 14px 16px;
        min-height: 118px;
        box-sizing: border-box;
    }

    .metric-accent {
        width: 28px;
        height: 3px;
        background: #315EFB;
        border-radius: 999px;
        margin-bottom: 10px;
    }

    .metric-label {
        color: #64748B;
        font-size: 0.82rem;
        font-weight: 600;
        margin-bottom: 6px;
    }

    .metric-value {
        color: #111827;
        font-size: 1.18rem;
        font-weight: 800;
        line-height: 1.25;
        margin-bottom: 8px;
        word-break: break-word;
    }

    .metric-note {
        color: #64748B;
        font-size: 0.76rem;
        line-height: 1.35;
    }

    .chart-title {
        font-size: 1.02rem;
        font-weight: 700;
        color: #0F172A;
        margin-bottom: 0.15rem;
    }

    .chart-subtitle {
        font-size: 0.82rem;
        color: #64748B;
        margin-bottom: 0.55rem;
    }

    .insight-box {
        border-left: 3px solid #315EFB;
        background: #F8FAFC;
        padding: 10px 12px;
        border-radius: 10px;
        color: #475569;
        font-size: 0.82rem;
        margin-top: 0.5rem;
        margin-bottom: 0.2rem;
    }

    div[data-testid="stDataFrame"] {
        border: 1px solid #E5E7EB;
        border-radius: 16px;
        overflow: hidden;
    }

    hr {
        margin-top: 1.3rem !important;
        margin-bottom: 1.15rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# HELPER FUNCTIONS
# =========================================================
def get_project_root() -> Path:
    """Return project root assuming this file is inside /pages."""
    try:
        return Path(__file__).resolve().parents[1]
    except Exception:
        return Path.cwd()


def clean_numeric(series: pd.Series) -> pd.Series:
    """Convert numeric-like strings such as RM 1,000 to float."""
    return pd.to_numeric(
        series.astype(str).str.replace(r"[^\d\.\-]", "", regex=True),
        errors="coerce",
    ).fillna(0)


def parse_campaign_date(series: pd.Series) -> pd.Series:
    """
    Parse campaign dates safely.

    The advertising dataset uses Malaysia/Excel-style dates such as 7/3/2024,
    which should be interpreted as DD/MM/YYYY instead of MM/DD/YYYY.
    The fallbacks keep the page working if some files already contain ISO dates.
    """
    s = series.astype(str).str.strip()

    parsed = pd.to_datetime(s, format="%d/%m/%Y", errors="coerce")
    parsed = parsed.fillna(pd.to_datetime(s, format="%Y-%m-%d", errors="coerce"))
    parsed = parsed.fillna(pd.to_datetime(s, dayfirst=True, errors="coerce"))

    return parsed


def find_column(df: pd.DataFrame, candidates: list[str]):
    """Find a column by case-insensitive matching."""
    lower_map = {str(col).strip().lower(): col for col in df.columns}
    for candidate in candidates:
        key = candidate.strip().lower()
        if key in lower_map:
            return lower_map[key]
    return None


def first_mode(series: pd.Series):
    """Return most frequent value for campaign-level categorical columns."""
    s = series.dropna().astype(str)
    if s.empty:
        return "Unspecified"
    mode_values = s.mode()
    if len(mode_values) > 0:
        return mode_values.iloc[0]
    return s.iloc[0]


@st.cache_data(show_spinner=False)
def load_dataset() -> pd.DataFrame:
    """Load dataset from the project data folder, with dataset/ as fallback."""
    root = get_project_root()

    candidate_paths = [
        root / "data" / "advertising_dataset.csv",
        root / "dataset" / "advertising_dataset.csv",
        root / "data" / "transformed_unified_advertising_dataset.csv",
        root / "dataset" / "transformed_unified_advertising_dataset.csv",
    ]

    for data_path in candidate_paths:
        if data_path.exists():
            return pd.read_csv(data_path)

    searched_paths = "\n".join(str(path) for path in candidate_paths)
    raise FileNotFoundError(
        "Dataset not found. Please ensure your advertising dataset is saved as "
        "advertising_dataset.csv inside either the data or dataset folder.\n\n"
        f"Searched paths:\n{searched_paths}"
    )


def standardize_dataset(raw_df: pd.DataFrame) -> pd.DataFrame:
    """Standardize common column names so the page works with your dataset."""
    df = raw_df.copy()

    aliases = {
        "Company": ["Company", "Client", "ClientName", "Advertiser", "Brand", "Business", "Organization"],
        "CampaignID": ["CampaignID", "Campaign ID", "Campaign_Id", "Campaign", "CampaignName", "Campaign Name"],
        "StartDate": ["StartDate", "Start Date", "CampaignStartDate", "Campaign Start Date", "Start"],
        "EndDate": ["EndDate", "End Date", "CampaignEndDate", "Campaign End Date", "End"],
        "Cost": ["Cost", "Spend", "TotalSpend", "Total Spend", "AmountSpent", "Amount Spent", "AdSpend", "Ad Spend"],
        "Budget": ["Budget", "CampaignBudget", "Campaign Budget", "TotalBudget", "Total Budget"],
        "Impressions": ["Impressions", "Impression", "TotalImpressions", "Total Impressions"],
        "Clicks": ["Clicks", "Click", "TotalClicks", "Total Clicks"],
        "CPM": ["CPM", "AvgCPM", "Average CPM", "CostPerMille", "Cost Per Mille"],
        "Objective": ["Objective", "CampaignObjective", "Campaign Objective"],
        "AdPlatform": ["AdPlatform", "Ad Platform", "Platform", "Channel", "MediaPlatform"],
        "AdFormat": ["AdFormat", "Ad Format", "Format", "CreativeFormat"],
    }

    actual_cols = {std: find_column(df, candidates) for std, candidates in aliases.items()}

    if actual_cols["Company"] is not None:
        df["Company"] = df[actual_cols["Company"]].astype(str).replace({"nan": "Unknown Client"})
    else:
        df["Company"] = "Unknown Client"

    if actual_cols["CampaignID"] is not None:
        df["CampaignID"] = df[actual_cols["CampaignID"]].astype(str).replace({"nan": "Unknown Campaign"})
    else:
        df["CampaignID"] = "Campaign-" + (df.index + 1).astype(str)

    for col in ["Objective", "AdPlatform", "AdFormat"]:
        if actual_cols[col] is not None:
            df[col] = df[actual_cols[col]].astype(str).replace({"nan": "Unspecified"})
        else:
            df[col] = "Unspecified"

    for col in ["Cost", "Budget", "Impressions", "Clicks"]:
        if actual_cols[col] is not None:
            df[col] = clean_numeric(df[actual_cols[col]])
        else:
            df[col] = 0.0

    if actual_cols["StartDate"] is not None:
        df["StartDate"] = parse_campaign_date(df[actual_cols["StartDate"]])
    else:
        df["StartDate"] = pd.NaT

    if actual_cols["EndDate"] is not None:
        df["EndDate"] = parse_campaign_date(df[actual_cols["EndDate"]])
    else:
        df["EndDate"] = pd.NaT

    today_ts = pd.Timestamp.today().normalize()
    df["ActivityDate"] = df["EndDate"].fillna(df["StartDate"]).fillna(today_ts)
    df["StartDate"] = df["StartDate"].fillna(df["ActivityDate"])
    df["EndDate"] = df["EndDate"].fillna(df["ActivityDate"])

    if actual_cols["CPM"] is not None:
        df["CPM"] = clean_numeric(df[actual_cols["CPM"]])
    else:
        df["CPM"] = 0.0

    computed_cpm = np.where(df["Impressions"] > 0, (df["Cost"] / df["Impressions"]) * 1000, 0)
    df["CPM"] = np.where(df["CPM"] > 0, df["CPM"], computed_cpm)

    for col in ["Cost", "Budget", "Impressions", "Clicks", "CPM"]:
        df[col] = df[col].clip(lower=0)

    return df


def build_campaign_level_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convert ad-level records into campaign-level records.
    This prevents one campaign with many ads from being counted as many reinvestments.
    """
    if df.empty:
        return pd.DataFrame()

    campaign_df = (
        df.groupby(["Company", "CampaignID"], as_index=False)
        .agg(
            StartDate=("StartDate", "min"),
            EndDate=("EndDate", "max"),
            ActivityDate=("ActivityDate", "max"),
            Cost=("Cost", "sum"),
            Budget=("Budget", "sum"),
            Impressions=("Impressions", "sum"),
            Clicks=("Clicks", "sum"),
            AdCount=("CampaignID", "count"),
            AdPlatform=("AdPlatform", first_mode),
            Objective=("Objective", first_mode),
            AdFormat=("AdFormat", first_mode),
        )
    )

    campaign_df["CPM"] = np.where(
        campaign_df["Impressions"] > 0,
        (campaign_df["Cost"] / campaign_df["Impressions"]) * 1000,
        0,
    )

    # CampaignID is a reference/registration ID only.
    # Chronological investment behaviour should follow actual campaign dates.
    campaign_df = campaign_df.sort_values(["Company", "StartDate", "EndDate", "CampaignID"]).reset_index(drop=True)
    return campaign_df


def pct_change(current: float, previous: float) -> float:
    if previous <= 0 and current > 0:
        return 100.0
    if previous <= 0 and current <= 0:
        return 0.0
    return ((current - previous) / previous) * 100


def classify_change(change_pct: float) -> str:
    if pd.isna(change_pct):
        return "New Client"
    if change_pct > 10:
        return "Growing Investor"
    if change_pct < -10:
        return "Declining Investor"
    return "Stable Investor"


def format_rm(value: float) -> str:
    return f"RM {value:,.0f}"


def compact_money(value: float) -> str:
    if abs(value) >= 1_000_000:
        return f"RM {value / 1_000_000:.2f}M"
    if abs(value) >= 1_000:
        return f"RM {value / 1_000:.1f}K"
    return f"RM {value:,.0f}"


def style_plotly(fig, height=360, showlegend=True):
    fig.update_layout(
        height=height,
        template="plotly_white",
        paper_bgcolor="white",
        plot_bgcolor="white",
        margin=dict(l=5, r=45, t=20, b=10),
        font=dict(family="Arial, sans-serif", size=12, color=TEXT),
        showlegend=showlegend,
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=11),
        ),
    )
    fig.update_xaxes(
        showgrid=True,
        gridcolor="#EDF2F7",
        zeroline=False,
        linecolor=BORDER,
        tickfont=dict(color=MUTED),
        title_font=dict(color=MUTED, size=12),
    )
    fig.update_yaxes(
        showgrid=False,
        zeroline=False,
        linecolor=BORDER,
        tickfont=dict(color=MUTED),
        title_font=dict(color=MUTED, size=12),
    )
    return fig


def metric_card(title: str, value: str, note: str):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-accent"></div>
            <div class="metric-label">{title}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def insight_box(text: str):
    st.markdown(f'<div class="insight-box">{text}</div>', unsafe_allow_html=True)


def get_period_start(series: pd.Series, granularity: str) -> pd.Series:
    if granularity == "Monthly":
        return series.dt.to_period("M").dt.start_time
    if granularity == "Quarterly":
        return series.dt.to_period("Q").dt.start_time
    return series.dt.to_period("Y").dt.start_time


def build_reinvestment_analysis(campaign_df: pd.DataFrame, reference_date: pd.Timestamp, lapse_threshold: int) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Build two datasets:
    1. company_summary: one row per company
    2. campaign_sequence: one row per campaign with previous-campaign comparison
    """
    if campaign_df.empty:
        return pd.DataFrame(), pd.DataFrame()

    sequence_rows = []

    for company, group in campaign_df.groupby("Company"):
        g = group.sort_values(["StartDate", "EndDate", "CampaignID"]).reset_index(drop=True)

        for idx, row in g.iterrows():
            prev_row = g.iloc[idx - 1] if idx > 0 else None

            previous_spend = float(prev_row["Cost"]) if prev_row is not None else np.nan
            spend_change_rm = float(row["Cost"]) - previous_spend if prev_row is not None else np.nan
            spend_change_pct = pct_change(float(row["Cost"]), previous_spend) if prev_row is not None else np.nan

            raw_gap_days = (
                (row["StartDate"] - prev_row["EndDate"]).days
                if prev_row is not None and pd.notna(row["StartDate"]) and pd.notna(prev_row["EndDate"])
                else np.nan
            )

            # GapDays means inactive days between campaigns.
            # If campaigns overlap, the inactive gap is treated as 0 instead of a negative value.
            gap_days = max(int(raw_gap_days), 0) if pd.notna(raw_gap_days) else np.nan

            sequence_rows.append(
                {
                    "Company": company,
                    "CampaignID": row["CampaignID"],
                    "CampaignOrder": idx + 1,
                    "InvestmentType": "Initial Investment" if idx == 0 else "Reinvestment",
                    "StartDate": row["StartDate"],
                    "EndDate": row["EndDate"],
                    "ActivityDate": row["ActivityDate"],
                    "Cost": float(row["Cost"]),
                    "PreviousCampaignID": prev_row["CampaignID"] if prev_row is not None else np.nan,
                    "PreviousEndDate": prev_row["EndDate"] if prev_row is not None else pd.NaT,
                    "PreviousCampaignSpend": previous_spend,
                    "SpendChangeRM": spend_change_rm,
                    "SpendChangePct": spend_change_pct,
                    "ChangePattern": classify_change(spend_change_pct),
                    "GapDays": gap_days,
                    "ReturnedAfterLapse": bool(pd.notna(gap_days) and gap_days > lapse_threshold),
                    "Impressions": float(row["Impressions"]),
                    "Clicks": float(row["Clicks"]),
                    "CPM": float(row["CPM"]),
                    "AdPlatform": row["AdPlatform"],
                    "Objective": row["Objective"],
                    "AdFormat": row["AdFormat"],
                    "AdCount": int(row["AdCount"]),
                }
            )

    sequence_df = pd.DataFrame(sequence_rows)

    summary_rows = []

    for company, group in sequence_df.groupby("Company"):
        g = group.sort_values("CampaignOrder").reset_index(drop=True)

        total_campaigns = int(len(g))
        reinvestment_count = max(total_campaigns - 1, 0)

        initial_investment = float(g.iloc[0]["Cost"]) if total_campaigns >= 1 else 0.0
        reinvestment_amount = float(g.loc[g["InvestmentType"] == "Reinvestment", "Cost"].sum())
        total_investment = float(g["Cost"].sum())

        latest_campaign = g.iloc[-1]
        previous_campaign = g.iloc[-2] if total_campaigns >= 2 else None

        latest_spend = float(latest_campaign["Cost"])
        previous_spend = float(previous_campaign["Cost"]) if previous_campaign is not None else np.nan
        latest_change_rm = latest_spend - previous_spend if previous_campaign is not None else np.nan
        latest_change_pct = pct_change(latest_spend, previous_spend) if previous_campaign is not None else np.nan

        valid_gaps = g.loc[g["InvestmentType"] == "Reinvestment", "GapDays"].dropna()
        avg_gap_days = float(valid_gaps.mean()) if len(valid_gaps) > 0 else np.nan
        longest_gap_days = float(valid_gaps.max()) if len(valid_gaps) > 0 else np.nan

        last_campaign_date = latest_campaign["EndDate"]
        days_since_last = (reference_date - last_campaign_date).days if pd.notna(last_campaign_date) else 0
        days_since_last = max(int(days_since_last), 0)

        returned_after_lapse = bool(g["ReturnedAfterLapse"].any())
        currently_lapsed = bool(days_since_last > lapse_threshold)

        base_pattern = classify_change(latest_change_pct)
        labels = [base_pattern]

        if returned_after_lapse:
            labels.append("Returned After Lapse")

        if currently_lapsed:
            labels.append("Currently Lapsed")

        client_pattern = " + ".join(labels)

        if returned_after_lapse and currently_lapsed:
            lapse_pattern = "Returned + Currently Lapsed"
        elif returned_after_lapse:
            lapse_pattern = "Returned After Lapse"
        elif currently_lapsed:
            lapse_pattern = "Currently Lapsed"
        else:
            lapse_pattern = "No Lapse Pattern"

        summary_rows.append(
            {
                "Company": company,
                "TotalCampaigns": total_campaigns,
                "ReinvestmentCount": reinvestment_count,
                "InitialInvestmentAmount": initial_investment,
                "TotalReinvestmentAmount": reinvestment_amount,
                "TotalInvestment": total_investment,
                "LatestCampaignSpend": latest_spend,
                "PreviousCampaignSpend": previous_spend,
                "LatestChangeRM": latest_change_rm,
                "LatestChangePct": latest_change_pct,
                "LatestChangePattern": base_pattern,
                "AverageGapDays": avg_gap_days,
                "LongestGapDays": longest_gap_days,
                "ReturnedAfterLapse": returned_after_lapse,
                "CurrentlyLapsed": currently_lapsed,
                "LapsePattern": lapse_pattern,
                "LastCampaignDate": last_campaign_date,
                "DaysSinceLastCampaign": days_since_last,
                "ClientPattern": client_pattern,
                "TotalImpressions": float(g["Impressions"].sum()),
                "AverageCPM": (total_investment / g["Impressions"].sum() * 1000) if g["Impressions"].sum() > 0 else 0,
            }
        )

    company_summary = pd.DataFrame(summary_rows)

    if not company_summary.empty:
        company_summary = company_summary.sort_values(
            by=["ReinvestmentCount", "TotalReinvestmentAmount"],
            ascending=[False, False],
        ).reset_index(drop=True)

    return company_summary, sequence_df


# =========================================================
# LOAD DATA
# =========================================================
try:
    raw_df = load_dataset()
    df = standardize_dataset(raw_df)
except Exception as e:
    st.error("Unable to load the advertising dataset.")
    st.exception(e)
    st.stop()

if df.empty:
    st.error("The advertising dataset is empty.")
    st.stop()


# =========================================================
# PAGE HEADER
# =========================================================
st.markdown('<div class="page-title">Investment Analysis</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="page-subtitle">Historical reinvestment and lapse pattern analysis based on chronological campaign dates, reinvestment amount, campaign-to-campaign spend changes, and inactive gaps.</div>',
    unsafe_allow_html=True,
)


# =========================================================
# FILTERS
# =========================================================
min_date = df["ActivityDate"].min().normalize()
max_date = df["ActivityDate"].max().normalize()

if pd.isna(min_date) or pd.isna(max_date):
    st.error("No valid campaign dates are available for investment analysis.")
    st.stop()

filter_cols = st.columns([1.2, 1.1, 1.0, 1.0, 0.9, 0.9])

with filter_cols[0]:
    date_option = st.selectbox(
        "Date Range",
        ["All Time", "Last 12 Months", "Last 6 Months", "Last 3 Months", "Custom Range"],
        index=0,
    )

with filter_cols[1]:
    company_options = ["All Companies"] + sorted(df["Company"].dropna().astype(str).unique().tolist())
    selected_company = st.selectbox("Company", company_options, index=0)

with filter_cols[2]:
    platform_options = ["All Platforms"] + sorted(df["AdPlatform"].dropna().astype(str).unique().tolist())
    selected_platform = st.selectbox("Platform", platform_options, index=0)

with filter_cols[3]:
    objective_options = ["All Objectives"] + sorted(df["Objective"].dropna().astype(str).unique().tolist())
    selected_objective = st.selectbox("Objective", objective_options, index=0)

with filter_cols[4]:
    granularity = st.selectbox("View By", ["Monthly", "Quarterly", "Yearly"], index=0)

with filter_cols[5]:
    lapse_threshold = st.selectbox("Lapse Gap", [30, 60, 90, 120, 180], index=2)

if date_option == "All Time":
    start_dt, end_dt = min_date, max_date
elif date_option == "Last 12 Months":
    start_dt, end_dt = max(min_date, max_date - pd.DateOffset(months=12)), max_date
elif date_option == "Last 6 Months":
    start_dt, end_dt = max(min_date, max_date - pd.DateOffset(months=6)), max_date
elif date_option == "Last 3 Months":
    start_dt, end_dt = max(min_date, max_date - pd.DateOffset(months=3)), max_date
else:
    custom_range = st.date_input(
        "Custom Date Range",
        value=(min_date.date(), max_date.date()),
        min_value=min_date.date(),
        max_value=max_date.date(),
    )
    if isinstance(custom_range, tuple) and len(custom_range) == 2:
        start_dt, end_dt = pd.Timestamp(custom_range[0]), pd.Timestamp(custom_range[1])
    else:
        start_dt, end_dt = min_date, max_date

filtered_df = df[(df["ActivityDate"] >= start_dt) & (df["ActivityDate"] <= end_dt)].copy()

if selected_company != "All Companies":
    filtered_df = filtered_df[filtered_df["Company"] == selected_company]

if selected_platform != "All Platforms":
    filtered_df = filtered_df[filtered_df["AdPlatform"] == selected_platform]

if selected_objective != "All Objectives":
    filtered_df = filtered_df[filtered_df["Objective"] == selected_objective]

if filtered_df.empty:
    st.warning("No investment data found for the selected filters. Try adjusting the date range or filters.")
    st.stop()

campaign_df = build_campaign_level_df(filtered_df)

if campaign_df.empty:
    st.warning("No campaign-level data is available after applying the selected filters.")
    st.stop()

reference_date = pd.Timestamp(end_dt).normalize()
company_summary, sequence_df = build_reinvestment_analysis(campaign_df, reference_date, lapse_threshold)

if company_summary.empty:
    st.warning("No reinvestment analysis can be generated from the selected data.")
    st.stop()


# =========================================================
# KPI OVERVIEW
# =========================================================
st.markdown('<div class="section-title">Historical Investment Overview</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-subtitle">Summary of company reinvestment behaviour, repeat campaign activity, reinvestment value, and lapse status.</div>',
    unsafe_allow_html=True,
)

total_companies = int(company_summary["Company"].nunique())
repeat_companies = int((company_summary["ReinvestmentCount"] > 0).sum())
reinvestment_campaigns = int(company_summary["ReinvestmentCount"].sum())
total_reinvestment_amount = float(company_summary["TotalReinvestmentAmount"].sum())
avg_reinvestment_count = (
    company_summary.loc[company_summary["ReinvestmentCount"] > 0, "ReinvestmentCount"].mean()
    if repeat_companies > 0
    else 0
)
currently_lapsed_companies = int(company_summary["CurrentlyLapsed"].sum())

kpi_cols = st.columns(6)

with kpi_cols[0]:
    metric_card("Total Companies", f"{total_companies:,}", "Companies available in the selected period")

with kpi_cols[1]:
    metric_card("Repeat Investing Companies", f"{repeat_companies:,}", "Companies with more than one campaign")

with kpi_cols[2]:
    metric_card("Reinvestment Campaigns", f"{reinvestment_campaigns:,}", "Total campaigns excluding first campaign per company")

with kpi_cols[3]:
    metric_card("Total Reinvestment Amount", compact_money(total_reinvestment_amount), "Spend from repeat campaigns only")

with kpi_cols[4]:
    metric_card("Avg Reinvestment Count", f"{avg_reinvestment_count:.1f}", "Average reinvestments among repeat investors")

with kpi_cols[5]:
    metric_card("Currently Lapsed Companies", f"{currently_lapsed_companies:,}", f"No new campaign beyond {lapse_threshold} days")

st.divider()


# =========================================================
# REINVESTMENT FREQUENCY & AMOUNT
# =========================================================
st.markdown('<div class="section-title">Reinvestment Frequency & Amount</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-subtitle">Analyse how many times companies reinvested and how much they spent after their initial campaign.</div>',
    unsafe_allow_html=True,
)

c1, c2 = st.columns(2)

with c1:
    st.markdown('<div class="chart-title">Reinvestment Count by Company</div>', unsafe_allow_html=True)
    st.markdown('<div class="chart-subtitle">Number of repeat campaigns launched after the company’s first campaign.</div>', unsafe_allow_html=True)

    top_count = (
        company_summary.sort_values("ReinvestmentCount", ascending=False)
        .head(10)
        .sort_values("ReinvestmentCount", ascending=True)
    )

    fig_count = go.Figure()
    fig_count.add_trace(
        go.Bar(
            x=top_count["ReinvestmentCount"],
            y=top_count["Company"],
            orientation="h",
            marker=dict(color=BLUE),
            text=top_count["ReinvestmentCount"],
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>Reinvestment Count: %{x}<extra></extra>",
        )
    )
    fig_count.update_xaxes(title="Reinvestment Count")
    fig_count.update_yaxes(title="")
    fig_count = style_plotly(fig_count, height=390, showlegend=False)
    st.plotly_chart(fig_count, use_container_width=True, config={"displayModeBar": False})

    insight_box(
        "<b>Insight:</b> Companies with higher reinvestment counts are repeat clients that continuously purchased additional advertising campaigns."
    )

with c2:
    st.markdown('<div class="chart-title">Total Reinvestment Amount by Company</div>', unsafe_allow_html=True)
    st.markdown('<div class="chart-subtitle">Total spend from campaigns after the initial campaign.</div>', unsafe_allow_html=True)

    top_amount = (
        company_summary.sort_values("TotalReinvestmentAmount", ascending=False)
        .head(10)
        .sort_values("TotalReinvestmentAmount", ascending=True)
    )

    fig_amount = go.Figure()
    fig_amount.add_trace(
        go.Bar(
            x=top_amount["TotalReinvestmentAmount"],
            y=top_amount["Company"],
            orientation="h",
            marker=dict(color="#1F3B69"),
            text=[compact_money(v) for v in top_amount["TotalReinvestmentAmount"]],
            textposition="outside",
            hovertemplate="<b>%{y}</b><br>Total Reinvestment Amount: RM %{x:,.2f}<extra></extra>",
        )
    )
    fig_amount.update_xaxes(title="Total Reinvestment Amount (RM)")
    fig_amount.update_yaxes(title="")
    fig_amount = style_plotly(fig_amount, height=390, showlegend=False)
    st.plotly_chart(fig_amount, use_container_width=True, config={"displayModeBar": False})

    insight_box(
        "<b>Insight:</b> High reinvestment amount indicates stronger historical spending commitment, even if the number of repeat campaigns is not the highest."
    )

st.divider()


# =========================================================
# CAMPAIGN-TO-CAMPAIGN CHANGE
# =========================================================
st.markdown('<div class="section-title">Campaign-to-Campaign Investment Change</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-subtitle">Compare the latest campaign spend with the previous campaign spend to identify growing, stable, or declining investors.</div>',
    unsafe_allow_html=True,
)

change_df = company_summary[company_summary["ReinvestmentCount"] > 0].copy()

if change_df.empty:
    st.info("No repeat campaign is available in the selected data, so campaign-to-campaign change cannot be calculated.")
else:
    change_plot = (
        change_df.reindex(change_df["LatestChangeRM"].abs().sort_values(ascending=False).index)
        .head(15)
        .sort_values("LatestChangeRM", ascending=True)
    )

    colors = [
        GREEN if v > 0 else RED if v < 0 else GRAY
        for v in change_plot["LatestChangeRM"]
    ]

    fig_change = go.Figure()
    fig_change.add_trace(
        go.Bar(
            x=change_plot["LatestChangeRM"],
            y=change_plot["Company"],
            orientation="h",
            marker=dict(color=colors),
            text=[
                f"{'+' if v > 0 else ''}{compact_money(v)}"
                for v in change_plot["LatestChangeRM"]
            ],
            textposition="outside",
            hovertemplate=(
                "<b>%{y}</b><br>"
                "Latest Spend Change: RM %{x:,.2f}<br>"
                "<extra></extra>"
            ),
        )
    )
    fig_change.add_vline(x=0, line_color=GRAY, line_width=1)
    fig_change.update_xaxes(title="Latest Campaign Spend Change vs Previous Campaign (RM)")
    fig_change.update_yaxes(title="")
    fig_change = style_plotly(fig_change, height=440, showlegend=False)
    st.plotly_chart(fig_change, use_container_width=True, config={"displayModeBar": False})

    insight_box(
        "<b>Insight:</b> Positive values show companies that increased their latest campaign investment. Negative values show companies that reduced spending compared with their previous campaign."
    )

st.divider()


# =========================================================
# GAP & LAPSE PATTERN
# =========================================================
st.markdown('<div class="section-title">Reinvestment Gap & Lapse Pattern</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-subtitle">Analyse how long companies waited before reinvesting and whether they returned after an inactive period.</div>',
    unsafe_allow_html=True,
)

c3, c4 = st.columns(2)

with c3:
    st.markdown('<div class="chart-title">Average Inactive Gap Between Campaigns</div>', unsafe_allow_html=True)
    st.markdown('<div class="chart-subtitle">Average inactive days between one campaign ending and the next campaign starting.</div>', unsafe_allow_html=True)

    gap_df = company_summary[company_summary["AverageGapDays"].notna()].copy()

    if gap_df.empty:
        st.info("No campaign gap is available because the selected companies do not have repeat campaigns.")
    else:
        gap_plot = (
            gap_df.sort_values("AverageGapDays", ascending=False)
            .head(10)
            .sort_values("AverageGapDays", ascending=True)
        )

        fig_gap = go.Figure()
        fig_gap.add_trace(
            go.Bar(
                x=gap_plot["AverageGapDays"],
                y=gap_plot["Company"],
                orientation="h",
                marker=dict(color=AMBER),
                text=[f"{v:.0f} days" for v in gap_plot["AverageGapDays"]],
                textposition="outside",
                hovertemplate="<b>%{y}</b><br>Average Inactive Gap: %{x:.0f} days<extra></extra>",
            )
        )
        fig_gap.add_vline(x=lapse_threshold, line_dash="dash", line_color=RED)
        fig_gap.update_xaxes(title="Average Inactive Gap Days")
        fig_gap.update_yaxes(title="")
        fig_gap = style_plotly(fig_gap, height=380, showlegend=False)
        st.plotly_chart(fig_gap, use_container_width=True, config={"displayModeBar": False})

        insight_box(
            f"<b>Insight:</b> Average inactive gaps above the selected <b>{lapse_threshold}-day lapse gap</b> indicate slower reinvestment cycles or possible inactive periods."
        )

with c4:
    st.markdown('<div class="chart-title">Lapse Pattern Distribution</div>', unsafe_allow_html=True)
    st.markdown('<div class="chart-subtitle">Companies grouped by whether they returned after a lapse or are currently lapsed.</div>', unsafe_allow_html=True)

    lapse_order = ["No Lapse Pattern", "Returned After Lapse", "Currently Lapsed", "Returned + Currently Lapsed"]
    lapse_counts = (
        company_summary["LapsePattern"]
        .value_counts()
        .reindex(lapse_order, fill_value=0)
        .reset_index()
    )
    lapse_counts.columns = ["Lapse Pattern", "Companies"]

    fig_lapse = go.Figure()
    fig_lapse.add_trace(
        go.Bar(
            x=lapse_counts["Lapse Pattern"],
            y=lapse_counts["Companies"],
            marker=dict(color=[PATTERN_COLORS.get(x, GRAY) for x in lapse_counts["Lapse Pattern"]]),
            text=lapse_counts["Companies"],
            textposition="outside",
            hovertemplate="<b>%{x}</b><br>Companies: %{y}<extra></extra>",
        )
    )
    fig_lapse.update_xaxes(title="")
    fig_lapse.update_yaxes(title="Number of Companies")
    fig_lapse = style_plotly(fig_lapse, height=380, showlegend=False)
    st.plotly_chart(fig_lapse, use_container_width=True, config={"displayModeBar": False})

    insight_box(
        "<b>Insight:</b> Returned after lapse means the company had a long inactive gap but later reinvested in another campaign."
    )

st.divider()


# =========================================================
# HISTORICAL INVESTMENT TREND
# =========================================================
st.markdown('<div class="section-title">Historical Investment Trend</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-subtitle">Review how company investment changed over time based on the selected monthly, quarterly, or yearly view.</div>',
    unsafe_allow_html=True,
)

trend_df = campaign_df.copy()
trend_df["Period"] = get_period_start(trend_df["ActivityDate"], granularity)

if selected_company == "All Companies":
    top_clients = company_summary.sort_values("TotalInvestment", ascending=False).head(10)["Company"].tolist()
    trend_df = trend_df[trend_df["Company"].isin(top_clients)]
    trend_title_note = "Top 10 companies by total investment"
else:
    trend_title_note = selected_company

trend_group = trend_df.groupby(["Period", "Company"], as_index=False)["Cost"].sum()

fig_trend = px.line(
    trend_group,
    x="Period",
    y="Cost",
    color="Company",
    markers=True,
    title=None,
)

fig_trend.update_traces(
    line=dict(width=2.6),
    marker=dict(size=6),
    hovertemplate="<b>%{fullData.name}</b><br>Period: %{x|%Y-%m-%d}<br>Investment: RM %{y:,.2f}<extra></extra>",
)
fig_trend.update_xaxes(title="Period")
fig_trend.update_yaxes(title="Investment Amount (RM)")
fig_trend = style_plotly(fig_trend, height=430, showlegend=True)
st.plotly_chart(fig_trend, use_container_width=True, config={"displayModeBar": False})

insight_box(
    f"<b>Insight:</b> This trend shows the historical investment pattern for <b>{trend_title_note}</b>. Rising lines suggest increasing advertising commitment, while falling lines suggest reduced spending."
)

st.divider()


# =========================================================
# CLIENT REINVESTMENT TABLE
# =========================================================
st.markdown('<div class="section-title">Client Reinvestment Table</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="section-subtitle">Detailed company-level view of reinvestment count, amount, latest change, inactive gaps, and historical investment pattern.</div>',
    unsafe_allow_html=True,
)

table_df = company_summary.copy()
table_df["Last Campaign Date"] = table_df["LastCampaignDate"].dt.strftime("%Y-%m-%d")

display_table = table_df[
    [
        "Company",
        "TotalCampaigns",
        "ReinvestmentCount",
        "InitialInvestmentAmount",
        "TotalReinvestmentAmount",
        "LatestCampaignSpend",
        "PreviousCampaignSpend",
        "LatestChangeRM",
        "LatestChangePct",
        "AverageGapDays",
        "LongestGapDays",
        "Last Campaign Date",
        "DaysSinceLastCampaign",
        "ClientPattern",
    ]
].rename(
    columns={
        "Company": "Company",
        "TotalCampaigns": "Total Campaigns",
        "ReinvestmentCount": "Reinvestment Count",
        "InitialInvestmentAmount": "Initial Investment Amount",
        "TotalReinvestmentAmount": "Total Reinvestment Amount",
        "LatestCampaignSpend": "Latest Campaign Spend",
        "PreviousCampaignSpend": "Previous Campaign Spend",
        "LatestChangeRM": "Latest Change RM",
        "LatestChangePct": "Latest Change %",
        "AverageGapDays": "Average Inactive Gap Days",
        "LongestGapDays": "Longest Inactive Gap Days",
        "DaysSinceLastCampaign": "Days Since Last Campaign",
        "ClientPattern": "Client Pattern",
    }
)

st.dataframe(
    display_table,
    use_container_width=True,
    hide_index=True,
    height=460,
    column_config={
        "Total Campaigns": st.column_config.NumberColumn("Total Campaigns", format="%d"),
        "Reinvestment Count": st.column_config.NumberColumn("Reinvestment Count", format="%d"),
        "Initial Investment Amount": st.column_config.NumberColumn("Initial Investment Amount", format="RM %.2f"),
        "Total Reinvestment Amount": st.column_config.NumberColumn("Total Reinvestment Amount", format="RM %.2f"),
        "Latest Campaign Spend": st.column_config.NumberColumn("Latest Campaign Spend", format="RM %.2f"),
        "Previous Campaign Spend": st.column_config.NumberColumn("Previous Campaign Spend", format="RM %.2f"),
        "Latest Change RM": st.column_config.NumberColumn("Latest Change RM", format="RM %.2f"),
        "Latest Change %": st.column_config.NumberColumn("Latest Change %", format="%.2f%%"),
        "Average Inactive Gap Days": st.column_config.NumberColumn("Average Inactive Gap Days", format="%.0f"),
        "Longest Inactive Gap Days": st.column_config.NumberColumn("Longest Inactive Gap Days", format="%.0f"),
        "Days Since Last Campaign": st.column_config.NumberColumn("Days Since Last Campaign", format="%d"),
    },
)

insight_box(
    "<b>How to read this table:</b> Reinvestment Count shows how many times the company purchased another campaign after the first one. Latest Change compares the latest campaign spend against the previous campaign. Client Pattern summarises whether the company is growing, stable, declining, returned after lapse, or currently lapsed."
)


# =========================================================
# OPTIONAL CAMPAIGN SEQUENCE TABLE
# =========================================================
with st.expander("View campaign sequence details"):
    sequence_view = sequence_df.copy()
    sequence_view["Start Date"] = sequence_view["StartDate"].dt.strftime("%Y-%m-%d")
    sequence_view["End Date"] = sequence_view["EndDate"].dt.strftime("%Y-%m-%d")
    sequence_view["Previous End Date"] = sequence_view["PreviousEndDate"].dt.strftime("%Y-%m-%d")

    sequence_display = sequence_view[
        [
            "Company",
            "CampaignID",
            "CampaignOrder",
            "InvestmentType",
            "Start Date",
            "End Date",
            "PreviousCampaignID",
            "Previous End Date",
            "Cost",
            "PreviousCampaignSpend",
            "SpendChangeRM",
            "SpendChangePct",
            "GapDays",
            "ReturnedAfterLapse",
            "ChangePattern",
        ]
    ].rename(
        columns={
            "CampaignID": "Campaign ID",
            "CampaignOrder": "Campaign Order",
            "InvestmentType": "Investment Type",
            "PreviousCampaignID": "Previous Campaign ID",
            "Cost": "Campaign Spend",
            "PreviousCampaignSpend": "Previous Campaign Spend",
            "SpendChangeRM": "Spend Change RM",
            "SpendChangePct": "Spend Change %",
            "GapDays": "Inactive Gap Days",
            "ReturnedAfterLapse": "Returned After Lapse",
            "ChangePattern": "Change Pattern",
        }
    )

    st.dataframe(
        sequence_display,
        use_container_width=True,
        hide_index=True,
        height=420,
        column_config={
            "Campaign Order": st.column_config.NumberColumn("Campaign Order", format="%d"),
            "Campaign Spend": st.column_config.NumberColumn("Campaign Spend", format="RM %.2f"),
            "Previous Campaign Spend": st.column_config.NumberColumn("Previous Campaign Spend", format="RM %.2f"),
            "Spend Change RM": st.column_config.NumberColumn("Spend Change RM", format="RM %.2f"),
            "Spend Change %": st.column_config.NumberColumn("Spend Change %", format="%.2f%%"),
            "Inactive Gap Days": st.column_config.NumberColumn("Inactive Gap Days", format="%.0f"),
        },
    )


# =========================================================
# METHODOLOGY
# =========================================================
with st.expander("Methodology"):
    st.markdown(
        f"""
        **Initial Investment** is the first campaign launched by a company in the selected data.

        **Reinvestment** refers to every campaign launched after the first campaign by the same company.

        **Reinvestment Count** is calculated as:

        `Total Campaigns by Company - 1`

        **Total Reinvestment Amount** is the total spend from all repeat campaigns, excluding the first campaign.

        **Campaign sequence** is calculated using actual campaign dates, not CampaignID. CampaignID is treated as a unique registration/reference number only. The dashboard sorts each company’s campaigns by **StartDate → EndDate → CampaignID**. CampaignID is only used as a tie-breaker when two campaigns have the same dates.

        **Latest Change RM and Latest Change %** compare the latest chronological campaign spend with the previous chronological campaign spend.

        **Average Inactive Gap Days** is the average number of inactive days between one campaign ending and the next campaign starting.

        **Longest Inactive Gap Days** is the maximum inactive period found between two consecutive campaigns. If campaigns overlap, the inactive gap is treated as 0 days.

        **Stable Investor** is defined as latest spend change between -10% and +10%.

        **Growing Investor** means latest campaign spend increased by more than 10%.

        **Declining Investor** means latest campaign spend decreased by more than 10%.

        **Returned After Lapse** means the company had an inactive campaign gap greater than the selected lapse gap threshold of **{lapse_threshold} days**, but later launched another campaign.

        **Currently Lapsed** means the company has no campaign after its latest campaign for more than **{lapse_threshold} days** relative to the selected analysis end date.
        """
    )
