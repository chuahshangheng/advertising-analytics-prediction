import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
from pathlib import Path
from html import escape
from textwrap import dedent
from urllib.parse import quote_plus
from io import BytesIO
import base64
import time
from pptx import Presentation

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(page_title="Campaign Ads", layout="wide")


# =========================================================
# DATA PATH
# =========================================================
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_PATH = BASE_DIR / "data" / "advertising_dataset.csv"
AD_TEMPLATE_PATH = BASE_DIR / "template" / "template_ads_deck_placeholders.pptx"
CAMPAIGN_TEMPLATE_PATH = BASE_DIR / "template" / "template_campaign_deck_placeholders.pptx"


# =========================================================
# REQUIRED COLUMNS
# =========================================================
REQUIRED_COLUMNS = [
    "PerformanceID",
    "CampaignID",
    "CampaignName",
    "Company",
    "Impressions",
    "Cost",
    "Budget",
    "AdPlatform",
    "CampaignDurationDays",
    "CPM",
    "StartDate",
]

PLATFORM_ORDER = ["Facebook", "Instagram"]

AD_FORMAT_COLORS = {
    "Image Ad": "#636EFA",
    "Carousel Ad": "#EF553B",
    "Story Ad": "#00CC96",
    "Video Ad": "#AB63FA",
    "Unknown": "#94a3b8",
}


def get_ad_format_color(ad_format: str) -> str:
    ad_format = normalize_display_text(ad_format, fallback="Unknown") if "normalize_display_text" in globals() else str(ad_format)
    return AD_FORMAT_COLORS.get(ad_format, "#64748b")


# =========================================================
# LOAD DATA
# =========================================================
@st.cache_data
def load_ads_data() -> pd.DataFrame:
    if not DATA_PATH.exists():
        raise FileNotFoundError(
            f"Data file not found: {DATA_PATH}\n"
            f"Please place 'unified_advertising_dataset.csv' inside the data folder."
        )

    df = pd.read_csv(DATA_PATH)
    df.columns = [str(col).strip() for col in df.columns]

    missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_cols:
        raise ValueError("Missing required columns: " + ", ".join(missing_cols))

    return df


# =========================================================
# PREPARE DATA
# =========================================================
def parse_campaign_date(series: pd.Series) -> pd.Series:
    """
    Parse campaign dates safely for Malaysia/Excel date format.

    Main expected format: DD/MM/YYYY
    Example: 7/3/2024 = 7 March 2024, not 3 July 2024.
    Fallbacks are included in case some rows are saved as YYYY-MM-DD
    or mixed Excel-style text dates.
    """
    s = series.astype(str).str.strip()

    parsed = pd.to_datetime(s, format="%d/%m/%Y", errors="coerce")
    parsed = parsed.fillna(pd.to_datetime(s, format="%Y-%m-%d", errors="coerce"))
    parsed = parsed.fillna(pd.to_datetime(s, dayfirst=True, errors="coerce"))

    return parsed


def prepare_ads_data(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    numeric_cols = [
        "Impressions",
        "Cost",
        "Budget",
        "CampaignDurationDays",
        "CPM",
    ]

    optional_numeric_cols = ["Clicks", "CTR", "CPC"]
    for col in optional_numeric_cols:
        if col in df.columns:
            numeric_cols.append(col)

    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    text_cols = [
        "PerformanceID",
        "CampaignID",
        "CampaignName",
        "Company",
        "AdPlatform",
    ]

    optional_text_cols = ["AdFormat", "TargetGender", "TargetAgeGroup", "TargetInterests", "Objective"]
    for col in optional_text_cols:
        if col in df.columns:
            text_cols.append(col)

    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Important: parse dates as Malaysia/Excel format first (DD/MM/YYYY).
    # This prevents values such as 7/3/2024 from being read as 3 July 2024.
    df["StartDateParsed"] = parse_campaign_date(df["StartDate"])

    if "EndDate" in df.columns:
        df["EndDateParsed"] = parse_campaign_date(df["EndDate"])
    else:
        df["EndDateParsed"] = pd.NaT

    df["Year"] = df["StartDateParsed"].dt.year

    return df


# =========================================================
# FORMAT HELPERS
# =========================================================
def format_int(value) -> str:
    return f"{int(round(value)):,}"


def format_rm(value) -> str:
    return f"RM {value:,.2f}"


def format_pct(value) -> str:
    return f"{value:.2f}%"


def trim_zero(text: str) -> str:
    return text.rstrip("0").rstrip(".")


def format_compact(value: float) -> str:
    abs_value = abs(value)

    if abs_value >= 1_000_000_000:
        return f"{trim_zero(f'{value / 1_000_000_000:.2f}')}B"
    if abs_value >= 1_000_000:
        return f"{trim_zero(f'{value / 1_000_000:.2f}')}M"
    if abs_value >= 1_000:
        return f"{trim_zero(f'{value / 1_000:.2f}')}K"

    if float(value).is_integer():
        return str(int(value))
    return trim_zero(f"{value:.2f}")


def format_compact_rm(value: float) -> str:
    return f"RM {format_compact(value)}"


def format_ppt_value_label(value: float, value_type: str = "number") -> str:
    """Create short data labels for exported PPT charts."""
    value = float(value) if pd.notna(value) else 0.0

    if value_type == "spend":
        return f"RM {format_compact(value)}"
    if value_type == "impressions":
        return format_compact(value)
    if value_type == "cpm":
        return f"RM {trim_zero(f'{value:.2f}')}"

    return format_compact(value)


def format_date_for_display(value) -> str:
    if pd.isna(value):
        return ""
    return pd.to_datetime(value).strftime("%d/%m/%Y")


def format_platforms(series: pd.Series) -> str:
    values = []
    for item in series.dropna().astype(str):
        item = item.strip()
        if item and item not in values:
            values.append(item)

    ordered = [p for p in PLATFORM_ORDER if p in values]
    others = [p for p in values if p not in ordered]
    return ", ".join(ordered + others)


def first_non_empty_text(series: pd.Series) -> str:
    values = []
    for item in series.dropna().astype(str):
        item = item.strip()
        if item and item.lower() != "nan":
            values.append(item)
    return values[0] if values else "-"


def format_unique_text_values(series: pd.Series, max_items: int = 6) -> str:
    values = []

    for item in series.dropna().astype(str):
        clean_item = item.strip()
        if not clean_item or clean_item.lower() == "nan":
            continue

        split_items = (
            clean_item
            .replace(";", ",")
            .replace("|", ",")
            .split(",")
        )

        for part in split_items:
            clean_part = part.strip()
            if clean_part and clean_part.lower() != "nan" and clean_part not in values:
                values.append(clean_part)

    if not values:
        return "-"

    if len(values) > max_items:
        shown_values = values[:max_items]
        remaining_count = len(values) - max_items
        return ", ".join(shown_values) + f" +{remaining_count} more"

    return ", ".join(values)


def format_budget_usage_subtitle(budget_value: float, used_pct: float, fallback: str = "") -> str:
    if pd.notna(budget_value) and float(budget_value) > 0:
        return f"of {format_compact_rm(float(budget_value))} budget • {float(used_pct):.1f}% used"
    return fallback


def normalize_display_text(value, fallback: str = "-") -> str:
    if pd.isna(value):
        return fallback
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return fallback
    return text


def truncate_text(text: str, max_len: int = 26) -> str:
    text = normalize_display_text(text, fallback="-")
    if len(text) <= max_len:
        return text
    return text[: max_len - 1].rstrip() + "…"


def join_meta_parts(parts) -> str:
    valid_parts = []
    for part in parts:
        clean = normalize_display_text(part, fallback="")
        if clean:
            valid_parts.append(clean)
    return " • ".join(valid_parts) if valid_parts else "-"


# =========================================================
# FILTER HELPERS
# =========================================================
def get_year_options(df: pd.DataFrame):
    years = (
        df["Year"]
        .dropna()
        .astype(int)
        .sort_values()
        .unique()
        .tolist()
    )
    return ["All"] + [str(year) for year in years]


def apply_year_filter(df: pd.DataFrame, selected_year: str) -> pd.DataFrame:
    if selected_year == "All":
        return df.copy()
    return df[df["Year"].astype("Int64") == int(selected_year)].copy()


def render_year_filter(year_options):
    if hasattr(st, "segmented_control"):
        return st.segmented_control(
            "Year",
            options=year_options,
            key="ads_year_filter",
            label_visibility="collapsed",
        )

    if hasattr(st, "pills"):
        return st.pills(
            "Year",
            options=year_options,
            key="ads_year_filter",
            label_visibility="collapsed",
        )

    return st.radio(
        "Year",
        options=year_options,
        horizontal=True,
        key="ads_year_filter",
        label_visibility="collapsed",
    )


# =========================================================
# KPI CALCULATIONS
# =========================================================
def calculate_kpis(df: pd.DataFrame) -> dict:
    total_campaigns = df["CampaignID"].nunique()
    total_companies = df["Company"].nunique()
    total_ads = df["PerformanceID"].nunique()
    total_spend = df["Cost"].sum()
    total_impressions = df["Impressions"].sum()

    overall_cpm = 0.0
    if total_impressions > 0:
        overall_cpm = (total_spend / total_impressions) * 1000

    campaign_budget_df = (
        df.groupby("CampaignID", as_index=False)
        .agg(CampaignBudget=("Budget", "mean"))
    )

    total_budget = campaign_budget_df["CampaignBudget"].sum()

    spend_pct = 0.0
    if total_budget > 0:
        spend_pct = (total_spend / total_budget) * 100

    return {
        "total_campaigns": total_campaigns,
        "total_companies": total_companies,
        "total_ads": total_ads,
        "total_spend": total_spend,
        "total_impressions": total_impressions,
        "overall_cpm": overall_cpm,
        "total_budget": total_budget,
        "spend_pct": spend_pct,
    }


# =========================================================
# TREND DATA PREPARATION
# =========================================================
def prepare_trend_data(df: pd.DataFrame, value_col: str, selected_year: str) -> pd.DataFrame:
    trend_df = df.copy()
    trend_df = trend_df.dropna(subset=["StartDateParsed"])

    if trend_df.empty:
        if selected_year == "All":
            return pd.DataFrame(columns=["XKey", "Label", "YearLabel", "Total", "Facebook", "Instagram"])
        return pd.DataFrame({
            "XKey": [str(i) for i in range(1, 13)],
            "Label": ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
            "YearLabel": [str(selected_year)] * 12,
            "Total": [0] * 12,
            "Facebook": [0] * 12,
            "Instagram": [0] * 12,
        })

    if selected_year == "All":
        trend_df["Period"] = trend_df["StartDateParsed"].dt.to_period("M")

        period_range = pd.period_range(
            start=trend_df["Period"].min(),
            end=trend_df["Period"].max(),
            freq="M"
        )

        base_df = pd.DataFrame({"Period": period_range})
        base_df["XKey"] = base_df["Period"].astype(str)
        base_df["Label"] = base_df["Period"].dt.strftime("%b")
        base_df["YearLabel"] = base_df["Period"].dt.strftime("%Y")

        total_df = (
            trend_df.groupby("Period", as_index=False)[value_col]
            .sum()
            .rename(columns={value_col: "Total"})
        )

        platform_df = (
            trend_df.groupby(["Period", "AdPlatform"], as_index=False)[value_col]
            .sum()
            .pivot(index="Period", columns="AdPlatform", values=value_col)
            .reset_index()
        )

        merged_df = base_df.merge(total_df, on="Period", how="left")
        merged_df = merged_df.merge(platform_df, on="Period", how="left")

    else:
        selected_year_int = int(selected_year)

        base_df = pd.DataFrame({"MonthNum": list(range(1, 13))})
        base_df["XKey"] = base_df["MonthNum"].astype(str)
        base_df["Label"] = pd.to_datetime(base_df["MonthNum"], format="%m").dt.strftime("%b")
        base_df["YearLabel"] = str(selected_year_int)

        trend_df["MonthNum"] = trend_df["StartDateParsed"].dt.month

        total_df = (
            trend_df.groupby("MonthNum", as_index=False)[value_col]
            .sum()
            .rename(columns={value_col: "Total"})
        )

        platform_df = (
            trend_df.groupby(["MonthNum", "AdPlatform"], as_index=False)[value_col]
            .sum()
            .pivot(index="MonthNum", columns="AdPlatform", values=value_col)
            .reset_index()
        )

        merged_df = base_df.merge(total_df, on="MonthNum", how="left")
        merged_df = merged_df.merge(platform_df, on="MonthNum", how="left")

    merged_df["Total"] = merged_df["Total"].fillna(0)

    for platform in PLATFORM_ORDER:
        if platform not in merged_df.columns:
            merged_df[platform] = 0
        merged_df[platform] = merged_df[platform].fillna(0)

    return merged_df


# =========================================================
# CHART HELPERS
# =========================================================
def add_all_year_custom_axis(fig, trend_df: pd.DataFrame):
    total_points = len(trend_df)
    if total_points == 0:
        return fig

    for idx, row in trend_df.reset_index(drop=True).iterrows():
        month_x = (idx + 0.5) / total_points

        fig.add_annotation(
            x=month_x,
            y=-0.045,
            xref="paper",
            yref="paper",
            text=str(row["Label"]),
            showarrow=False,
            font=dict(size=11, color="#64748b"),
            align="center",
        )

    grouped_years = (
        trend_df[["YearLabel"]]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    start_idx = 0
    for i, year_label in enumerate(grouped_years["YearLabel"].tolist()):
        year_count = (trend_df["YearLabel"] == year_label).sum()
        end_idx = start_idx + year_count - 1

        year_x = (start_idx + (year_count / 2)) / total_points

        fig.add_annotation(
            x=year_x,
            y=-0.105,
            xref="paper",
            yref="paper",
            text=f"<b>{year_label}</b>",
            showarrow=False,
            font=dict(size=14, color="#475569"),
            align="center",
        )

        if i < len(grouped_years) - 1:
            separator_x = (end_idx + 1) / total_points
            fig.add_shape(
                type="line",
                xref="paper",
                yref="paper",
                x0=separator_x,
                x1=separator_x,
                y0=-0.095,
                y1=-0.005,
                line=dict(color="rgba(148, 163, 184, 0.45)", width=1),
            )

        start_idx += year_count

    return fig


def build_trend_chart(trend_df: pd.DataFrame, title: str, value_type: str, selected_year: str):
    fig = go.Figure()

    x_values = trend_df["XKey"]
    hover_text = trend_df["Label"] + " " + trend_df["YearLabel"]

    fig.add_trace(
        go.Scatter(
            x=x_values,
            y=trend_df["Total"],
            mode="lines+markers",
            name="Total",
            line=dict(width=3),
            marker=dict(size=5),
            hovertemplate=(
                "<b>%{text}</b><br>"
                + ("Total Spend: RM %{y:,.2f}" if value_type == "spend" else "Total Impressions: %{y:,.0f}")
                + "<extra></extra>"
            ),
            text=hover_text,
        )
    )

    fig.add_trace(
        go.Scatter(
            x=x_values,
            y=trend_df["Facebook"],
            mode="lines+markers",
            name="Facebook",
            line=dict(width=2),
            marker=dict(size=4),
            hovertemplate=(
                "<b>%{text}</b><br>"
                + ("Facebook Spend: RM %{y:,.2f}" if value_type == "spend" else "Facebook Impressions: %{y:,.0f}")
                + "<extra></extra>"
            ),
            text=hover_text,
        )
    )

    fig.add_trace(
        go.Scatter(
            x=x_values,
            y=trend_df["Instagram"],
            mode="lines+markers",
            name="Instagram",
            line=dict(width=2),
            marker=dict(size=4),
            hovertemplate=(
                "<b>%{text}</b><br>"
                + ("Instagram Spend: RM %{y:,.2f}" if value_type == "spend" else "Instagram Impressions: %{y:,.0f}")
                + "<extra></extra>"
            ),
            text=hover_text,
        )
    )

    bottom_margin = 18
    if selected_year == "All":
        bottom_margin = 58

    fig.update_layout(
        title=title,
        template="plotly_white",
        height=340,
        margin=dict(l=20, r=20, t=45, b=bottom_margin),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        hovermode="x unified",
    )

    if selected_year == "All":
        fig.update_xaxes(
            title_text="",
            showticklabels=False,
            showgrid=False,
            ticks="",
            showline=False
        )
        fig = add_all_year_custom_axis(fig, trend_df)
    else:
        fig.update_xaxes(
            title_text="",
            tickmode="array",
            tickvals=x_values,
            ticktext=trend_df["Label"],
            showgrid=False,
            tickangle=0,
            tickfont=dict(size=11, color="#64748b")
        )

    if value_type == "spend":
        fig.update_yaxes(
            title_text="Spend (RM)",
            tickprefix="RM ",
            tickformat="~s",
            showgrid=True,
            gridcolor="rgba(148, 163, 184, 0.20)"
        )
    else:
        fig.update_yaxes(
            title_text="Impressions",
            tickformat="~s",
            showgrid=True,
            gridcolor="rgba(148, 163, 184, 0.20)"
        )

    return fig


def build_empty_chart(title: str, message: str = "No data available"):
    fig = go.Figure()
    fig.update_layout(
        title=title,
        template="plotly_white",
        height=320,
        margin=dict(l=20, r=20, t=45, b=20),
        xaxis=dict(visible=False),
        yaxis=dict(visible=False),
        annotations=[
            dict(
                text=message,
                x=0.5,
                y=0.5,
                xref="paper",
                yref="paper",
                showarrow=False,
                font=dict(size=14, color="#64748b"),
            )
        ],
    )
    return fig


def prepare_ad_format_chart_data(selected_df: pd.DataFrame) -> pd.DataFrame:
    if selected_df.empty:
        return pd.DataFrame(columns=[
            "AdFormat", "TotalSpend", "TotalImpressions", "TotalClicks",
            "AdsCount", "WeightedCPM", "WeightedCTR", "WeightedCPC"
        ])

    chart_df = selected_df.copy()

    if "AdFormat" not in chart_df.columns:
        chart_df["AdFormat"] = "Unknown"

    if "Clicks" not in chart_df.columns:
        chart_df["Clicks"] = 0

    chart_df["AdFormat"] = (
        chart_df["AdFormat"]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("", "Unknown")
    )

    summary_df = (
        chart_df.groupby("AdFormat", as_index=False)
        .agg(
            TotalSpend=("Cost", "sum"),
            TotalImpressions=("Impressions", "sum"),
            TotalClicks=("Clicks", "sum"),
            AdsCount=("PerformanceID", "nunique"),
        )
    )

    summary_df["WeightedCPM"] = 0.0
    valid_impr = summary_df["TotalImpressions"] > 0
    summary_df.loc[valid_impr, "WeightedCPM"] = (
        summary_df.loc[valid_impr, "TotalSpend"] / summary_df.loc[valid_impr, "TotalImpressions"] * 1000
    )

    summary_df["WeightedCTR"] = 0.0
    summary_df.loc[valid_impr, "WeightedCTR"] = (
        summary_df.loc[valid_impr, "TotalClicks"] / summary_df.loc[valid_impr, "TotalImpressions"] * 100
    )

    summary_df["WeightedCPC"] = 0.0
    valid_clicks = summary_df["TotalClicks"] > 0
    summary_df.loc[valid_clicks, "WeightedCPC"] = (
        summary_df.loc[valid_clicks, "TotalSpend"] / summary_df.loc[valid_clicks, "TotalClicks"]
    )

    return summary_df


def prepare_top_ads_chart_data(selected_df: pd.DataFrame, top_n: int = 10, sort_by: str = "Impressions") -> pd.DataFrame:
    if selected_df.empty:
        return pd.DataFrame(columns=["PerformanceID", "Impressions", "Cost", "AdFormat", "AdPlatform"])

    chart_df = selected_df.copy()

    if "AdFormat" not in chart_df.columns:
        chart_df["AdFormat"] = "Unknown"

    if "AdPlatform" not in chart_df.columns:
        chart_df["AdPlatform"] = "-"

    chart_df["AdFormat"] = (
        chart_df["AdFormat"]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("", "Unknown")
    )

    chart_df["AdPlatform"] = (
        chart_df["AdPlatform"]
        .fillna("")
        .astype(str)
        .str.strip()
        .replace("", "-")
    )

    ranking_df = (
        chart_df.groupby(["PerformanceID", "AdFormat", "AdPlatform"], as_index=False)
        .agg(
            Impressions=("Impressions", "sum"),
            Cost=("Cost", "sum"),
        )
    )

    sort_by = sort_by if sort_by in ["Impressions", "Cost"] else "Impressions"
    secondary_col = "Cost" if sort_by == "Impressions" else "Impressions"

    ranking_df = (
        ranking_df
        .sort_values(by=[sort_by, secondary_col], ascending=[False, False])
        .head(top_n)
        .reset_index(drop=True)
    )

    return ranking_df


def build_ad_format_bar_chart(
    chart_df: pd.DataFrame,
    value_col: str,
    title: str,
    value_type: str,
    show_values: bool = False,
):
    if chart_df.empty:
        return build_empty_chart(title)

    plot_df = chart_df.copy().sort_values(by=value_col, ascending=False).reset_index(drop=True)

    hover_template_map = {
        "spend": (
            "<b>%{x}</b><br>"
            "Total Spend: RM %{y:,.2f}<br>"
            "Impressions: %{customdata[0]:,.0f}<br>"
            "Weighted CPM: RM %{customdata[1]:,.2f}<extra></extra>"
        ),
        "impressions": (
            "<b>%{x}</b><br>"
            "Impressions: %{y:,.0f}<br>"
            "Total Spend: RM %{customdata[0]:,.2f}<br>"
            "Weighted CPM: RM %{customdata[1]:,.2f}<extra></extra>"
        ),
        "cpm": (
            "<b>%{x}</b><br>"
            "Weighted CPM: RM %{y:,.2f}<br>"
            "Total Spend: RM %{customdata[0]:,.2f}<br>"
            "Impressions: %{customdata[1]:,.0f}<extra></extra>"
        ),
    }

    if value_col == "TotalSpend":
        customdata = plot_df[["TotalImpressions", "WeightedCPM"]]
    elif value_col == "TotalImpressions":
        customdata = plot_df[["TotalSpend", "WeightedCPM"]]
    else:
        customdata = plot_df[["TotalSpend", "TotalImpressions"]]

    label_value_type = "spend" if value_col == "TotalSpend" else value_type
    value_labels = plot_df[value_col].apply(lambda x: format_ppt_value_label(x, label_value_type))

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=plot_df["AdFormat"],
            y=plot_df[value_col],
            customdata=customdata,
            hovertemplate=hover_template_map[value_type],
            text=value_labels if show_values else None,
            textposition="outside" if show_values else None,
            textfont=dict(size=12, color="#334155") if show_values else None,
            cliponaxis=False,
        )
    )

    fig.update_layout(
        title=title,
        template="plotly_white",
        height=320,
        margin=dict(l=20, r=20, t=45, b=20),
        showlegend=False,
        bargap=0.35,
    )

    fig.update_xaxes(
        title_text="Ad Format",
        showgrid=False,
        tickangle=0,
    )

    max_value = float(plot_df[value_col].max()) if not plot_df.empty else 0.0
    y_axis_range = [0, max_value * 1.18] if show_values and max_value > 0 else None

    if value_type == "spend":
        fig.update_yaxes(
            title_text="Spend (RM)",
            tickprefix="RM ",
            tickformat="~s",
            showgrid=True,
            gridcolor="rgba(148, 163, 184, 0.20)",
            range=y_axis_range,
        )
    elif value_type == "impressions":
        fig.update_yaxes(
            title_text="Impressions",
            tickformat="~s",
            showgrid=True,
            gridcolor="rgba(148, 163, 184, 0.20)",
            range=y_axis_range,
        )
    else:
        fig.update_yaxes(
            title_text="CPM (RM)",
            tickprefix="RM ",
            tickformat=",.2f",
            showgrid=True,
            gridcolor="rgba(148, 163, 184, 0.20)",
            range=y_axis_range,
        )

    return fig

def build_top_ads_chart(
    ranking_df: pd.DataFrame,
    title: str = "Top 10 Ads by Impressions",
    metric_col: str = "Impressions",
    value_title: str = "Impressions",
    value_prefix: str = "",
    show_values: bool = False,
):
    if ranking_df.empty:
        return build_empty_chart(title)

    metric_col = metric_col if metric_col in ["Impressions", "Cost"] else "Impressions"
    secondary_col = "Cost" if metric_col == "Impressions" else "Impressions"

    plot_df = (
        ranking_df
        .copy()
        .sort_values(by=[metric_col, secondary_col], ascending=[False, False])
        .reset_index(drop=True)
    )

    marker_colors = [get_ad_format_color(fmt) for fmt in plot_df["AdFormat"]]

    if metric_col == "Impressions":
        customdata = plot_df[["Cost", "AdFormat", "AdPlatform"]]
        hovertemplate = (
            "<b>%{y}</b><br>"
            "Impressions: %{x:,.0f}<br>"
            "Spend: RM %{customdata[0]:,.2f}<br>"
            "Ad Format: %{customdata[1]}<br>"
            "Platform: %{customdata[2]}<extra></extra>"
        )
        label_value_type = "impressions"
    else:
        customdata = plot_df[["Impressions", "AdFormat", "AdPlatform"]]
        hovertemplate = (
            "<b>%{y}</b><br>"
            "Spend: RM %{x:,.2f}<br>"
            "Impressions: %{customdata[0]:,.0f}<br>"
            "Ad Format: %{customdata[1]}<br>"
            "Platform: %{customdata[2]}<extra></extra>"
        )
        label_value_type = "spend"

    value_labels = plot_df[metric_col].apply(lambda x: format_ppt_value_label(x, label_value_type))

    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=plot_df[metric_col],
            y=plot_df["PerformanceID"],
            orientation="h",
            marker=dict(color=marker_colors),
            customdata=customdata,
            hovertemplate=hovertemplate,
            showlegend=False,
            text=value_labels if show_values else None,
            textposition="outside" if show_values else None,
            textfont=dict(size=12, color="#334155") if show_values else None,
            cliponaxis=False,
        )
    )

    for ad_format in plot_df["AdFormat"].drop_duplicates().tolist():
        fig.add_trace(
            go.Bar(
                x=[None],
                y=[None],
                name=ad_format,
                marker=dict(color=get_ad_format_color(ad_format)),
                showlegend=True,
                hoverinfo="skip",
            )
        )

    fig.update_layout(
        title=title,
        template="plotly_white",
        height=320,
        margin=dict(l=20, r=20, t=45, b=20),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
            font=dict(size=10),
        ),
        bargap=0.30,
    )

    max_value = float(plot_df[metric_col].max()) if not plot_df.empty else 0.0
    x_axis_range = [0, max_value * 1.18] if show_values and max_value > 0 else None

    if metric_col == "Cost":
        fig.update_xaxes(
            title_text=value_title,
            tickprefix="RM ",
            tickformat="~s",
            showgrid=True,
            gridcolor="rgba(148, 163, 184, 0.20)",
            range=x_axis_range,
        )
    else:
        fig.update_xaxes(
            title_text=value_title,
            tickformat="~s",
            showgrid=True,
            gridcolor="rgba(148, 163, 184, 0.20)",
            range=x_axis_range,
        )

    fig.update_yaxes(
        title_text="",
        autorange="reversed",
        showgrid=False,
    )

    return fig


# =========================================================
# CAMPAIGN LIST PREPARATION
# =========================================================
def prepare_campaign_list(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=[
            "Campaign ID", "CampaignName", "Company", "AdPlatform", "Objective",
            "StartDateParsed", "EndDateParsed", "CampaignDurationDays",
            "TotalAds", "TotalSpend", "CampaignBudget", "BudgetUsedPct",
            "TotalImpressions", "AverageCPM", "Total Spend Text",
            "Campaign Budget Text", "Budget Usage Text", "Budget Used Text",
            "Total Impressions Text", "Average CPM Text", "Total Ads Text",
            "Objective Text", "Target Gender Text", "Target Age Group Text", "Target Interests Text",
            "Meta Line 1", "Meta Line 2"
        ])

    agg_dict = {
        "AdPlatform": ("AdPlatform", format_platforms),
        "StartDateParsed": ("StartDateParsed", "min"),
        "EndDateParsed": ("EndDateParsed", "max"),
        "CampaignDurationDays": ("CampaignDurationDays", "max"),
        "TotalAds": ("PerformanceID", "nunique"),
        "TotalSpend": ("Cost", "sum"),
        "CampaignBudget": ("Budget", "mean"),
        "TotalImpressions": ("Impressions", "sum"),
    }

    if "Objective" in df.columns:
        agg_dict["Objective"] = ("Objective", first_non_empty_text)

    if "TargetGender" in df.columns:
        agg_dict["TargetGender"] = ("TargetGender", format_unique_text_values)

    if "TargetAgeGroup" in df.columns:
        agg_dict["TargetAgeGroup"] = ("TargetAgeGroup", format_unique_text_values)

    if "TargetInterests" in df.columns:
        agg_dict["TargetInterests"] = ("TargetInterests", format_unique_text_values)

    campaign_df = (
        df.groupby(["CampaignID", "CampaignName", "Company"], as_index=False)
        .agg(**agg_dict)
    )

    if "Objective" not in campaign_df.columns:
        campaign_df["Objective"] = "-"

    if "TargetGender" not in campaign_df.columns:
        campaign_df["TargetGender"] = "-"

    if "TargetAgeGroup" not in campaign_df.columns:
        campaign_df["TargetAgeGroup"] = "-"

    if "TargetInterests" not in campaign_df.columns:
        campaign_df["TargetInterests"] = "-"

    derived_end = campaign_df["StartDateParsed"] + pd.to_timedelta(
        campaign_df["CampaignDurationDays"].clip(lower=1) - 1,
        unit="D"
    )
    campaign_df["EndDateParsed"] = campaign_df["EndDateParsed"].fillna(derived_end)

    valid_dates = campaign_df["StartDateParsed"].notna() & campaign_df["EndDateParsed"].notna()
    campaign_df.loc[valid_dates, "CampaignDurationDays"] = (
        (campaign_df.loc[valid_dates, "EndDateParsed"] - campaign_df.loc[valid_dates, "StartDateParsed"]).dt.days + 1
    )

    campaign_df["AverageCPM"] = 0.0
    valid_impressions = campaign_df["TotalImpressions"] > 0
    campaign_df.loc[valid_impressions, "AverageCPM"] = (
        campaign_df.loc[valid_impressions, "TotalSpend"] / campaign_df.loc[valid_impressions, "TotalImpressions"] * 1000
    )

    campaign_df["BudgetUsedPct"] = 0.0
    valid_budget = campaign_df["CampaignBudget"] > 0
    campaign_df.loc[valid_budget, "BudgetUsedPct"] = (
        campaign_df.loc[valid_budget, "TotalSpend"] / campaign_df.loc[valid_budget, "CampaignBudget"] * 100
    )

    campaign_df = campaign_df.sort_values(
        by=["StartDateParsed", "TotalSpend"],
        ascending=[False, False]
    ).reset_index(drop=True)

    campaign_df["Campaign ID"] = campaign_df["CampaignID"]
    campaign_df["Start Date"] = campaign_df["StartDateParsed"].apply(format_date_for_display)
    campaign_df["End Date"] = campaign_df["EndDateParsed"].apply(format_date_for_display)
    campaign_df["Campaign Duration (Days)"] = campaign_df["CampaignDurationDays"].apply(
        lambda x: format_int(x) if pd.notna(x) else ""
    )
    campaign_df["Total Ads Text"] = campaign_df["TotalAds"].apply(
        lambda x: f"{format_int(x)} ad" if int(round(x)) == 1 else f"{format_int(x)} ads"
    )
    campaign_df["Total Spend Text"] = campaign_df["TotalSpend"].apply(format_rm)
    campaign_df["Campaign Budget Text"] = campaign_df["CampaignBudget"].apply(format_rm)
    campaign_df["Budget Usage Text"] = campaign_df.apply(
        lambda row: f"{row['Total Spend Text']} / {row['Campaign Budget Text']}" if row["CampaignBudget"] > 0 else row["Total Spend Text"],
        axis=1
    )
    campaign_df["Budget Used Text"] = campaign_df["BudgetUsedPct"].apply(lambda x: f"{x:.1f}% used")
    campaign_df["Total Impressions Text"] = campaign_df["TotalImpressions"].apply(format_int)
    campaign_df["Average CPM Text"] = campaign_df["AverageCPM"].apply(format_rm)
    campaign_df["Objective Text"] = campaign_df["Objective"].fillna("-").astype(str).str.strip().replace("", "-")
    campaign_df["Target Gender Text"] = campaign_df["TargetGender"].fillna("-").astype(str).str.strip().replace("", "-")
    campaign_df["Target Age Group Text"] = campaign_df["TargetAgeGroup"].fillna("-").astype(str).str.strip().replace("", "-")
    campaign_df["Target Interests Text"] = campaign_df["TargetInterests"].fillna("-").astype(str).str.strip().replace("", "-")

    campaign_df["Meta Line 1"] = campaign_df.apply(
        lambda row: f"{row['Company']} • {row['AdPlatform']}" if row["AdPlatform"] else row["Company"],
        axis=1
    )

    campaign_df["Meta Line 2"] = campaign_df.apply(
        lambda row: f"{row['Campaign ID']} • {row['Start Date']} - {row['End Date']} • {row['Campaign Duration (Days)']} days"
        if row["Campaign Duration (Days)"]
        else f"{row['Campaign ID']} • {row['Start Date']} - {row['End Date']}",
        axis=1
    )

    return campaign_df


def apply_campaign_search(df: pd.DataFrame, keyword: str) -> pd.DataFrame:
    if not keyword or not keyword.strip():
        return df

    keyword = keyword.strip().lower()

    mask = (
        df["Campaign ID"].astype(str).str.lower().str.contains(keyword, na=False)
        | df["CampaignName"].astype(str).str.lower().str.contains(keyword, na=False)
        | df["Company"].astype(str).str.lower().str.contains(keyword, na=False)
    )

    return df[mask].copy()


def apply_campaign_sort(df: pd.DataFrame, sort_option: str) -> pd.DataFrame:
    if df.empty:
        return df

    sortable_df = df.copy()

    if sort_option == "Latest Start Date":
        return sortable_df.sort_values(
            by=["StartDateParsed", "TotalSpend"],
            ascending=[False, False],
            kind="mergesort",
            na_position="last"
        ).reset_index(drop=True)

    if sort_option == "Earliest Start Date":
        return sortable_df.sort_values(
            by=["StartDateParsed", "TotalSpend"],
            ascending=[True, False],
            kind="mergesort",
            na_position="last"
        ).reset_index(drop=True)

    if sort_option == "Highest Spend":
        return sortable_df.sort_values(
            by=["TotalSpend", "StartDateParsed"],
            ascending=[False, False],
            kind="mergesort",
            na_position="last"
        ).reset_index(drop=True)

    if sort_option == "Highest Impressions":
        return sortable_df.sort_values(
            by=["TotalImpressions", "StartDateParsed"],
            ascending=[False, False],
            kind="mergesort",
            na_position="last"
        ).reset_index(drop=True)

    if sort_option == "Lowest CPM":
        return sortable_df.sort_values(
            by=["AverageCPM", "StartDateParsed"],
            ascending=[True, False],
            kind="mergesort",
            na_position="last"
        ).reset_index(drop=True)

    if sort_option == "Highest Ads Count":
        return sortable_df.sort_values(
            by=["TotalAds", "StartDateParsed"],
            ascending=[False, False],
            kind="mergesort",
            na_position="last"
        ).reset_index(drop=True)

    return sortable_df.reset_index(drop=True)


# =========================================================
# ADS CARD LIST PREPARATION
# =========================================================
def prepare_ads_card_list(selected_df: pd.DataFrame) -> pd.DataFrame:
    if selected_df.empty:
        return pd.DataFrame(columns=[
            "Ad ID", "Ad Format", "Platform Text", "Meta Line 2",
            "Target Gender Search", "Target Age Group Search", "Target Interests Search",
            "TotalSpend", "TotalImpressions", "TotalClicks",
            "WeightedCPM", "WeightedCTR", "WeightedCPC",
            "Spend Text", "Spend Usage Text", "Impressions Text", "Clicks Text",
            "CPM Text", "CTR Text", "CPC Text"
        ])

    work_df = selected_df.copy()

    optional_text_defaults = {
        "AdFormat": "Unknown",
        "TargetGender": "-",
        "TargetAgeGroup": "-",
        "TargetInterests": "-",
        "AdPlatform": "-",
    }

    for col, fallback in optional_text_defaults.items():
        if col not in work_df.columns:
            work_df[col] = fallback
        work_df[col] = (
            work_df[col]
            .fillna("")
            .astype(str)
            .str.strip()
            .replace("", fallback)
        )

    if "Clicks" not in work_df.columns:
        work_df["Clicks"] = 0

    ads_card_df = (
        work_df.groupby("PerformanceID", as_index=False)
        .agg(
            AdFormat=("AdFormat", first_non_empty_text),
            AdPlatform=("AdPlatform", first_non_empty_text),
            TargetGender=("TargetGender", first_non_empty_text),
            TargetAgeGroup=("TargetAgeGroup", first_non_empty_text),
            TargetInterests=("TargetInterests", first_non_empty_text),
            TotalSpend=("Cost", "sum"),
            TotalImpressions=("Impressions", "sum"),
            TotalClicks=("Clicks", "sum"),
        )
    )

    ads_card_df["WeightedCPM"] = 0.0
    valid_impr = ads_card_df["TotalImpressions"] > 0
    ads_card_df.loc[valid_impr, "WeightedCPM"] = (
        ads_card_df.loc[valid_impr, "TotalSpend"] / ads_card_df.loc[valid_impr, "TotalImpressions"] * 1000
    )

    ads_card_df["WeightedCTR"] = 0.0
    ads_card_df.loc[valid_impr, "WeightedCTR"] = (
        ads_card_df.loc[valid_impr, "TotalClicks"] / ads_card_df.loc[valid_impr, "TotalImpressions"] * 100
    )

    ads_card_df["WeightedCPC"] = 0.0
    valid_clicks = ads_card_df["TotalClicks"] > 0
    ads_card_df.loc[valid_clicks, "WeightedCPC"] = (
        ads_card_df.loc[valid_clicks, "TotalSpend"] / ads_card_df.loc[valid_clicks, "TotalClicks"]
    )

    campaign_total_spend = ads_card_df["TotalSpend"].sum()
    ads_card_df["SpendUsagePct"] = 0.0
    if campaign_total_spend > 0:
        ads_card_df["SpendUsagePct"] = ads_card_df["TotalSpend"] / campaign_total_spend * 100

    ads_card_df = ads_card_df.sort_values(
        by=["TotalSpend", "TotalImpressions"],
        ascending=[False, False]
    ).reset_index(drop=True)

    ads_card_df["Ad ID"] = ads_card_df["PerformanceID"].astype(str).str.strip()
    ads_card_df["Ad Format"] = ads_card_df["AdFormat"].apply(lambda x: normalize_display_text(x, "Unknown"))
    ads_card_df["Platform Text"] = ads_card_df["AdPlatform"].apply(lambda x: normalize_display_text(x, "-"))

    ads_card_df["Target Gender Search"] = ads_card_df["TargetGender"].apply(lambda x: normalize_display_text(x, "-"))
    ads_card_df["Target Age Group Search"] = ads_card_df["TargetAgeGroup"].apply(lambda x: normalize_display_text(x, "-"))
    ads_card_df["Target Interests Search"] = ads_card_df["TargetInterests"].apply(lambda x: normalize_display_text(x, "-"))

    ads_card_df["Meta Line 2"] = ads_card_df.apply(
        lambda row: join_meta_parts([
            row["Target Gender Search"],
            row["Target Age Group Search"],
            truncate_text(row["Target Interests Search"], max_len=28),
        ]),
        axis=1
    )

    ads_card_df["Spend Text"] = ads_card_df["TotalSpend"].apply(format_rm)
    ads_card_df["Spend Usage Text"] = ads_card_df["SpendUsagePct"].apply(lambda x: f"{x:.1f}% of campaign spend")
    ads_card_df["Impressions Text"] = ads_card_df["TotalImpressions"].apply(format_int)
    ads_card_df["Clicks Text"] = ads_card_df["TotalClicks"].apply(format_int)
    ads_card_df["CPM Text"] = ads_card_df["WeightedCPM"].apply(format_rm)
    ads_card_df["CTR Text"] = ads_card_df["WeightedCTR"].apply(format_pct)
    ads_card_df["CPC Text"] = ads_card_df["WeightedCPC"].apply(format_rm)

    return ads_card_df


def apply_ads_card_search(df: pd.DataFrame, keyword: str) -> pd.DataFrame:
    if df.empty:
        return df

    if not keyword or not keyword.strip():
        return df

    keyword = keyword.strip().lower()

    mask = (
        df["Ad ID"].astype(str).str.lower().str.contains(keyword, na=False)
        | df["Target Gender Search"].astype(str).str.lower().str.contains(keyword, na=False)
        | df["Target Age Group Search"].astype(str).str.lower().str.contains(keyword, na=False)
        | df["Target Interests Search"].astype(str).str.lower().str.contains(keyword, na=False)
    )

    return df[mask].copy()


def apply_ads_card_sort(df: pd.DataFrame, sort_option: str) -> pd.DataFrame:
    if df.empty:
        return df

    sortable_df = df.copy()

    if sort_option == "Highest Spend":
        return sortable_df.sort_values(
            by=["TotalSpend", "TotalImpressions"],
            ascending=[False, False],
            kind="mergesort"
        ).reset_index(drop=True)

    if sort_option == "Highest Impressions":
        return sortable_df.sort_values(
            by=["TotalImpressions", "TotalSpend"],
            ascending=[False, False],
            kind="mergesort"
        ).reset_index(drop=True)

    if sort_option == "Highest Clicks":
        return sortable_df.sort_values(
            by=["TotalClicks", "TotalImpressions"],
            ascending=[False, False],
            kind="mergesort"
        ).reset_index(drop=True)

    if sort_option == "Lowest CPM":
        return sortable_df.sort_values(
            by=["WeightedCPM", "TotalImpressions"],
            ascending=[True, False],
            kind="mergesort"
        ).reset_index(drop=True)

    if sort_option == "Highest CTR":
        return sortable_df.sort_values(
            by=["WeightedCTR", "TotalClicks"],
            ascending=[False, False],
            kind="mergesort"
        ).reset_index(drop=True)

    if sort_option == "Lowest CPC":
        return sortable_df.sort_values(
            by=["WeightedCPC", "TotalClicks"],
            ascending=[True, False],
            kind="mergesort"
        ).reset_index(drop=True)

    return sortable_df.reset_index(drop=True)



# =========================================================
# PPT EXPORT HELPERS
# =========================================================
def replace_text_in_text_frame(text_frame, replacements: dict):
    for paragraph in text_frame.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)

        new_text = full_text
        for placeholder, value in replacements.items():
            new_text = new_text.replace(placeholder, str(value))

        if new_text != full_text:
            if paragraph.runs:
                for run in paragraph.runs:
                    run.text = ""
                paragraph.runs[0].text = new_text
            else:
                paragraph.text = new_text


def replace_text_in_shape(shape, replacements: dict):
    if shape.has_text_frame:
        replace_text_in_text_frame(shape.text_frame, replacements)

    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                replace_text_in_text_frame(cell.text_frame, replacements)

    if hasattr(shape, "shapes"):
        for sub_shape in shape.shapes:
            replace_text_in_shape(sub_shape, replacements)


def generate_ad_pptx(row, campaign_summary) -> bytes:
    if not AD_TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"PPT template not found: {AD_TEMPLATE_PATH}. "
            f"Please place 'template_ads_deck_placeholders.pptx' inside the template folder."
        )

    prs = Presentation(str(AD_TEMPLATE_PATH))

    replacements = {
        "{CampaignName}": normalize_display_text(campaign_summary.get("CampaignName", "-")),
        "{CampaignID}": normalize_display_text(campaign_summary.get("Campaign ID", "-")),
        "{Company}": normalize_display_text(campaign_summary.get("Company", "-")),
        "{PerformanceID}": normalize_display_text(row.get("Ad ID", "-")),

        "{AdPlatform}": normalize_display_text(row.get("Platform Text", "-")),
        "{Objective}": normalize_display_text(campaign_summary.get("Objective Text", "-")),
        "{StartDate}": normalize_display_text(campaign_summary.get("Start Date", "-")),
        "{EndDate}": normalize_display_text(campaign_summary.get("End Date", "-")),
        "{CampaignDurationDays}": normalize_display_text(campaign_summary.get("Campaign Duration (Days)", "-")),

        "{Cost}": normalize_display_text(row.get("Spend Text", "-")),
        "{SpendUsage}": f"{float(row.get('SpendUsagePct', 0)):.1f}",
        "{AdFormat}": normalize_display_text(row.get("Ad Format", "-")),

        "{TargetGender}": normalize_display_text(row.get("Target Gender Search", "-")),
        "{TargetAgeGroup}": normalize_display_text(row.get("Target Age Group Search", "-")),
        "{TargetInterests}": normalize_display_text(row.get("Target Interests Search", "-")),

        "{Impressions}": normalize_display_text(row.get("Impressions Text", "-")),
        "{Clicks}": normalize_display_text(row.get("Clicks Text", "-")),
        "{CPM}": normalize_display_text(row.get("CPM Text", "-")).replace("RM ", ""),
        "{CTR}": normalize_display_text(row.get("CTR Text", "-")),
        "{CPC}": normalize_display_text(row.get("CPC Text", "-")).replace("RM ", ""),
    }

    for slide in prs.slides:
        for shape in slide.shapes:
            replace_text_in_shape(shape, replacements)

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()


# =========================================================
# CAMPAIGN PPT EXPORT HELPERS
# =========================================================
def get_shape_text(shape) -> str:
    text_parts = []

    if shape.has_text_frame:
        for paragraph in shape.text_frame.paragraphs:
            text_parts.append("".join(run.text for run in paragraph.runs))

    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for paragraph in cell.text_frame.paragraphs:
                    text_parts.append("".join(run.text for run in paragraph.runs))

    if hasattr(shape, "shapes"):
        for sub_shape in shape.shapes:
            text_parts.append(get_shape_text(sub_shape))

    return "\n".join(text_parts)


def remove_shape(shape):
    element = shape._element
    element.getparent().remove(element)


def chart_to_png_bytes(fig, width: int = 900, height: int = 520) -> BytesIO:
    try:
        png_bytes = fig.to_image(format="png", width=width, height=height, scale=2)
    except Exception as e:
        raise RuntimeError(
            "Chart export failed. Please install Kaleido with: pip install kaleido"
        ) from e

    return BytesIO(png_bytes)


def insert_image_at_placeholder(prs: Presentation, placeholder: str, image_stream: BytesIO) -> bool:
    for slide in prs.slides:
        for shape in list(slide.shapes):
            if placeholder in get_shape_text(shape):
                left, top, width, height = shape.left, shape.top, shape.width, shape.height
                remove_shape(shape)
                image_stream.seek(0)
                slide.shapes.add_picture(image_stream, left, top, width=width, height=height)
                return True

    return False


def prepare_ppt_chart(fig, show_title: bool = False, chart_kind: str = "vertical"):
    fig = go.Figure(fig)

    if not show_title:
        fig.update_layout(title="")

    # More margin is reserved for exported charts so data labels are not clipped.
    right_margin = 95 if chart_kind == "horizontal" else 30
    top_margin = 34 if not show_title else 48

    fig.update_layout(
        template="plotly_white",
        height=420,
        margin=dict(l=60, r=right_margin, t=top_margin, b=55),
        font=dict(family="Arial", size=12, color="#334155"),
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        uniformtext_minsize=10,
        uniformtext_mode="show",
    )

    return fig

def build_campaign_ppt_charts(selected_df: pd.DataFrame) -> dict:
    ad_format_df = prepare_ad_format_chart_data(selected_df)

    spend_by_format_fig = build_ad_format_bar_chart(
        ad_format_df,
        value_col="TotalSpend",
        title="",
        value_type="spend",
        show_values=True,
    )

    impressions_by_format_fig = build_ad_format_bar_chart(
        ad_format_df,
        value_col="TotalImpressions",
        title="",
        value_type="impressions",
        show_values=True,
    )

    top_ads_by_spend_df = prepare_top_ads_chart_data(selected_df, top_n=10, sort_by="Cost")
    top_ads_by_impressions_df = prepare_top_ads_chart_data(selected_df, top_n=10, sort_by="Impressions")

    top_ads_by_spend_fig = build_top_ads_chart(
        top_ads_by_spend_df,
        title="",
        metric_col="Cost",
        value_title="Spend (RM)",
        show_values=True,
    )

    top_ads_by_impressions_fig = build_top_ads_chart(
        top_ads_by_impressions_df,
        title="",
        metric_col="Impressions",
        value_title="Impressions",
        show_values=True,
    )

    return {
        "{SpendByAdFormatChart}": chart_to_png_bytes(
            prepare_ppt_chart(spend_by_format_fig, chart_kind="vertical")
        ),
        "{ImpressionsByAdFormatChart}": chart_to_png_bytes(
            prepare_ppt_chart(impressions_by_format_fig, chart_kind="vertical")
        ),
        "{TopAdsBySpendChart}": chart_to_png_bytes(
            prepare_ppt_chart(top_ads_by_spend_fig, chart_kind="horizontal"),
            width=920,
            height=520,
        ),
        "{TopAdsByImpressionsChart}": chart_to_png_bytes(
            prepare_ppt_chart(top_ads_by_impressions_fig, chart_kind="horizontal"),
            width=920,
            height=520,
        ),
    }

def generate_campaign_pptx(campaign_summary, selected_df: pd.DataFrame) -> bytes:
    if not CAMPAIGN_TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"Campaign PPT template not found: {CAMPAIGN_TEMPLATE_PATH}. "
            f"Please place 'template_campaign_deck_placeholders.pptx' inside the template folder."
        )

    prs = Presentation(str(CAMPAIGN_TEMPLATE_PATH))

    replacements = {
        "{CampaignName}": normalize_display_text(campaign_summary.get("CampaignName", "-")),
        "{CampaignID}": normalize_display_text(campaign_summary.get("Campaign ID", "-")),
        "{Company}": normalize_display_text(campaign_summary.get("Company", "-")),
        "{AdPlatform}": normalize_display_text(campaign_summary.get("AdPlatform", "-")),
        "{Objective}": normalize_display_text(campaign_summary.get("Objective Text", "-")),

        "{StartDate}": normalize_display_text(campaign_summary.get("Start Date", "-")),
        "{EndDate}": normalize_display_text(campaign_summary.get("End Date", "-")),
        "{CampaignDurationDays}": normalize_display_text(campaign_summary.get("Campaign Duration (Days)", "-")),

        "{TotalAds}": normalize_display_text(campaign_summary.get("Total Ads Text", "-")),
        "{TotalSpend}": normalize_display_text(campaign_summary.get("Total Spend Text", "-")),
        "{CampaignBudget}": normalize_display_text(campaign_summary.get("Campaign Budget Text", "-")),
        "{BudgetUsedPct}": f"{float(campaign_summary.get('BudgetUsedPct', 0)):.1f}",

        "{TargetGender}": normalize_display_text(campaign_summary.get("Target Gender Text", "-")),
        "{TargetAgeGroup}": normalize_display_text(campaign_summary.get("Target Age Group Text", "-")),
        "{TargetInterests}": normalize_display_text(campaign_summary.get("Target Interests Text", "-")),

        "{TotalImpressions}": normalize_display_text(campaign_summary.get("Total Impressions Text", "-")),
        "{AverageCPM}": normalize_display_text(campaign_summary.get("Detail Avg CPM Text", campaign_summary.get("Average CPM Text", "-"))),
        "{TotalClicks}": normalize_display_text(campaign_summary.get("Total Clicks Text", "-")),
        "{CTR}": normalize_display_text(campaign_summary.get("Detail CTR Text", "-")),
        "{AverageCPC}": normalize_display_text(campaign_summary.get("Detail Avg CPC Text", "-")),
    }

    for slide in prs.slides:
        for shape in list(slide.shapes):
            replace_text_in_shape(shape, replacements)

    chart_streams = build_campaign_ppt_charts(selected_df)
    for placeholder, image_stream in chart_streams.items():
        inserted = insert_image_at_placeholder(prs, placeholder, image_stream)
        if not inserted:
            raise ValueError(f"Chart placeholder not found in campaign template: {placeholder}")

    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()


def create_campaign_ppt_export_request_link(campaign_id: str) -> str:
    campaign_id_url = quote_plus(str(campaign_id))
    export_token = str(time.time_ns())

    return (
        f'<a class="export-btn-link" '
        f'href="?export_campaign={campaign_id_url}&campaign_export_token={export_token}" '
        f'target="_self" '
        f'onclick="this.classList.add(\'export-btn-loading\'); this.innerHTML=\'⏳ <span>Exporting...</span>\';" '
        f'title="Export Campaign PPTX">📄 <span>Export</span></a>'
    )

def create_ppt_export_request_link(campaign_id: str, ad_id: str) -> str:
    campaign_id_url = quote_plus(str(campaign_id))
    ad_id_url = quote_plus(str(ad_id))
    export_token = str(time.time_ns())

    return (
        f'<a class="export-btn-link" '
        f'href="?campaign_id={campaign_id_url}&export_ad={ad_id_url}&export_token={export_token}" '
        f'target="_self" '
        f'onclick="this.classList.add(\'export-btn-loading\'); this.innerHTML=\'⏳ <span>Exporting...</span>\';" '
        f'title="Export PPTX">📄 <span>Export</span></a>'
    )


def trigger_pptx_download(
    ppt_bytes: bytes,
    filename: str,
    clean_url: bool = False,
    campaign_id: str = "",
    clean_to_list: bool = False,
):
    safe_filename = filename.replace("/", "-").replace("\\", "-").replace(":", "-")
    b64 = base64.b64encode(ppt_bytes).decode("utf-8")

    clean_script = ""
    if clean_url and clean_to_list:
        clean_script = "window.parent.history.replaceState(null, '', window.parent.location.pathname);"
    elif clean_url and campaign_id:
        campaign_id_url = quote_plus(str(campaign_id))
        clean_script = f"window.parent.history.replaceState(null, '', '?campaign_id={campaign_id_url}');"

    components.html(
        f"""
        <html>
        <body>
            <a id="download-link"
               href="data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,{b64}"
               download="{safe_filename}">
            </a>
            <script>
                document.getElementById("download-link").click();
                try {{
                    window.parent.document.querySelectorAll(".export-loading-card").forEach(function(el) {{ el.remove(); }});
                    window.parent.document.querySelectorAll(".export-btn-loading").forEach(function(el) {{
                        el.classList.remove("export-btn-loading");
                        el.innerHTML = "📄 <span>Export</span>";
                    }});
                }} catch (e) {{}}
                {clean_script}
            </script>
        </body>
        </html>
        """,
        height=0,
    )


def render_export_loading_box(title: str, message: str):
    placeholder = st.empty()
    placeholder.markdown(
        f"""
        <div class="export-loading-card">
            <div class="export-loading-spinner"></div>
            <div>
                <div class="export-loading-title">{escape(str(title))}</div>
                <div class="export-loading-message">{escape(str(message))}</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )
    return placeholder


def clear_export_loading(loading_box=None, progress_box=None):
    if progress_box is not None:
        progress_box.empty()
    if loading_box is not None:
        loading_box.empty()

# =========================================================
# DETAIL VIEW HELPERS
# =========================================================
def prepare_campaign_detail(df: pd.DataFrame, campaign_id: str):
    selected_df = df[df["CampaignID"].astype(str) == str(campaign_id)].copy()

    if selected_df.empty:
        return None, pd.DataFrame(), pd.DataFrame()

    campaign_summary = prepare_campaign_list(selected_df).iloc[0].copy()

    total_spend = pd.to_numeric(selected_df["Cost"], errors="coerce").fillna(0).sum()
    total_impressions = pd.to_numeric(selected_df["Impressions"], errors="coerce").fillna(0).sum()

    if "Clicks" in selected_df.columns:
        total_clicks = pd.to_numeric(selected_df["Clicks"], errors="coerce").fillna(0).sum()
    else:
        total_clicks = 0.0

    avg_cpm = (total_spend / total_impressions * 1000) if total_impressions > 0 else 0.0
    ctr = (total_clicks / total_impressions * 100) if total_impressions > 0 else 0.0
    avg_cpc = (total_spend / total_clicks) if total_clicks > 0 else 0.0

    campaign_summary["Total Clicks"] = total_clicks
    campaign_summary["Total Clicks Text"] = format_int(total_clicks)
    campaign_summary["Detail CTR"] = ctr
    campaign_summary["Detail CTR Text"] = format_pct(ctr)
    campaign_summary["Detail Avg CPM"] = avg_cpm
    campaign_summary["Detail Avg CPM Text"] = format_rm(avg_cpm)
    campaign_summary["Detail Avg CPC"] = avg_cpc
    campaign_summary["Detail Avg CPC Text"] = format_rm(avg_cpc)

    detail_columns = [
        "PerformanceID",
        "AdPlatform",
        "Cost",
        "Impressions",
    ]

    optional_detail_columns = [
        "Clicks",
        "CTR",
        "CPC",
        "CPM",
        "AdFormat",
        "TargetGender",
        "TargetAgeGroup",
        "TargetInterests",
    ]

    for col in optional_detail_columns:
        if col in selected_df.columns:
            detail_columns.append(col)

    ads_df = selected_df[detail_columns].copy()

    return campaign_summary, ads_df, selected_df


def render_detail_view(df: pd.DataFrame, campaign_id: str):
    campaign_summary, ads_df, selected_df = prepare_campaign_detail(df, campaign_id)

    if campaign_summary is None:
        st.warning("Selected campaign not found.")
        if st.button("← Back to Campaign List"):
            st.query_params.clear()
            st.rerun()
        return

    top_left, top_right = st.columns([6, 1])

    with top_left:
        st.markdown(
            f'<div class="section-title">{escape(str(campaign_summary["CampaignName"]))}</div>',
            unsafe_allow_html=True
        )
        st.markdown(
            f'<div class="section-note">{escape(str(campaign_summary["Meta Line 1"]))}<br>{escape(str(campaign_summary["Meta Line 2"]))}</div>',
            unsafe_allow_html=True
        )

    with top_right:
        st.write("")
        if st.button("← Back", use_container_width=True):
            st.query_params.clear()
            st.rerun()

    d1, d2, d3, d4, d5 = st.columns(5)

    with d1:
        render_kpi_card("Total Ads", str(campaign_summary["Total Ads Text"]), "")

    with d2:
        render_kpi_card(
            "Total Spend (RM)",
            str(campaign_summary["Total Spend Text"]),
            format_budget_usage_subtitle(
                campaign_summary["CampaignBudget"],
                campaign_summary["BudgetUsedPct"],
                fallback=""
            )
        )

    with d3:
        render_kpi_card(
            "Total Impressions",
            str(campaign_summary["Total Impressions Text"]),
            f"Avg. CPM: {campaign_summary['Detail Avg CPM Text']}"
        )

    with d4:
        render_kpi_card(
            "Total Clicks",
            str(campaign_summary["Total Clicks Text"]),
            f"CTR: {campaign_summary['Detail CTR Text']}"
        )

    with d5:
        render_kpi_card("Avg. CPC (RM)", str(campaign_summary["Detail Avg CPC Text"]), "")

    ad_format_df = prepare_ad_format_chart_data(selected_df)
    top_ads_df = prepare_top_ads_chart_data(selected_df, top_n=10, sort_by="Impressions")

    vc1, vc2, vc3 = st.columns(3)

    with vc1:
        spend_by_format_fig = build_ad_format_bar_chart(
            ad_format_df,
            value_col="TotalSpend",
            title="Total Spend by Ad Format",
            value_type="spend"
        )
        st.plotly_chart(spend_by_format_fig, use_container_width=True, config={"displayModeBar": False})

    with vc2:
        impressions_by_format_fig = build_ad_format_bar_chart(
            ad_format_df,
            value_col="TotalImpressions",
            title="Total Impressions by Ad Format",
            value_type="impressions"
        )
        st.plotly_chart(impressions_by_format_fig, use_container_width=True, config={"displayModeBar": False})

    with vc3:
        top_ads_fig = build_top_ads_chart(
            top_ads_df,
            title="Top 10 Ads by Impressions",
            metric_col="Impressions",
            value_title="Impressions"
        )
        st.plotly_chart(top_ads_fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown('<div class="section-title">Ads List</div>', unsafe_allow_html=True)

    ads_search_col, ads_sort_col = st.columns([7.8, 1.4], gap="small")

    with ads_search_col:
        st.markdown(
            '<div class="search-box-label">Search by ad ID, target gender, target age group, or target interests.</div>',
            unsafe_allow_html=True
        )
        ads_search_keyword = st.text_input(
            "Search ads list",
            placeholder="Search ad ID, target gender, target age group, or target interests",
            key=f"ads_detail_search_{campaign_id}",
            label_visibility="collapsed"
        )

    with ads_sort_col:
        st.markdown('<div class="sort-box-label">Sort by</div>', unsafe_allow_html=True)
        ads_sort_option = st.selectbox(
            "Sort ads list",
            options=[
                "Highest Spend",
                "Highest Impressions",
                "Highest Clicks",
                "Lowest CPM",
                "Highest CTR",
                "Lowest CPC",
            ],
            key=f"ads_detail_sort_{campaign_id}",
            label_visibility="collapsed"
        )

    ads_card_df = prepare_ads_card_list(selected_df)
    ads_card_df = apply_ads_card_search(ads_card_df, ads_search_keyword)
    ads_card_df = apply_ads_card_sort(ads_card_df, ads_sort_option)

    export_ad_id = st.query_params.get("export_ad")
    export_token = st.query_params.get("export_token")

    if export_ad_id and export_token:
        processed_token_key = f"processed_export_token_{campaign_id}"

        if st.session_state.get(processed_token_key) != str(export_token):
            st.session_state[processed_token_key] = str(export_token)

            export_source_df = prepare_ads_card_list(selected_df)
            selected_export_df = export_source_df[
                export_source_df["Ad ID"].astype(str) == str(export_ad_id)
            ]

            if not selected_export_df.empty:
                loading_box = render_export_loading_box(
                    "Generating ad report",
                    "Preparing the PowerPoint file. Please wait a moment..."
                )
                progress_box = st.empty()
                try:
                    export_row = selected_export_df.iloc[0]
                    progress_box.progress(20, text="Preparing ad data...")
                    progress_box.progress(55, text="Replacing PowerPoint placeholders...")
                    ppt_bytes = generate_ad_pptx(export_row, campaign_summary)
                    progress_box.progress(90, text="Starting download...")
                    ppt_filename = f"{export_row['Ad ID']}_Ad_Performance_Report.pptx"
                    trigger_pptx_download(
                        ppt_bytes,
                        ppt_filename,
                        clean_url=True,
                        campaign_id=campaign_id
                    )
                    progress_box.progress(100, text="Download is starting...")
                    time.sleep(0.3)
                    clear_export_loading(loading_box, progress_box)
                except Exception as e:
                    clear_export_loading(loading_box, progress_box)
                    st.error(f"Export failed: {e}")
            else:
                st.error("Selected ad was not found for export.")

    if export_ad_id and not export_token:
        components.html(
            f"""
            <script>
                window.parent.history.replaceState(null, '', '?campaign_id={quote_plus(str(campaign_id))}');
            </script>
            """,
            height=0,
        )

    st.caption(f"Showing {len(ads_card_df):,} ad(s)")

    if ads_card_df.empty:
        st.markdown(
            '<div class="campaign-empty">No ads found for the current search and sort.</div>',
            unsafe_allow_html=True
        )
    else:
        render_ads_list_header()
        for _, row in ads_card_df.iterrows():
            render_ads_card(row, campaign_summary, campaign_id)


# =========================================================
# CAMPAIGN CARD RENDERING
# =========================================================
def render_campaign_list_header():
    st.markdown(
        """
        <div class="campaign-list-header">
            <div class="campaign-grid campaign-grid-header">
                <div>Campaign</div>
                <div>Objectives</div>
                <div>Budget usage</div>
                <div>Impressions</div>
                <div>CPM</div>
                <div>Ads</div>
                <div>Export</div>
                <div></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


def render_campaign_card(row):
    campaign_name = escape(str(row["CampaignName"]))
    meta_line_1 = escape(str(row["Meta Line 1"]))
    meta_line_2 = escape(str(row["Meta Line 2"]))
    objective_text = escape(str(row["Objective Text"]))
    budget_usage = escape(str(row["Budget Usage Text"]))
    budget_used_text = escape(str(row["Budget Used Text"]))
    impressions_text = escape(str(row["Total Impressions Text"]))
    cpm_text = escape(str(row["Average CPM Text"]))
    ads_text = escape(str(row["Total Ads Text"]))
    campaign_id_url = quote_plus(str(row["Campaign ID"]))
    export_link = create_campaign_ppt_export_request_link(row["Campaign ID"])

    html = dedent(f"""
    <div class="campaign-card-row">
        <div class="campaign-grid campaign-grid-row">
            <div class="campaign-main">
                <div class="campaign-name">{campaign_name}</div>
                <div class="campaign-meta">{meta_line_1}</div>
                <div class="campaign-meta">{meta_line_2}</div>
            </div>
            <div class="campaign-metric">
                <div class="metric-main">{objective_text}</div>
            </div>
            <div class="campaign-metric">
                <div class="metric-main">{budget_usage}</div>
                <div class="metric-sub">{budget_used_text}</div>
            </div>
            <div class="campaign-metric">
                <div class="metric-main">{impressions_text}</div>
            </div>
            <div class="campaign-metric">
                <div class="metric-main">{cpm_text}</div>
            </div>
            <div class="campaign-metric">
                <div class="metric-main">{ads_text}</div>
            </div>
            <div class="campaign-export-wrap">
                {export_link}
            </div>
            <div class="campaign-arrow-wrap">
                <a class="campaign-arrow-btn-link" href="?campaign_id={campaign_id_url}" target="_self">›</a>
            </div>
        </div>
    </div>
    """).strip()

    st.markdown(html, unsafe_allow_html=True)


# =========================================================
# ADS CARD RENDERING
# =========================================================
def render_ads_list_header():
    st.markdown(
        """
        <div class="ads-list-header">
            <div class="ads-grid ads-grid-header">
                <div>Ad</div>
                <div>Ad Format</div>
                <div>Spend</div>
                <div>Impressions</div>
                <div>Clicks</div>
                <div>CPM</div>
                <div>CTR</div>
                <div>CPC</div>
                <div></div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


def render_ads_card(row, campaign_summary, campaign_id):
    ad_id = escape(str(row["Ad ID"]))
    platform_text = escape(str(row["Platform Text"]))
    meta_line_2 = escape(str(row["Meta Line 2"]))
    ad_format_text = escape(str(row["Ad Format"]))
    spend_text = escape(str(row["Spend Text"]))
    spend_usage_text = escape(str(row["Spend Usage Text"]))
    impressions_text = escape(str(row["Impressions Text"]))
    clicks_text = escape(str(row["Clicks Text"]))
    cpm_text = escape(str(row["CPM Text"]))
    ctr_text = escape(str(row["CTR Text"]))
    cpc_text = escape(str(row["CPC Text"]))

    export_link = create_ppt_export_request_link(
        campaign_id=campaign_id,
        ad_id=row["Ad ID"]
    )

    html = dedent(f"""
    <div class="ads-card-row">
        <div class="ads-grid ads-grid-row">
            <div class="campaign-main">
                <div class="campaign-name">{ad_id}</div>
                <div class="campaign-meta">{platform_text}</div>
                <div class="campaign-meta">{meta_line_2}</div>
            </div>
            <div class="campaign-metric">
                <div class="metric-main">{ad_format_text}</div>
            </div>
            <div class="campaign-metric">
                <div class="metric-main">{spend_text}</div>
                <div class="metric-sub">{spend_usage_text}</div>
            </div>
            <div class="campaign-metric">
                <div class="metric-main">{impressions_text}</div>
            </div>
            <div class="campaign-metric">
                <div class="metric-main">{clicks_text}</div>
            </div>
            <div class="campaign-metric">
                <div class="metric-main">{cpm_text}</div>
            </div>
            <div class="campaign-metric">
                <div class="metric-main">{ctr_text}</div>
            </div>
            <div class="campaign-metric">
                <div class="metric-main">{cpc_text}</div>
            </div>
            <div class="campaign-arrow-wrap">
                {export_link}
            </div>
        </div>
    </div>
    """).strip()

    st.markdown(html, unsafe_allow_html=True)


# =========================================================
# PAGE / CARD STYLES
# =========================================================
def inject_page_css():
    st.markdown(
        """
        <style>
        .page-title {
            font-size: 28px;
            font-weight: 700;
            color: #1e293b;
            margin-top: 0px;
            margin-bottom: 0px;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        .overview-text {
            font-size: 14px;
            color: #8a94a6;
            margin-top: -2px;
            margin-bottom: 2px;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        .filter-label-inline {
            font-size: 12px;
            color: #7c8799;
            font-weight: 600;
            text-align: right;
            margin-top: 6px;
            white-space: nowrap;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        .sort-box-label {
            font-size: 12px;
            color: #7c8799;
            font-weight: 600;
            margin-bottom: 6px;
            text-align: left;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        .search-box-label {
            font-size: 13px;
            color: #8a94a6;
            margin-bottom: 6px;
            text-align: left;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        .section-title {
            font-size: 22px;
            font-weight: 700;
            color: #1e293b;
            margin-top: 10px;
            margin-bottom: 6px;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        .section-note {
            font-size: 13px;
            color: #8a94a6;
            margin-top: 0px;
            margin-bottom: 10px;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        .kpi-card {
            background: #ffffff;
            border: 1px solid #d9e2ec;
            border-radius: 18px;
            padding: 18px 18px 16px 18px;
            box-shadow: 0 4px 14px rgba(15, 23, 42, 0.05);
            min-height: 150px;
            display: flex;
            flex-direction: column;
            justify-content: flex-start;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        .kpi-accent {
            width: 36px;
            height: 3px;
            border-radius: 999px;
            background: #2563eb;
            margin-bottom: 14px;
        }

        .kpi-label {
            font-size: 13px;
            color: #475569;
            font-weight: 600;
            margin-bottom: 12px;
        }

        .kpi-value {
            font-size: 27px;
            color: #0f172a;
            font-weight: 700;
            line-height: 1.2;
            margin-bottom: 14px;
            word-break: break-word;
        }

        .kpi-subtitle {
            font-size: 14px;
            color: #64748b;
            font-weight: 500;
            line-height: 1.4;
            margin-top: auto;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }

        .campaign-list-header,
        .ads-list-header {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-bottom: none;
            border-radius: 16px 16px 0 0;
            padding: 10px 16px;
            margin-top: 8px;
        }

        .campaign-card-row,
        .ads-card-row {
            background: #ffffff;
            border-left: 1px solid #e2e8f0;
            border-right: 1px solid #e2e8f0;
            border-bottom: 1px solid #e2e8f0;
            padding: 16px;
        }

        .campaign-card-row:last-child,
        .ads-card-row:last-child {
            border-radius: 0 0 16px 16px;
        }

        .campaign-grid {
            display: grid;
            grid-template-columns: 3fr 1.55fr 2fr 1.35fr 1fr 0.9fr 0.85fr 0.45fr;
            gap: 14px;
            align-items: center;
        }

        .ads-grid {
            display: grid;
            grid-template-columns: 2.3fr 1.15fr 1.15fr 1.2fr 0.95fr 0.85fr 0.85fr 0.85fr 0.9fr;
            gap: 16px;
            align-items: center;
        }

        .campaign-grid-header,
        .ads-grid-header {
            font-size: 13px;
            color: #64748b;
            font-weight: 700;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        .campaign-main {
            min-width: 0;
        }

        .campaign-name {
            font-size: 16px;
            font-weight: 700;
            color: #0f172a;
            margin-bottom: 4px;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        .campaign-meta {
            font-size: 13px;
            color: #64748b;
            line-height: 1.45;
            font-family: "Segoe UI", Arial, sans-serif;
            word-break: break-word;
        }

        .campaign-metric {
            min-width: 0;
        }

        .metric-main {
            font-size: 15px;
            font-weight: 700;
            color: #0f172a;
            line-height: 1.35;
            font-family: "Segoe UI", Arial, sans-serif;
            word-break: break-word;
        }

        .metric-sub {
            font-size: 13px;
            color: #64748b;
            margin-top: 3px;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        .campaign-export-wrap,
        .campaign-arrow-wrap {
            display: flex;
            justify-content: flex-end;
        }

        .campaign-arrow-btn-link {
            width: 34px;
            height: 34px;
            border: 1px solid #d1d9e6;
            border-radius: 10px;
            display: flex;
            align-items: center;
            justify-content: center;
            color: #475569;
            font-size: 18px;
            font-weight: 700;
            background: #ffffff;
            font-family: "Segoe UI", Arial, sans-serif;
            text-decoration: none;
        }

        .export-btn-link {
            min-width: 88px;
            height: 36px;
            padding: 0 12px;
            border: 1px solid #d1d9e6;
            border-radius: 10px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            color: #475569;
            font-size: 13px;
            font-weight: 700;
            background: #ffffff;
            font-family: "Segoe UI", Arial, sans-serif;
            text-decoration: none;
            white-space: nowrap;
            transition: all 0.2s ease;
        }

        .export-btn-link:hover {
            background: #0f766e;
            color: #ffffff;
            border-color: #0f766e;
            transform: translateY(-1px);
        }

        .export-btn-link:active {
            transform: translateY(0px);
        }

        .export-btn-disabled,
        .export-btn-loading {
            opacity: 0.70;
            pointer-events: none;
            cursor: wait;
        }

        .export-loading-card {
            display: flex;
            align-items: center;
            gap: 14px;
            background: #ffffff;
            border: 1px solid #d9e2ec;
            border-radius: 16px;
            padding: 14px 16px;
            margin: 14px 0 10px 0;
            box-shadow: 0 6px 18px rgba(15, 23, 42, 0.08);
            font-family: "Segoe UI", Arial, sans-serif;
        }

        .export-loading-spinner {
            width: 28px;
            height: 28px;
            border: 3px solid #e2e8f0;
            border-top-color: #0f766e;
            border-radius: 50%;
            animation: export-spin 0.85s linear infinite;
            flex-shrink: 0;
        }

        .export-loading-title {
            color: #0f172a;
            font-size: 15px;
            font-weight: 700;
            margin-bottom: 2px;
        }

        .export-loading-message {
            color: #64748b;
            font-size: 13px;
            font-weight: 500;
        }

        @keyframes export-spin {
            from { transform: rotate(0deg); }
            to { transform: rotate(360deg); }
        }

        .campaign-empty {
            background: #ffffff;
            border: 1px dashed #d9e2ec;
            border-radius: 16px;
            padding: 22px;
            color: #64748b;
            font-size: 14px;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        @media (max-width: 1500px) {
            .ads-grid {
                grid-template-columns: 2.1fr 1.05fr 1.05fr 1.1fr 0.9fr 0.8fr 0.8fr 0.8fr 0.85fr;
                gap: 12px;
            }

            .metric-main {
                font-size: 14px;
            }
        }

        @media (max-width: 1300px) {
            .kpi-card {
                min-height: 145px;
            }

            .kpi-value {
                font-size: 24px;
            }

            .kpi-subtitle {
                font-size: 13px;
            }

            .campaign-grid {
                grid-template-columns: 2.5fr 1.25fr 1.8fr 1.2fr 0.85fr 0.75fr 0.8fr 0.45fr;
                gap: 10px;
            }

            .ads-grid {
                grid-template-columns: 1.9fr 0.95fr 0.95fr 1fr 0.8fr 0.75fr 0.75fr 0.75fr 0.8fr;
                gap: 10px;
            }

            .campaign-name {
                font-size: 15px;
            }

            .metric-main {
                font-size: 13px;
            }

            .campaign-meta {
                font-size: 12px;
            }
        }
        </style>
        """,
        unsafe_allow_html=True
    )


def render_kpi_card(title: str, value: str, subtitle: str = ""):
    safe_title = escape(str(title))
    safe_value = escape(str(value))
    safe_subtitle = escape(str(subtitle)) if subtitle else "&nbsp;"
    safe_title_attr = escape(str(subtitle)) if subtitle else ""

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-accent"></div>
            <div class="kpi-label">{safe_title}</div>
            <div class="kpi-value">{safe_value}</div>
            <div class="kpi-subtitle" title="{safe_title_attr}">{safe_subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


# =========================================================
# MAIN
# =========================================================
inject_page_css()

if "ads_year_filter" not in st.session_state:
    st.session_state["ads_year_filter"] = "All"

if "campaign_sort_option" not in st.session_state:
    st.session_state["campaign_sort_option"] = "Latest Start Date"

st.markdown('<div class="page-title">Campaign Ads</div>', unsafe_allow_html=True)

try:
    raw_df = load_ads_data()
    df = prepare_ads_data(raw_df)
except Exception as e:
    st.error(str(e))
    st.stop()

selected_year_for_view = st.session_state.get("ads_year_filter", "All")
if selected_year_for_view not in get_year_options(df):
    st.session_state["ads_year_filter"] = "All"
    selected_year_for_view = "All"

filtered_df_for_view = apply_year_filter(df, selected_year_for_view)

selected_campaign_id = st.query_params.get("campaign_id")
if selected_campaign_id:
    render_detail_view(filtered_df_for_view, selected_campaign_id)
    st.stop()

# ---------------------------------------------------------
# TOP ROW: OVERVIEW + FAR-RIGHT FILTER
# ---------------------------------------------------------
top_left, top_spacer, top_right = st.columns([7.6, 1.2, 2.2], gap="small")

with top_left:
    st.markdown(
        '<div class="overview-text">Overview of advertising campaigns, spend, impressions, and cost efficiency.</div>',
        unsafe_allow_html=True
    )

with top_right:
    label_col, control_col = st.columns([1.7, 4.3], gap="small")

    with label_col:
        st.markdown('<div class="filter-label-inline">Filter by</div>', unsafe_allow_html=True)

    with control_col:
        year_options = get_year_options(df)
        selected_year = render_year_filter(year_options)

filtered_df = apply_year_filter(df, selected_year)

# ---------------------------------------------------------
# CAMPAIGN PPT EXPORT REQUEST
# ---------------------------------------------------------
export_campaign_id = st.query_params.get("export_campaign")
campaign_export_token = st.query_params.get("campaign_export_token")

if export_campaign_id and campaign_export_token:
    processed_token_key = f"processed_campaign_export_token_{export_campaign_id}"

    if st.session_state.get(processed_token_key) != str(campaign_export_token):
        st.session_state[processed_token_key] = str(campaign_export_token)

        loading_box = render_export_loading_box(
            "Generating campaign report",
            "Building charts and preparing the PowerPoint file. Please wait a moment..."
        )
        progress_box = st.empty()

        selected_campaign_df = filtered_df[
            filtered_df["CampaignID"].astype(str) == str(export_campaign_id)
        ].copy()

        if selected_campaign_df.empty:
            clear_export_loading(loading_box, progress_box)
            st.error("Selected campaign was not found for export.")
        else:
            try:
                progress_box.progress(15, text="Preparing campaign data...")
                campaign_summary, _, selected_export_df = prepare_campaign_detail(
                    filtered_df,
                    str(export_campaign_id)
                )

                progress_box.progress(35, text="Creating charts for the report...")
                ppt_bytes = generate_campaign_pptx(campaign_summary, selected_export_df)
                progress_box.progress(85, text="Finalising PowerPoint slides...")

                safe_campaign_name = normalize_display_text(
                    campaign_summary.get("CampaignName", "Campaign"),
                    fallback="Campaign"
                )
                safe_campaign_id = normalize_display_text(
                    campaign_summary.get("Campaign ID", export_campaign_id),
                    fallback=str(export_campaign_id)
                )
                ppt_filename = f"{safe_campaign_id}_{safe_campaign_name}_Campaign_Report.pptx"

                trigger_pptx_download(
                    ppt_bytes,
                    ppt_filename,
                    clean_url=True,
                    clean_to_list=True,
                )
                progress_box.progress(100, text="Download is starting...")
                time.sleep(0.3)
                clear_export_loading(loading_box, progress_box)
            except Exception as e:
                clear_export_loading(loading_box, progress_box)
                st.error(f"Campaign export failed: {e}")

if export_campaign_id and not campaign_export_token:
    components.html(
        """
        <script>
            window.parent.history.replaceState(null, '', window.parent.location.pathname);
        </script>
        """,
        height=0,
    )

# ---------------------------------------------------------
# KPI CALCULATION ON FILTERED DATA
# ---------------------------------------------------------
kpis = calculate_kpis(filtered_df)

company_word = "company" if kpis["total_companies"] == 1 else "companies"

total_campaigns_text = format_int(kpis["total_campaigns"])
total_ads_text = format_int(kpis["total_ads"])
total_spend_text = format_compact_rm(kpis["total_spend"])
total_impressions_text = format_compact(kpis["total_impressions"])
overall_cpm_text = format_rm(kpis["overall_cpm"])

total_campaigns_subtitle = f"Across {format_int(kpis['total_companies'])} {company_word}"
total_spend_subtitle = format_budget_usage_subtitle(
    kpis["total_budget"],
    kpis["spend_pct"],
    fallback="Across all campaigns"
)

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    render_kpi_card(
        title="Total Campaigns",
        value=total_campaigns_text,
        subtitle=total_campaigns_subtitle
    )

with col2:
    render_kpi_card(
        title="Total Ads",
        value=total_ads_text,
        subtitle=""
    )

with col3:
    render_kpi_card(
        title="Total Spend (RM)",
        value=total_spend_text,
        subtitle=total_spend_subtitle
    )

with col4:
    render_kpi_card(
        title="Total Impressions",
        value=total_impressions_text,
        subtitle="Across all campaigns"
    )

with col5:
    render_kpi_card(
        title="Overall CPM",
        value=overall_cpm_text,
        subtitle="Weighted across all campaigns"
    )

# ---------------------------------------------------------
# TIME-SERIES CHARTS (AFFECTED BY GLOBAL FILTER)
# ---------------------------------------------------------
spend_trend_df = prepare_trend_data(filtered_df, value_col="Cost", selected_year=selected_year)
impression_trend_df = prepare_trend_data(filtered_df, value_col="Impressions", selected_year=selected_year)

chart_col1, chart_col2 = st.columns(2)

with chart_col1:
    spend_fig = build_trend_chart(
        trend_df=spend_trend_df,
        title="Advertising Spend Trend",
        value_type="spend",
        selected_year=selected_year
    )
    st.plotly_chart(spend_fig, use_container_width=True, config={"displayModeBar": False})

with chart_col2:
    impression_fig = build_trend_chart(
        trend_df=impression_trend_df,
        title="Impressions Trend",
        value_type="impressions",
        selected_year=selected_year
    )
    st.plotly_chart(impression_fig, use_container_width=True, config={"displayModeBar": False})

# ---------------------------------------------------------
# CAMPAIGN LIST SECTION
# ---------------------------------------------------------
st.markdown('<div class="section-title">Campaign List</div>', unsafe_allow_html=True)

search_col, sort_col = st.columns([7.8, 1.4], gap="small")

with search_col:
    st.markdown(
        '<div class="search-box-label">Search by company, campaign name, or campaign ID.</div>',
        unsafe_allow_html=True
    )
    search_keyword = st.text_input(
        "Search campaign list",
        placeholder="Search company, campaign name, or campaign ID",
        label_visibility="collapsed"
    )

with sort_col:
    st.markdown('<div class="sort-box-label">Sort by</div>', unsafe_allow_html=True)
    sort_option = st.selectbox(
        "Sort by",
        options=[
            "Latest Start Date",
            "Earliest Start Date",
            "Highest Spend",
            "Highest Impressions",
            "Lowest CPM",
            "Highest Ads Count",
        ],
        key="campaign_sort_option",
        label_visibility="collapsed"
    )

campaign_list_df = prepare_campaign_list(filtered_df)
campaign_list_df = apply_campaign_search(campaign_list_df, search_keyword)
campaign_list_df = apply_campaign_sort(campaign_list_df, sort_option)

st.caption(f"Showing {len(campaign_list_df):,} campaign(s)")

if campaign_list_df.empty:
    st.markdown(
        '<div class="campaign-empty">No campaigns found for the current filter and search.</div>',
        unsafe_allow_html=True
    )
else:
    render_campaign_list_header()
    for _, row in campaign_list_df.iterrows():
        render_campaign_card(row)