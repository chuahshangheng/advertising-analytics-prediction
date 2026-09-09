import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

from datetime import date, datetime
from io import BytesIO
from pathlib import Path

from utils.load_models import load_all_models
from utils.hybrid_predict import hybrid_predict


# =========================================================
# PAGE TITLE
# =========================================================
st.title("Prediction")

BASE_DIR = Path(__file__).resolve().parent.parent
TEMPLATE_PATH = BASE_DIR / "template" / "template_ads.xlsx"

# =========================================================
# TEMPLATE COLUMN NAMES (EXCEL HEADERS)
# =========================================================
REQUIRED_COLUMNS = [
    "Ad Platform",
    "Ad Format",
    "Objective",
    "Cost (RM)",
    "Start Date (d/m/yyyy)",
    "End Date (d/m/yyyy)",
    "Target Gender",
    "Target Age Group",
    "Target Interests",
]

# =========================================================
# INTERNAL COLUMN MAPPING
# =========================================================
COLUMN_MAPPING = {
    "Ad Platform": "AdPlatform",
    "Ad Format": "AdFormat",
    "Objective": "Objective",
    "Cost (RM)": "Cost",
    "Start Date (d/m/yyyy)": "StartDate",
    "End Date (d/m/yyyy)": "EndDate",
    "Target Gender": "TargetGender",
    "Target Age Group": "TargetAgeGroup",
    "Target Interests": "TargetInterests",
}

AD_PLATFORM_OPTIONS = ["Facebook", "Instagram"]
AD_FORMAT_OPTIONS = ["Carousel Ad", "Image Ad", "Story Ad", "Video Ad"]
OBJECTIVE_OPTIONS = [
    "Brand Awareness",
    "Reach",
    "Engagement",
    "Traffic",
    "Lead Generation",
    "Sales",
]
TARGET_GENDER_OPTIONS = ["All", "Male", "Female"]
TARGET_AGE_OPTIONS = ["18-24", "25-34", "35-44", "All"]


# =========================================================
# LOAD MODELS
# =========================================================
@st.cache_resource
def get_loaded_models():
    return load_all_models()


loaded = get_loaded_models()


# =========================================================
# HELPER: BUILD DAILY IMPRESSION CHART
# =========================================================
def build_daily_impression_chart(duration, current_cost, current_prediction):
    current_daily_budget = current_cost / duration
    current_daily_impressions = current_prediction / duration

    if current_daily_budget <= 0:
        current_daily_budget = 1

    max_display_budget = max(current_daily_budget * 1.8, 100)
    x_line = np.linspace(0, max_display_budget, 50)

    slope = current_daily_impressions / current_daily_budget
    y_line = slope * x_line

    hover_mask = np.abs(x_line - current_daily_budget) > max_display_budget * 0.03
    hover_x = x_line[hover_mask]
    hover_y = y_line[hover_mask]

    fig = go.Figure()

    fig.add_trace(
        go.Scatter(
            x=x_line,
            y=y_line,
            mode="lines",
            line=dict(width=3, color="#14b8a6"),
            hoverinfo="skip",
            showlegend=False,
        )
    )

    fig.add_trace(
        go.Scatter(
            x=hover_x,
            y=hover_y,
            mode="markers",
            marker=dict(
                size=16,
                color="rgba(0,0,0,0)",
                line=dict(width=0, color="rgba(0,0,0,0)")
            ),
            customdata=np.stack([hover_x, hover_y], axis=-1),
            hovertemplate=(
                "<b>Daily budget</b>: RM%{customdata[0]:,.2f}<br>"
                "<b>Daily impressions</b>: %{customdata[1]:,.0f}"
                "<extra></extra>"
            ),
            hoverlabel=dict(
                bgcolor="white",
                bordercolor="#94a3b8",
                font_size=14,
                font_family="Arial",
                font_color="#1f2937",
            ),
            showlegend=False,
        )
    )

    fig.add_trace(
        go.Scatter(
            x=[current_daily_budget],
            y=[current_daily_impressions],
            mode="markers",
            marker=dict(
                size=22,
                color="rgba(20, 184, 166, 0.20)",
                line=dict(width=0, color="rgba(20, 184, 166, 0)")
            ),
            hoverinfo="skip",
            showlegend=False,
        )
    )

    fig.add_trace(
        go.Scatter(
            x=[current_daily_budget],
            y=[current_daily_impressions],
            mode="markers",
            marker=dict(
                size=11,
                color="#14b8a6",
                line=dict(width=2, color="white")
            ),
            hoverinfo="skip",
            showlegend=False,
        )
    )

    fig.add_annotation(
        x=current_daily_budget,
        y=current_daily_impressions,
        text=(
            f"<b>Daily budget</b>: <b>RM{current_daily_budget:,.2f}</b><br>"
            f"<b>Daily impressions</b>: <b>{current_daily_impressions:,.0f}</b>"
        ),
        showarrow=False,
        xanchor="left",
        yanchor="middle",
        xshift=14,
        bgcolor="#d1fae5",
        bordercolor="#0f766e",
        borderwidth=2,
        font=dict(size=14, color="#065f46"),
        align="left",
    )

    fig.add_vline(
        x=current_daily_budget,
        line_width=1,
        line_color="#9ca3af"
    )

    fig.update_layout(
        title="Estimated daily impressions",
        xaxis_title="Budget per day (RM)",
        yaxis_title="Daily impressions",
        height=420,
        margin=dict(l=20, r=20, t=50, b=20),
        template="plotly_white",
        showlegend=False,
        hovermode="closest",
        hoverdistance=100,
    )

    fig.update_xaxes(
        tickformat=",.0f",
        showgrid=True,
        tickfont=dict(size=11, color="#64748b"),
        title_font=dict(size=13, color="#475569")
    )

    fig.update_yaxes(
        tickformat=",.0f",
        showgrid=True,
        tickfont=dict(size=11, color="#64748b"),
        title_font=dict(size=13, color="#475569")
    )

    return fig


# =========================================================
# HELPER FUNCTIONS
# =========================================================
def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def canonicalize_value(value, valid_options):
    text = normalize_text(value)
    if text == "":
        return ""

    option_map = {str(option).strip().lower(): option for option in valid_options}
    return option_map.get(text.lower())


def load_template_bytes():
    if TEMPLATE_PATH.exists():
        return TEMPLATE_PATH.read_bytes()
    return None


def standardize_uploaded_columns(df):
    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]
    return df


def check_required_columns(df):
    missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    return missing_cols


def convert_to_internal_columns(df):
    df = df.copy()
    return df.rename(columns=COLUMN_MAPPING)


def format_date_columns_for_display(df):
    df = df.copy()
    date_cols = ["Start Date (d/m/yyyy)", "End Date (d/m/yyyy)"]

    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce").dt.strftime("%d/%m/%Y")
            df[col] = df[col].fillna("")

    return df


def format_numeric_columns_for_display(df):
    df = df.copy()

    if "Cost (RM)" in df.columns:
        df["Cost (RM)"] = pd.to_numeric(df["Cost (RM)"], errors="coerce").apply(
            lambda x: f"{x:,.2f}" if pd.notna(x) else ""
        )

    if "Predicted Impression" in df.columns:
        df["Predicted Impression"] = pd.to_numeric(df["Predicted Impression"], errors="coerce").apply(
            lambda x: f"{int(round(x)):,}" if pd.notna(x) else ""
        )

    if "Predicted CPM (RM)" in df.columns:
        df["Predicted CPM (RM)"] = pd.to_numeric(df["Predicted CPM (RM)"], errors="coerce").apply(
            lambda x: f"{x:,.2f}" if pd.notna(x) else ""
        )

    if "Campaign Duration (Days)" in df.columns:
        df["Campaign Duration (Days)"] = pd.to_numeric(
            df["Campaign Duration (Days)"], errors="coerce"
        ).apply(
            lambda x: f"{int(round(x)):,}" if pd.notna(x) else ""
        )

    return df


def validate_and_predict_row(row, loaded_models):
    target_interest_options = loaded_models["cat_meta"]["target_interest_options"]
    errors = []

    ad_platform_raw = row.get("AdPlatform")
    ad_format_raw = row.get("AdFormat")
    objective_raw = row.get("Objective")
    target_gender_raw = row.get("TargetGender")
    target_age_raw = row.get("TargetAgeGroup")
    target_interests_raw = row.get("TargetInterests")

    ad_platform = canonicalize_value(ad_platform_raw, AD_PLATFORM_OPTIONS)
    ad_format = canonicalize_value(ad_format_raw, AD_FORMAT_OPTIONS)
    objective = canonicalize_value(objective_raw, OBJECTIVE_OPTIONS)
    target_gender = canonicalize_value(target_gender_raw, TARGET_GENDER_OPTIONS)
    target_age = canonicalize_value(target_age_raw, TARGET_AGE_OPTIONS)
    target_interests = canonicalize_value(target_interests_raw, target_interest_options)

    if normalize_text(ad_platform_raw) == "":
        errors.append("Missing Ad Platform")
    elif ad_platform is None:
        errors.append("Invalid Ad Platform")

    if normalize_text(ad_format_raw) == "":
        errors.append("Missing Ad Format")
    elif ad_format is None:
        errors.append("Invalid Ad Format")

    if normalize_text(objective_raw) == "":
        errors.append("Missing Objective")
    elif objective is None:
        errors.append("Invalid Objective")

    if normalize_text(target_gender_raw) == "":
        errors.append("Missing Target Gender")
    elif target_gender is None:
        errors.append("Invalid Target Gender")

    if normalize_text(target_age_raw) == "":
        errors.append("Missing Target Age Group")
    elif target_age is None:
        errors.append("Invalid Target Age Group")

    if normalize_text(target_interests_raw) == "":
        errors.append("Missing Target Interests")
    elif target_interests is None:
        errors.append("Invalid Target Interests")

    cost_raw = row.get("Cost")
    cost_value = pd.to_numeric(pd.Series([cost_raw]), errors="coerce").iloc[0]

    if normalize_text(cost_raw) == "":
        errors.append("Missing Cost")
    elif pd.isna(cost_value):
        errors.append("Invalid Cost")
    elif cost_value < 0:
        errors.append("Cost cannot be negative")

    start_raw = row.get("StartDate")
    end_raw = row.get("EndDate")

    start_date = pd.to_datetime(start_raw, errors="coerce", dayfirst=True)
    end_date = pd.to_datetime(end_raw, errors="coerce", dayfirst=True)

    if normalize_text(start_raw) == "":
        errors.append("Missing Start Date")
    elif pd.isna(start_date):
        errors.append("Invalid Start Date")

    if normalize_text(end_raw) == "":
        errors.append("Missing End Date")
    elif pd.isna(end_date):
        errors.append("Invalid End Date")

    duration = None
    month_num = None

    if not pd.isna(start_date) and not pd.isna(end_date):
        start_date = start_date.date()
        end_date = end_date.date()

        if end_date < start_date:
            errors.append("End date must be later than or equal to Start date")
        else:
            duration = (end_date - start_date).days + 1
            month_num = start_date.month

            if not pd.isna(cost_value):
                min_required_cost = duration * 20
                if cost_value < min_required_cost:
                    errors.append(
                        f"Minimum cost must be RM{min_required_cost:,.2f} "
                        f"for {duration} campaign days (RM20/day)"
                    )

    if errors:
        return {
            "is_valid": False,
            "predicted_impressions": pd.NA,
            "predicted_cpm": pd.NA,
            "campaign_duration_days": pd.NA,
            "note": "; ".join(errors),
            "model_used": "",
            "range_status": "",
        }

    sample_input = pd.DataFrame([{
        "AdFormat": ad_format,
        "Cost": float(cost_value),
        "Objective": objective,
        "TargetGender": target_gender,
        "TargetAgeGroup": target_age,
        "TargetInterests": target_interests,
        "AdPlatform": ad_platform,
        "CampaignDurationDays": duration,
        "Year": 2025,
        "MonthNum": month_num,
    }])

    result = hybrid_predict(sample_input, loaded_models)
    predicted_impressions = result["predicted_impressions"]

    if result["model_used"] == "CatBoost":
        predicted_cpm = round((float(cost_value) / predicted_impressions) * 1000, 2) if predicted_impressions > 0 else 0.0
    else:
        predicted_cpm = result["predicted_cpm"]

    note = "Primary model used"
    if result["warning"]:
        note = result["warning"]

    return {
        "is_valid": True,
        "predicted_impressions": predicted_impressions,
        "predicted_cpm": predicted_cpm,
        "campaign_duration_days": duration,
        "note": note,
        "model_used": result["model_used"],
        "range_status": result["range_status"],
    }


def process_uploaded_file(original_df, loaded_models):
    display_df = original_df.copy()
    internal_df = convert_to_internal_columns(original_df)

    display_df["Predicted Impression"] = pd.NA
    display_df["Predicted CPM (RM)"] = pd.NA
    display_df["Campaign Duration (Days)"] = pd.NA
    display_df["Note"] = ""

    summary = {
        "total_rows": len(display_df),
        "valid_rows": 0,
        "invalid_rows": 0,
        "catboost_rows": 0,
        "catboost_cpm_rows": 0,
    }

    for idx, row in internal_df.iterrows():
        row_result = validate_and_predict_row(row, loaded_models)

        display_df.at[idx, "Predicted Impression"] = row_result["predicted_impressions"]
        display_df.at[idx, "Predicted CPM (RM)"] = row_result["predicted_cpm"]
        display_df.at[idx, "Campaign Duration (Days)"] = row_result["campaign_duration_days"]
        display_df.at[idx, "Note"] = row_result["note"]

        if row_result["is_valid"]:
            summary["valid_rows"] += 1

            if row_result["model_used"] == "CatBoost":
                summary["catboost_rows"] += 1
            elif row_result["model_used"] == "CatBoost CPM":
                summary["catboost_cpm_rows"] += 1
        else:
            summary["invalid_rows"] += 1

    return display_df, summary


def prepare_result_dataframe_for_display(df):
    display_df = df.copy()
    display_df = format_date_columns_for_display(display_df)
    display_df = format_numeric_columns_for_display(display_df)
    return display_df


def dataframe_to_excel_bytes(df):
    output = BytesIO()

    export_df = df.copy()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="prediction_results")

        worksheet = writer.sheets["prediction_results"]

        header_map = {}
        for idx, col_name in enumerate(export_df.columns, start=1):
            header_map[col_name] = idx

        for row_idx in range(2, len(export_df) + 2):
            if "Cost (RM)" in header_map:
                worksheet.cell(row=row_idx, column=header_map["Cost (RM)"]).number_format = '#,##0.00'

            if "Predicted Impression" in header_map:
                worksheet.cell(row=row_idx, column=header_map["Predicted Impression"]).number_format = '#,##0'

            if "Predicted CPM (RM)" in header_map:
                worksheet.cell(row=row_idx, column=header_map["Predicted CPM (RM)"]).number_format = '#,##0.00'

            if "Campaign Duration (Days)" in header_map:
                worksheet.cell(row=row_idx, column=header_map["Campaign Duration (Days)"]).number_format = '#,##0'

        for column_cells in worksheet.columns:
            max_length = 0
            column_letter = column_cells[0].column_letter

            for cell in column_cells:
                cell_value = "" if cell.value is None else str(cell.value)
                max_length = max(max_length, len(cell_value))

            worksheet.column_dimensions[column_letter].width = min(max_length + 2, 40)

    output.seek(0)
    return output.getvalue()


# =========================================================
# SINGLE PREDICTION TAB
# =========================================================
def render_single_prediction():
    left, right = st.columns([1, 1])

    with left:
        st.subheader("Campaign Input")

        target_interest_options = loaded["cat_meta"]["target_interest_options"]

        with st.form("prediction_form"):
            ad_platform = st.selectbox("Ad Platform", AD_PLATFORM_OPTIONS)
            ad_format = st.selectbox("Ad Format", AD_FORMAT_OPTIONS)
            objective = st.selectbox("Objective", OBJECTIVE_OPTIONS)

            cost = st.number_input(
                "Cost (RM)",
                min_value=0.0,
                value=3000.0,
                step=100.0,
                format="%.2f"
            )

            start_date = st.date_input("Start date", value=date.today())
            end_date = st.date_input("End date", value=date.today())
            target_gender = st.selectbox("Target Gender", TARGET_GENDER_OPTIONS)
            target_age = st.selectbox("Target Age Group", TARGET_AGE_OPTIONS)
            target_interests = st.selectbox(
                "Target Interests",
                options=target_interest_options,
                index=0
            )

            submitted = st.form_submit_button("Predict")

    with right:
        st.subheader("Prediction Result")

        if submitted:
            if end_date < start_date:
                st.error("Invalid input: End date must be later than or equal to Start date.")
            else:
                year = 2025
                duration = (end_date - start_date).days + 1
                month_num = start_date.month
                min_required_cost = duration * 20

                if cost < min_required_cost:
                    st.error(
                        f"Invalid input: Minimum cost must be RM{min_required_cost:,.2f} "
                        f"for {duration} campaign days (RM20/day)."
                    )
                    st.write(f"**Minimum Required Cost:** RM{min_required_cost:,.2f}")
                else:
                    sample_input = pd.DataFrame([{
                        "AdFormat": ad_format,
                        "Cost": cost,
                        "Objective": objective,
                        "TargetGender": target_gender,
                        "TargetAgeGroup": target_age,
                        "TargetInterests": target_interests,
                        "AdPlatform": ad_platform,
                        "CampaignDurationDays": duration,
                        "Year": year,
                        "MonthNum": month_num,
                    }])

                    result = hybrid_predict(sample_input, loaded)
                    predicted_impressions = result["predicted_impressions"]

                    if result["model_used"] == "CatBoost":
                        predicted_cpm = (cost / predicted_impressions) * 1000 if predicted_impressions > 0 else 0.0
                    else:
                        predicted_cpm = result["predicted_cpm"]

                    metric_col1, metric_col2 = st.columns(2)

                    with metric_col1:
                        st.metric("Predicted Impressions", f"{predicted_impressions:,}")

                    with metric_col2:
                        st.metric("Predicted CPM", f"RM{predicted_cpm:,.2f}")

                    st.write(f"**Model Used:** {result['model_used']}")
                    st.write(f"**Range Status:** {result['range_status']}")
                    st.write(f"**Campaign Duration:** {duration:,} {'day' if duration == 1 else 'days'}")
                    st.write(f"**Cost:** RM{cost:,.2f}")

                    daily_chart = build_daily_impression_chart(
                        duration=duration,
                        current_cost=cost,
                        current_prediction=predicted_impressions,
                    )

                    st.plotly_chart(
                        daily_chart,
                        use_container_width=True,
                        config={"displayModeBar": False}
                    )

                    st.caption("These are estimates and don't guarantee results.")

                    if result["warning"]:
                        st.warning(result["warning"])
                    else:
                        st.success("Primary model used.")
        else:
            st.info("Fill in the campaign details and click Predict.")


# =========================================================
# BATCH UPLOAD TAB
# =========================================================
def render_batch_prediction():
    st.subheader("Batch Upload Prediction")
    st.write(
        "Download the template, fill in one campaign per row, upload the completed file, "
        "and export the prediction results."
    )

    template_bytes = load_template_bytes()

    if template_bytes is not None:
        st.download_button(
            label="Download Template",
            data=template_bytes,
            file_name="template_ads.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning("Template file not found. Please place 'template_ads.xlsx' inside the 'template' folder.")

    st.markdown("**Required columns:**")
    st.caption(", ".join(REQUIRED_COLUMNS))

    uploaded_file = st.file_uploader(
        "Upload completed template",
        type=["xlsx", "xlsm"]
    )

    if uploaded_file is None:
        st.info("Upload a completed Excel file to start batch prediction.")
        return

    try:
        uploaded_df = pd.read_excel(uploaded_file)
        uploaded_df = standardize_uploaded_columns(uploaded_df)
    except Exception as e:
        st.error(f"Unable to read the uploaded file. Error: {e}")
        return

    st.markdown("**Uploaded file preview**")
    preview_df = format_date_columns_for_display(uploaded_df)
    preview_df = format_numeric_columns_for_display(preview_df)
    st.dataframe(preview_df, use_container_width=True)

    missing_cols = check_required_columns(uploaded_df)
    if missing_cols:
        st.error("Missing required columns: " + ", ".join(missing_cols))
        return

    if st.button("Run Batch Prediction", type="primary"):
        result_df, summary = process_uploaded_file(uploaded_df, loaded)
        display_result_df = prepare_result_dataframe_for_display(result_df)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Total Rows", f"{summary['total_rows']:,}")
        c2.metric("Valid Rows", f"{summary['valid_rows']:,}")
        c3.metric("Invalid Rows", f"{summary['invalid_rows']:,}")
        c4.metric("CatBoost Rows", f"{summary['catboost_rows']:,}")
        c5.metric("CatBoost CPM Rows", f"{summary['catboost_cpm_rows']:,}")

        st.markdown("**Processed result preview**")
        st.dataframe(display_result_df, use_container_width=True)

        output_bytes = dataframe_to_excel_bytes(result_df)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_name = f"prediction_results_{timestamp}.xlsx"

        st.download_button(
            label="Download Prediction Results",
            data=output_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# =========================================================
# SUBMENU TABS
# =========================================================
tab1, tab2 = st.tabs(["Single Prediction", "Batch Upload Prediction"])

with tab1:
    render_single_prediction()

with tab2:
    render_batch_prediction()