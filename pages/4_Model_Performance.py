# =========================================================
# MODEL PERFORMANCE PAGE
# Start directly with model selection
# Includes "All Model Comparison" option
# Includes Linear Regression in comparison
# Shows before/after tuning metrics from exported summary CSV
# Controlled image size for diagnostic plots
# =========================================================

import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path


# =========================================================
# 1. PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="Model Performance",
    layout="wide"
)


# =========================================================
# 2. PATH SETTINGS
# =========================================================

BASE_DIR = Path(__file__).resolve().parents[1]
MODEL_OUTPUT_DIR = BASE_DIR / "model_outputs"


# =========================================================
# 3. IMAGE DISPLAY SETTINGS
# =========================================================

IMAGE_WIDTH = 560
SINGLE_IMAGE_WIDTH = 650


# =========================================================
# 4. HELPER FUNCTIONS
# =========================================================

def safe_model_name(model_name):
    """
    Convert model name into file-safe format.
    Must match the export file naming.
    """
    return (
        str(model_name)
        .replace(" ", "_")
        .replace("/", "_")
        .replace("\\", "_")
        .replace("(", "")
        .replace(")", "")
    )


def format_number(value, decimals=4):
    """
    Format numeric value safely.
    """
    try:
        return f"{float(value):,.{decimals}f}"
    except Exception:
        return "-"


def get_plot_path(task_name, file_name):
    """
    Get plot path.
    """
    return MODEL_OUTPUT_DIR / task_name / "figures" / file_name


def get_csv_path(task_name, folder_name, file_name):
    """
    Get CSV path.
    """
    return MODEL_OUTPUT_DIR / task_name / folder_name / file_name


def is_linear_model(model_name):
    """
    Identify Linear Regression model.
    Works for:
    - Linear Regression
    - LinearRegression_CPM
    - Linear Regression CPM
    """
    name = str(model_name).lower()
    return "linear" in name and "regression" in name


def get_model_row(eval_df, model_name):
    """
    Get selected model row from evaluation table.
    """
    row = eval_df[eval_df["Model"] == model_name]

    if row.empty:
        return None

    return row.iloc[0]


def show_table(df, title):
    """
    Show dataframe table.
    """
    st.markdown(f"#### {title}")

    if df is not None and not df.empty:
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info(f"{title} not available.")


def show_image(path, title, width=IMAGE_WIDTH):
    """
    Show image with controlled size.
    """
    st.markdown(f"#### {title}")

    if path.exists():
        st.image(str(path), width=width)
    else:
        st.info(f"{title} not available.")


# =========================================================
# 5. METRIC TABLE FROM SUMMARY CSV
# =========================================================

def load_before_after_metrics(task_name, model_name):
    """
    Load before-after metrics CSV exported from modelling notebook.

    Expected path:
    model_outputs/<task_name>/summary/<model_name>_before_after_metrics.csv
    """

    model_safe = safe_model_name(model_name)

    metrics_path = get_csv_path(
        task_name,
        "summary",
        f"{model_safe}_before_after_metrics.csv"
    )

    if metrics_path.exists():
        return pd.read_csv(metrics_path)

    return None


def format_metric_table(metrics_df):
    """
    Format metric table for display.
    """

    if metrics_df is None or metrics_df.empty:
        return None

    display_df = metrics_df.copy()

    # Remove Model column if exists, because selected model is already shown in heading
    if "Model" in display_df.columns:
        display_df = display_df.drop(columns=["Model"])

    rename_cols = {
        "Train_R2": "Train R²",
        "Test_R2": "Test R²",
        "Train_MAE": "Train MAE",
        "Test_MAE": "Test MAE",
        "Train_RMSE": "Train RMSE",
        "Test_RMSE": "Test RMSE",
        "Train_MAPE": "Train MAPE (%)",
        "Test_MAPE": "Test MAPE (%)"
    }

    display_df = display_df.rename(columns=rename_cols)

    for col in [
        "Train R²",
        "Test R²",
        "Train MAE",
        "Test MAE",
        "Train RMSE",
        "Test RMSE",
        "Train MAPE (%)",
        "Test MAPE (%)"
    ]:
        if col in display_df.columns:
            if "R²" in col:
                display_df[col] = display_df[col].apply(lambda x: format_number(x, 4))
            else:
                display_df[col] = display_df[col].apply(lambda x: format_number(x, 2))

    return display_df


# =========================================================
# 6. FALLBACK METRICS FROM EVALUATION TABLE
# =========================================================

def create_fallback_metric_table(eval_df, model_name):
    """
    Fallback metric table if before_after_metrics.csv is not found.
    """

    row = get_model_row(eval_df, model_name)

    if row is None:
        return None

    version_name = row.get("Search_Method", "Baseline Only")

    return pd.DataFrame({
        "Model Version": [version_name],
        "Train R²": [format_number(row.get("Train_R2"), 4)],
        "Test R²": [format_number(row.get("Test_R2"), 4)],
        "Train MAE": [format_number(row.get("Train_MAE"), 2)],
        "Test MAE": [format_number(row.get("Test_MAE"), 2)],
        "Train RMSE": [format_number(row.get("Train_RMSE"), 2)],
        "Test RMSE": [format_number(row.get("Test_RMSE"), 2)],
        "Train MAPE (%)": [format_number(row.get("Train_MAPE"), 2)],
        "Test MAPE (%)": [format_number(row.get("Test_MAPE"), 2)]
    })


# =========================================================
# 7. ADJUSTED BAR CHART
# =========================================================

def create_adjusted_bar_chart(df, x_col, y_col, title, y_title, higher_is_better=True):
    """
    Create Altair bar chart with adjusted y-axis range.
    Uses custom baseline so bars remain visible when y-axis is zoomed.
    """

    chart_df = df[[x_col, y_col]].copy()
    chart_df[y_col] = pd.to_numeric(chart_df[y_col], errors="coerce")
    chart_df = chart_df.dropna(subset=[y_col])

    if chart_df.empty:
        st.info(f"{title} is not available.")
        return

    min_value = chart_df[y_col].min()
    max_value = chart_df[y_col].max()

    if min_value == max_value:
        y_min = min_value * 0.95
        y_max = max_value * 1.05
    else:
        value_range = max_value - min_value
        y_min = min_value - (value_range * 0.20)
        y_max = max_value + (value_range * 0.20)

    if not higher_is_better:
        y_min = max(0, y_min)

    chart_df["Baseline"] = y_min

    sort_order = chart_df.sort_values(
        by=y_col,
        ascending=not higher_is_better
    )[x_col].tolist()

    bar_chart = (
        alt.Chart(chart_df)
        .mark_bar()
        .encode(
            x=alt.X(
                f"{x_col}:N",
                sort=sort_order,
                title="Model",
                axis=alt.Axis(labelAngle=-35)
            ),
            y=alt.Y(
                f"{y_col}:Q",
                title=y_title,
                scale=alt.Scale(domain=[y_min, y_max], zero=False)
            ),
            y2=alt.Y2("Baseline:Q"),
            tooltip=[
                alt.Tooltip(f"{x_col}:N", title="Model"),
                alt.Tooltip(f"{y_col}:Q", title=y_title, format=",.4f")
            ]
        )
        .properties(
            title=title,
            height=320
        )
    )

    text_chart = (
        alt.Chart(chart_df)
        .mark_text(
            align="center",
            baseline="bottom",
            dy=-5
        )
        .encode(
            x=alt.X(
                f"{x_col}:N",
                sort=sort_order
            ),
            y=alt.Y(
                f"{y_col}:Q",
                scale=alt.Scale(domain=[y_min, y_max], zero=False)
            ),
            text=alt.Text(f"{y_col}:Q", format=",.4f")
        )
    )

    st.altair_chart(bar_chart + text_chart, use_container_width=True)


# =========================================================
# 8. ALL MODEL COMPARISON
# =========================================================

def show_all_model_comparison(eval_df, task_title):
    """
    Show comparison of all final models.
    Includes Linear Regression baseline and tuned models.
    """

    st.divider()
    st.subheader(f"{task_title} - All Model Comparison")

    comparison_df = eval_df.copy()

    if comparison_df.empty:
        st.info("No model results found.")
        return

    metric_cols = [
        "Train_R2",
        "Test_R2",
        "Train_MAE",
        "Test_MAE",
        "Train_RMSE",
        "Test_RMSE",
        "Train_MAPE",
        "Test_MAPE"
    ]

    for col in metric_cols:
        if col in comparison_df.columns:
            comparison_df[col] = pd.to_numeric(comparison_df[col], errors="coerce")

    comparison_df = comparison_df.sort_values(
        by=["Test_RMSE", "Test_R2"],
        ascending=[True, False]
    ).reset_index(drop=True)

    comparison_df.insert(0, "Rank", range(1, len(comparison_df) + 1))

    best_model = comparison_df.iloc[0]

    st.markdown(
        """
        This section compares all final model results, including the baseline Linear Regression model 
        and the models improved through hyperparameter tuning. The ranking is mainly based on the 
        lowest Test RMSE, while Test R², Test MAE and Test MAPE are used as supporting metrics.
        """
    )

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Best Model", best_model["Model"])

    with col2:
        st.metric("Test R²", format_number(best_model.get("Test_R2"), 4))

    with col3:
        st.metric("Test RMSE", format_number(best_model.get("Test_RMSE"), 2))

    with col4:
        st.metric("Test MAPE (%)", format_number(best_model.get("Test_MAPE"), 2))

    display_cols = [
        "Rank",
        "Model",
        "Target",
        "Search_Method",
        "Train_R2",
        "Test_R2",
        "Train_MAE",
        "Test_MAE",
        "Train_RMSE",
        "Test_RMSE",
        "Train_MAPE",
        "Test_MAPE"
    ]

    existing_cols = [col for col in display_cols if col in comparison_df.columns]
    display_df = comparison_df[existing_cols].copy()

    display_df = display_df.rename(columns={
        "Search_Method": "Search Method",
        "Train_R2": "Train R²",
        "Test_R2": "Test R²",
        "Train_MAE": "Train MAE",
        "Test_MAE": "Test MAE",
        "Train_RMSE": "Train RMSE",
        "Test_RMSE": "Test RMSE",
        "Train_MAPE": "Train MAPE (%)",
        "Test_MAPE": "Test MAPE (%)"
    })

    for col in [
        "Train R²",
        "Test R²",
        "Train MAE",
        "Test MAE",
        "Train RMSE",
        "Test RMSE",
        "Train MAPE (%)",
        "Test MAPE (%)"
    ]:
        if col in display_df.columns:
            if "R²" in col:
                display_df[col] = display_df[col].apply(lambda x: format_number(x, 4))
            else:
                display_df[col] = display_df[col].apply(lambda x: format_number(x, 2))

    st.markdown("#### Ranked Model Comparison")
    st.dataframe(display_df, use_container_width=True, hide_index=True)

    st.markdown("#### Model Metric Comparison")

    st.caption(
        "The chart y-axis is adjusted to focus on the model performance range, "
        "so small differences between models are easier to compare."
    )

    col_a, col_b = st.columns(2)

    with col_a:
        if "Test_R2" in comparison_df.columns:
            create_adjusted_bar_chart(
                df=comparison_df,
                x_col="Model",
                y_col="Test_R2",
                title="Test R² Comparison",
                y_title="Test R²",
                higher_is_better=True
            )

    with col_b:
        if "Test_RMSE" in comparison_df.columns:
            create_adjusted_bar_chart(
                df=comparison_df,
                x_col="Model",
                y_col="Test_RMSE",
                title="Test RMSE Comparison",
                y_title="Test RMSE",
                higher_is_better=False
            )

    col_c, col_d = st.columns(2)

    with col_c:
        if "Test_MAE" in comparison_df.columns:
            create_adjusted_bar_chart(
                df=comparison_df,
                x_col="Model",
                y_col="Test_MAE",
                title="Test MAE Comparison",
                y_title="Test MAE",
                higher_is_better=False
            )

    with col_d:
        if "Test_MAPE" in comparison_df.columns:
            create_adjusted_bar_chart(
                df=comparison_df,
                x_col="Model",
                y_col="Test_MAPE",
                title="Test MAPE Comparison",
                y_title="Test MAPE (%)",
                higher_is_better=False
            )

    st.info(
        f"{best_model['Model']} is ranked as the best overall model because it achieved "
        f"the lowest Test RMSE. However, Test R², Test MAE and Test MAPE should also be "
        f"reviewed together to support the final model selection."
    )


# =========================================================
# 9. SELECTED MODEL METRICS
# =========================================================

def create_selected_model_metric_table(eval_df, model_name, task_name):
    """
    Create selected model metric table.

    Priority:
    1. Read full before/after metrics from summary CSV
    2. Fallback to evaluation_table.csv if summary CSV is not available
    """

    metrics_df = load_before_after_metrics(
        task_name=task_name,
        model_name=model_name
    )

    if metrics_df is not None:
        return format_metric_table(metrics_df)

    return create_fallback_metric_table(
        eval_df=eval_df,
        model_name=model_name
    )


def show_metric_cards(eval_df, model_name):
    """
    Show final selected model metric cards.
    """

    row = get_model_row(eval_df, model_name)

    if row is None:
        return

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Test R²", format_number(row.get("Test_R2"), 4))

    with col2:
        st.metric("Test MAE", format_number(row.get("Test_MAE"), 2))

    with col3:
        st.metric("Test RMSE", format_number(row.get("Test_RMSE"), 2))

    with col4:
        st.metric("Test MAPE (%)", format_number(row.get("Test_MAPE"), 2))


# =========================================================
# 10. PLOT DISPLAY FUNCTIONS
# =========================================================

def show_before_after_images(task_name, model_name, image_suffix, title):
    """
    Show before and after tuning images side-by-side.
    """

    model_safe = safe_model_name(model_name)

    before_path = get_plot_path(
        task_name,
        f"{model_safe}_Before_Tuning_{image_suffix}.png"
    )

    after_path = get_plot_path(
        task_name,
        f"{model_safe}_After_Tuning_{image_suffix}.png"
    )

    st.markdown(f"#### {title}")

    col1, col2 = st.columns(2)

    with col1:
        st.caption("Before Tuning")
        if before_path.exists():
            st.image(str(before_path), width=IMAGE_WIDTH)
        else:
            st.info("Before tuning plot not available.")

    with col2:
        st.caption("After Tuning")
        if after_path.exists():
            st.image(str(after_path), width=IMAGE_WIDTH)
        else:
            st.info("After tuning plot not available.")


def show_baseline_image(task_name, model_name, image_suffix, title):
    """
    Show baseline image for Linear Regression.
    """

    model_safe = safe_model_name(model_name)

    possible_paths = [
        get_plot_path(task_name, f"{model_safe}_Before_Tuning_{image_suffix}.png"),
        get_plot_path(task_name, f"{model_safe}_{image_suffix}.png")
    ]

    st.markdown(f"#### {title}")

    for path in possible_paths:
        if path.exists():
            st.image(str(path), width=IMAGE_WIDTH)
            return

    st.info(f"{title} not available.")


def show_learning_curve(task_name, model_name):
    """
    Show learning curve.
    """

    model_safe = safe_model_name(model_name)

    if is_linear_model(model_name):
        possible_paths = [
            get_plot_path(task_name, f"{model_safe}_learning_curve.png"),
            get_plot_path(task_name, f"{model_safe}_Before_Tuning_learning_curve.png")
        ]
    else:
        possible_paths = [
            get_plot_path(task_name, f"{model_safe}_After_Tuning_learning_curve.png"),
            get_plot_path(task_name, f"{model_safe}_learning_curve.png")
        ]

    st.markdown("#### Learning Curve")

    for path in possible_paths:
        if path.exists():
            st.image(str(path), width=IMAGE_WIDTH)
            return

    st.info("Learning curve not available.")


def show_feature_importance(task_name, model_name):
    """
    Show feature importance or coefficient importance.
    """

    model_safe = safe_model_name(model_name)

    if is_linear_model(model_name):
        possible_paths = [
            get_plot_path(task_name, f"{model_safe}_Before_Tuning_feature_importance.png"),
            get_plot_path(task_name, f"{model_safe}_feature_importance.png")
        ]
        title = "Coefficient Importance"
    else:
        possible_paths = [
            get_plot_path(task_name, f"{model_safe}_After_Tuning_feature_importance.png"),
            get_plot_path(task_name, f"{model_safe}_feature_importance.png")
        ]
        title = "Feature Importance"

    st.markdown(f"#### {title}")

    for path in possible_paths:
        if path.exists():
            st.image(str(path), width=SINGLE_IMAGE_WIDTH)
            return

    st.info(f"{title} not available for this model.")


def show_prediction_output(task_name, model_name):
    """
    Show prediction CSV output.
    """

    model_safe = safe_model_name(model_name)

    if is_linear_model(model_name):
        possible_paths = [
            get_csv_path(task_name, "predictions", f"{model_safe}_Before_Tuning_predictions.csv"),
            get_csv_path(task_name, "predictions", f"{model_safe}_predictions.csv")
        ]
    else:
        possible_paths = [
            get_csv_path(task_name, "predictions", f"{model_safe}_After_Tuning_predictions.csv")
        ]

    for path in possible_paths:
        if path.exists():
            pred_df = pd.read_csv(path)
            st.markdown("#### Prediction Output")
            st.caption(f"Showing first 100 rows from {path.name}")
            st.dataframe(pred_df.head(100), use_container_width=True, hide_index=True)
            return

    st.info("Prediction output not available.")


# =========================================================
# 11. SELECTED MODEL SECTION
# =========================================================

def show_selected_model_section(task_name, model_name, eval_df):
    """
    Show selected model diagnostic section.
    """

    st.divider()
    st.subheader(f"{model_name} Model Diagnostics")

    if is_linear_model(model_name):
        st.caption(
            "This is the baseline model, so no before-and-after tuning comparison is shown."
        )
    else:
        st.caption(
            "This section shows before and after hyperparameter tuning diagnostics."
        )

    show_metric_cards(eval_df, model_name)

    metric_df = create_selected_model_metric_table(
        eval_df=eval_df,
        model_name=model_name,
        task_name=task_name
    )

    show_table(metric_df, "Evaluation Metrics")

    if is_linear_model(model_name):

        col1, col2 = st.columns(2)

        with col1:
            show_baseline_image(
                task_name,
                model_name,
                "actual_vs_predicted",
                "Actual vs Predicted"
            )

        with col2:
            show_baseline_image(
                task_name,
                model_name,
                "residual_plot",
                "Residual Plot"
            )

        col3, col4 = st.columns(2)

        with col3:
            show_baseline_image(
                task_name,
                model_name,
                "residual_distribution",
                "Residual Distribution"
            )

        with col4:
            show_learning_curve(task_name, model_name)

        show_feature_importance(task_name, model_name)

    else:

        show_before_after_images(
            task_name,
            model_name,
            "actual_vs_predicted",
            "Actual vs Predicted: Before vs After Tuning"
        )

        show_before_after_images(
            task_name,
            model_name,
            "residual_plot",
            "Residual Plot: Before vs After Tuning"
        )

        show_before_after_images(
            task_name,
            model_name,
            "residual_distribution",
            "Residual Distribution: Before vs After Tuning"
        )

        col1, col2 = st.columns(2)

        with col1:
            show_learning_curve(task_name, model_name)

        with col2:
            show_feature_importance(task_name, model_name)

    with st.expander("Show Prediction Output"):
        show_prediction_output(task_name, model_name)


# =========================================================
# 12. TASK PAGE
# =========================================================

def show_task_page(task_name, task_title):
    """
    Show full model performance task page.
    """

    task_dir = MODEL_OUTPUT_DIR / task_name
    eval_path = task_dir / "evaluation_table.csv"

    if not eval_path.exists():
        st.warning(f"No evaluation table found for {task_title}.")
        st.code(str(eval_path))
        return

    eval_df = pd.read_csv(eval_path)

    st.subheader(f"{task_title} Model Details")

    model_list = eval_df["Model"].dropna().unique().tolist()
    model_options = ["All Model Comparison"] + model_list

    selected_model = st.selectbox(
        "Search or select model",
        model_options,
        key=f"{task_name}_model_filter"
    )

    if selected_model == "All Model Comparison":
        show_all_model_comparison(
            eval_df=eval_df,
            task_title=task_title
        )
    else:
        show_selected_model_section(
            task_name=task_name,
            model_name=selected_model,
            eval_df=eval_df
        )


# =========================================================
# 13. PAGE LAYOUT
# =========================================================

st.title("Model Performance")

st.write(
    "This page presents the model-level diagnostic outputs and model comparison "
    "for the Impression Prediction and CPM Prediction models."
)

tab1, tab2 = st.tabs(
    [
        "Impression Prediction",
        "CPM Prediction"
    ]
)

with tab1:
    show_task_page(
        task_name="impression",
        task_title="Impression Prediction"
    )

with tab2:
    show_task_page(
        task_name="cpm",
        task_title="CPM Prediction"
    )