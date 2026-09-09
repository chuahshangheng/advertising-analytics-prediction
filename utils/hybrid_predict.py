import numpy as np
import pandas as pd


def prepare_catboost_input(sample_df: pd.DataFrame, model_meta: dict) -> pd.DataFrame:
    feature_columns = model_meta["feature_columns"]
    categorical_cols = model_meta["categorical_cols"]
    numeric_cols = model_meta["numeric_cols"]

    df = sample_df.copy()
    df = df.reindex(columns=feature_columns)

    for col in categorical_cols:
        if col not in df.columns:
            df[col] = "Missing"
        df[col] = df[col].fillna("Missing").astype(str)

    for col in numeric_cols:
        if col not in df.columns:
            df[col] = 0
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    return df


def decode_prediction(pred, target_mode: str):
    if target_mode in ["Log Impressions", "Log CPM"]:
        pred = np.expm1(pred)
    return max(float(pred), 0.0)


def is_in_catboost_range(sample_df: pd.DataFrame, cat_meta: dict) -> bool:
    cost_value = float(pd.to_numeric(sample_df["Cost"], errors="coerce").iloc[0])
    cost_min = float(cat_meta.get("cost_min", 1000))
    cost_max = float(cat_meta.get("cost_max", 4999))
    return cost_min <= cost_value <= cost_max


def hybrid_predict(sample_df: pd.DataFrame, loaded: dict) -> dict:
    cat_model = loaded["cat_model"]
    cat_meta = loaded["cat_meta"]

    cat_model_cpm = loaded["cat_model_cpm"]
    cat_model_cpm_meta = loaded["cat_model_cpm_meta"]

    cost_value = float(pd.to_numeric(sample_df["Cost"], errors="coerce").iloc[0])

    if is_in_catboost_range(sample_df, cat_meta):
        prepared_df = prepare_catboost_input(sample_df, cat_meta)

        pred = cat_model.predict(prepared_df)[0]
        pred = decode_prediction(pred, cat_meta.get("target_mode", "Impressions"))

        predicted_impressions = int(round(max(pred, 0)))

        if predicted_impressions > 0:
            predicted_cpm = round((cost_value / predicted_impressions) * 1000, 2)
        else:
            predicted_cpm = 0.0

        return {
            "predicted_impressions": predicted_impressions,
            "predicted_cpm": predicted_cpm,
            "model_used": "CatBoost",
            "range_status": "In training range",
            "warning": "",
        }

    else:
        # IMPORTANT:
        # remove Cost before CPM model prediction
        cpm_input = sample_df.drop(columns=["Cost"], errors="ignore")

        prepared_cpm_df = prepare_catboost_input(cpm_input, cat_model_cpm_meta)

        pred_cpm = cat_model_cpm.predict(prepared_cpm_df)[0]

        predicted_cpm = decode_prediction(
            pred_cpm,
            cat_model_cpm_meta.get("target_mode", "Raw CPM")
        )

        predicted_cpm = max(predicted_cpm, 0.01)
        predicted_cpm = round(predicted_cpm, 2)

        predicted_impressions = int(round((cost_value / predicted_cpm) * 1000))

        return {
            "predicted_impressions": predicted_impressions,
            "predicted_cpm": predicted_cpm,
            "model_used": "CatBoost CPM",
            "range_status": "Outside training range",
            "warning": "CPM model used because Cost is outside original CatBoost training range.",
        }



        