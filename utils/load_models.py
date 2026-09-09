from pathlib import Path
import json
from catboost import CatBoostRegressor

BASE_DIR = Path(__file__).resolve().parent.parent
MODEL_DIR = BASE_DIR / "models" / "models"

def load_all_models():
    # Original CatBoost model (in-range -> predicts impressions)
    cat_model = CatBoostRegressor()
    cat_model.load_model(str(MODEL_DIR / "catboost_model.cbm"))

    with open(MODEL_DIR / "catboost_meta.json", "r", encoding="utf-8") as f:
        cat_meta = json.load(f)

    # CPM CatBoost model (out-of-range -> predicts CPM)
    cat_model_cpm = CatBoostRegressor()
    cat_model_cpm.load_model(str(MODEL_DIR / "cat_model_cpm.cbm"))

    with open(MODEL_DIR / "cat_model_cpm_meta.json", "r", encoding="utf-8") as f:
        cat_model_cpm_meta = json.load(f)

    return {
        "cat_model": cat_model,
        "cat_meta": cat_meta,
        "cat_model_cpm": cat_model_cpm,
        "cat_model_cpm_meta": cat_model_cpm_meta,
    }