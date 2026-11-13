#!/usr/bin/env python
"""Regenerate figures from completed results and saved checkpoints."""

from __future__ import annotations

import joblib
import pandas as pd
from sklearn.datasets import load_digits
from sklearn.model_selection import train_test_split

from run_experiment import PROJECT_DIR, RESULT_DIR, save_figures


def main() -> None:
    digits = load_digits()
    x = digits.data.astype("float64") / 16.0
    y = digits.target.astype(int)
    _, x_test, _, y_test = train_test_split(
        x, y, test_size=0.20, random_state=20261113, stratify=y
    )
    checkpoints = {
        "baseline": PROJECT_DIR / "checkpoints" / "baseline_tree.joblib",
        "ablation_bagging": PROJECT_DIR
        / "checkpoints"
        / "bagging_all_features.joblib",
        "proposed": PROJECT_DIR / "checkpoints" / "proposed_random_forest.joblib",
    }
    models = {
        method: joblib.load(path)["model"] for method, path in checkpoints.items()
    }
    save_figures(
        digits=digits,
        metrics_df=pd.read_csv(RESULT_DIR / "per_seed_metrics.csv"),
        noise_df=pd.read_csv(RESULT_DIR / "noise_robustness_raw.csv"),
        occlusion_df=pd.read_csv(RESULT_DIR / "occlusion_robustness_raw.csv"),
        learning_df=pd.read_csv(RESULT_DIR / "learning_curves_raw.csv"),
        tree_curve_df=pd.read_csv(RESULT_DIR / "tree_count_curve.csv"),
        fitted_models=models,
        x_test=x_test,
        y_test=y_test,
    )
    print("Figures regenerated from completed results.")


if __name__ == "__main__":
    main()
