#!/usr/bin/env python
"""Independently reload the saved checkpoints and verify reported accuracy."""

from __future__ import annotations

import json
from pathlib import Path

import joblib
import numpy as np
from sklearn.datasets import load_digits
from sklearn.metrics import accuracy_score, f1_score
from sklearn.model_selection import train_test_split


ROOT = Path(__file__).resolve().parents[1]


def json_safe(value):
    if isinstance(value, dict):
        return {key: json_safe(item) for key, item in value.items()}
    if isinstance(value, list):
        return [json_safe(item) for item in value]
    if isinstance(value, (float, np.floating)):
        return float(value) if np.isfinite(value) else None
    return value


def main() -> None:
    digits = load_digits()
    x = digits.data.astype(np.float64) / 16.0
    y = digits.target.astype(int)
    _, x_test, _, y_test = train_test_split(
        x, y, test_size=0.20, random_state=20261113, stratify=y
    )
    checkpoints = [
        ROOT / "checkpoints" / "baseline_tree.joblib",
        ROOT / "checkpoints" / "bagging_all_features.joblib",
        ROOT / "checkpoints" / "proposed_random_forest.joblib",
    ]
    report = {}
    for path in checkpoints:
        artifact = joblib.load(path)
        model = artifact["model"]
        prediction = model.predict(x_test)
        report[path.name] = {
            "accuracy": float(accuracy_score(y_test, prediction)),
            "macro_f1": float(f1_score(y_test, prediction, average="macro")),
            "stored_accuracy_matches": bool(
                np.isclose(
                    accuracy_score(y_test, prediction),
                    artifact["metrics"]["accuracy"],
                )
            ),
            "stored_metrics": artifact["metrics"],
        }
    with (ROOT / "results" / "checkpoint_verification.json").open(
        "w", encoding="utf-8"
    ) as handle:
        json.dump(json_safe(report), handle, ensure_ascii=False, indent=2, allow_nan=False)
    print(json.dumps(json_safe(report), ensure_ascii=False, indent=2, allow_nan=False))


if __name__ == "__main__":
    main()
