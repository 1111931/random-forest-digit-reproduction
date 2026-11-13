#!/usr/bin/env python
"""Run the random-forest handwritten-digit reproduction end to end."""

from __future__ import annotations

import argparse
import json
import logging
import platform
import sys
import time
from pathlib import Path
from typing import Any

import joblib
import matplotlib
import numpy as np
import pandas as pd
import scipy
import sklearn
from scipy.stats import binomtest
from sklearn.base import clone
from sklearn.datasets import load_digits
from sklearn.ensemble import BaggingClassifier, RandomForestClassifier
from sklearn.metrics import (
    accuracy_score,
    balanced_accuracy_score,
    confusion_matrix,
    f1_score,
    log_loss,
)
from sklearn.model_selection import train_test_split
from sklearn.tree import DecisionTreeClassifier

matplotlib.use("Agg")
import matplotlib.pyplot as plt

plt.rcParams["font.sans-serif"] = [
    "Microsoft YaHei",
    "SimHei",
    "Noto Sans CJK SC",
    "DejaVu Sans",
]
plt.rcParams["axes.unicode_minus"] = False


PROJECT_DIR = Path(__file__).resolve().parents[1]
CHECKPOINT_DIR = PROJECT_DIR / "checkpoints"
FIGURE_DIR = PROJECT_DIR / "figures"
RESULT_DIR = PROJECT_DIR / "results"

MAIN_SPLIT_SEED = 20261113
CHECKPOINT_SEED = 20261113
EVAL_SEEDS = [20261113, 20261114, 20261115, 20261116, 20261117]
N_ESTIMATORS = 200
LABELS = np.arange(10)


def configure_logging() -> None:
    RESULT_DIR.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(RESULT_DIR / "experiment.log", mode="w", encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
        force=True,
    )


def json_safe(value: Any) -> Any:
    if isinstance(value, dict):
        return {key: json_safe(item) for key, item in value.items()}
    if isinstance(value, list):
        return [json_safe(item) for item in value]
    if isinstance(value, (float, np.floating)):
        return float(value) if np.isfinite(value) else None
    return value


def make_models(seed: int) -> dict[str, Any]:
    baseline = DecisionTreeClassifier(
        criterion="gini",
        min_samples_leaf=2,
        random_state=seed,
    )
    ablation = BaggingClassifier(
        estimator=DecisionTreeClassifier(min_samples_leaf=2, random_state=seed),
        n_estimators=N_ESTIMATORS,
        max_samples=1.0,
        max_features=1.0,
        bootstrap=True,
        bootstrap_features=False,
        n_jobs=-1,
        random_state=seed,
    )
    proposed = RandomForestClassifier(
        n_estimators=N_ESTIMATORS,
        criterion="gini",
        max_features="sqrt",
        bootstrap=True,
        oob_score=True,
        class_weight="balanced_subsample",
        min_samples_leaf=1,
        n_jobs=-1,
        random_state=seed,
    )
    return {
        "baseline": baseline,
        "ablation_bagging": ablation,
        "proposed": proposed,
    }


def multiclass_brier(y_true: np.ndarray, proba: np.ndarray) -> float:
    one_hot = np.eye(len(LABELS))[y_true]
    return float(np.mean(np.sum((proba - one_hot) ** 2, axis=1)))


def expected_calibration_error(
    y_true: np.ndarray, proba: np.ndarray, n_bins: int = 10
) -> float:
    confidence = proba.max(axis=1)
    prediction = proba.argmax(axis=1)
    bins = np.linspace(0.0, 1.0, n_bins + 1)
    ece = 0.0
    for lower, upper in zip(bins[:-1], bins[1:]):
        mask = (confidence > lower) & (confidence <= upper)
        if not np.any(mask):
            continue
        accuracy = np.mean(prediction[mask] == y_true[mask])
        mean_confidence = np.mean(confidence[mask])
        ece += np.mean(mask) * abs(accuracy - mean_confidence)
    return float(ece)


def evaluate_model(model: Any, x_test: np.ndarray, y_test: np.ndarray) -> dict[str, Any]:
    start = time.perf_counter()
    prediction = model.predict(x_test)
    inference_seconds = time.perf_counter() - start
    proba = model.predict_proba(x_test)
    return {
        "accuracy": float(accuracy_score(y_test, prediction)),
        "balanced_accuracy": float(balanced_accuracy_score(y_test, prediction)),
        "macro_f1": float(f1_score(y_test, prediction, average="macro")),
        "weighted_f1": float(f1_score(y_test, prediction, average="weighted")),
        "log_loss": float(log_loss(y_test, proba, labels=LABELS)),
        "multiclass_brier": multiclass_brier(y_test, proba),
        "ece_10bin": expected_calibration_error(y_test, proba),
        "inference_ms_per_sample": float(1000.0 * inference_seconds / len(y_test)),
    }


def bootstrap_accuracy_ci(
    y_true: np.ndarray, y_pred: np.ndarray, seed: int, n_bootstrap: int = 2000
) -> dict[str, float]:
    rng = np.random.default_rng(seed)
    correct = (y_true == y_pred).astype(np.float64)
    values = np.empty(n_bootstrap)
    for index in range(n_bootstrap):
        sample = rng.integers(0, len(y_true), len(y_true))
        values[index] = correct[sample].mean()
    return {
        "low": float(np.quantile(values, 0.025)),
        "high": float(np.quantile(values, 0.975)),
    }


def mcnemar_exact(
    baseline_pred: np.ndarray,
    proposed_pred: np.ndarray,
    y_true: np.ndarray,
) -> dict[str, Any]:
    baseline_correct = baseline_pred == y_true
    proposed_correct = proposed_pred == y_true
    only_baseline = int(np.sum(baseline_correct & ~proposed_correct))
    only_proposed = int(np.sum(~baseline_correct & proposed_correct))
    discordant = only_baseline + only_proposed
    p_value = (
        float(binomtest(only_proposed, discordant, p=0.5).pvalue)
        if discordant > 0
        else 1.0
    )
    return {
        "baseline_only_correct": only_baseline,
        "proposed_only_correct": only_proposed,
        "discordant": discordant,
        "two_sided_exact_p": p_value,
    }


def measure_model_size(model: Any) -> int:
    import io

    stream = io.BytesIO()
    joblib.dump(model, stream, compress=3)
    return len(stream.getvalue())


def corruption_noise(
    x_test: np.ndarray, sigma: float, rng: np.random.Generator
) -> np.ndarray:
    noisy = x_test + rng.normal(0.0, sigma, size=x_test.shape)
    return np.clip(noisy, 0.0, 1.0)


def corruption_occlusion(
    x_test: np.ndarray, block_size: int
) -> np.ndarray:
    corrupted = x_test.copy()
    start = (8 - block_size) // 2
    stop = start + block_size
    grid = corrupted.reshape(-1, 8, 8)
    grid[:, start:stop, start:stop] = 0.0
    return corrupted


def run_noise_robustness(
    models: dict[str, Any], x_test: np.ndarray, y_test: np.ndarray
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    sigmas = [0.0, 0.05, 0.10, 0.15, 0.20, 0.25]
    for seed_index, seed in enumerate(EVAL_SEEDS):
        rng = np.random.default_rng(seed + 101)
        for sigma in sigmas:
            corrupted = corruption_noise(x_test, sigma, rng)
            for method, model in models[seed].items():
                score = accuracy_score(y_test, model.predict(corrupted))
                rows.append(
                    {
                        "seed": seed,
                        "seed_index": seed_index,
                        "sigma": sigma,
                        "method": method,
                        "accuracy": float(score),
                    }
                )
    return pd.DataFrame(rows)


def run_occlusion_robustness(
    models: dict[str, Any], x_test: np.ndarray, y_test: np.ndarray
) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    block_sizes = [0, 1, 2, 3, 4]
    for seed in EVAL_SEEDS:
        for block_size in block_sizes:
            corrupted = corruption_occlusion(x_test, block_size)
            for method, model in models[seed].items():
                score = accuracy_score(y_test, model.predict(corrupted))
                rows.append(
                    {
                        "seed": seed,
                        "block_size": block_size,
                        "method": method,
                        "accuracy": float(score),
                    }
                )
    return pd.DataFrame(rows)


def run_learning_curves(
    x_train: np.ndarray,
    y_train: np.ndarray,
    x_test: np.ndarray,
    y_test: np.ndarray,
) -> pd.DataFrame:
    fractions = [0.10, 0.20, 0.40, 0.60, 0.80, 1.00]
    rows: list[dict[str, Any]] = []
    for seed in [20261113, 20261114, 20261115]:
        rng = np.random.default_rng(seed + 202)
        for fraction in fractions:
            if fraction == 1.0:
                indices = np.arange(len(y_train))
            else:
                n_per_class = max(2, int(np.floor(fraction * 145)))
                indices_list: list[int] = []
                for label in LABELS:
                    candidates = np.flatnonzero(y_train == label)
                    chosen = rng.choice(candidates, size=n_per_class, replace=False)
                    indices_list.extend(chosen.tolist())
                indices = np.asarray(indices_list)
            models = make_models(seed)
            for method in ["baseline", "proposed"]:
                model = models[method]
                model.fit(x_train[indices], y_train[indices])
                rows.append(
                    {
                        "seed": seed,
                        "train_fraction": fraction,
                        "train_samples": int(len(indices)),
                        "method": method,
                        "accuracy": float(
                            accuracy_score(y_test, model.predict(x_test))
                        ),
                    }
                )
            logging.info(
                "Learning curve seed=%s fraction=%.2f n=%d",
                seed,
                fraction,
                len(indices),
            )
    return pd.DataFrame(rows)


def staged_ensemble_curve(
    fitted_models: dict[str, Any], x_test: np.ndarray, y_test: np.ndarray
) -> pd.DataFrame:
    tree_counts = [1, 5, 10, 25, 50, 100, 200]
    rows: list[dict[str, Any]] = []
    for method in ["ablation_bagging", "proposed"]:
        model = fitted_models[method]
        estimators = np.asarray(model.estimators_)
        for count in tree_counts:
            selected = estimators[:count]
            probability = np.mean(
                [estimator.predict_proba(x_test) for estimator in selected], axis=0
            )
            prediction = probability.argmax(axis=1)
            rows.append(
                {
                    "method": method,
                    "tree_count": count,
                    "accuracy": float(accuracy_score(y_test, prediction)),
                }
            )
    return pd.DataFrame(rows)


def save_figures(
    digits: Any,
    metrics_df: pd.DataFrame,
    noise_df: pd.DataFrame,
    occlusion_df: pd.DataFrame,
    learning_df: pd.DataFrame,
    tree_curve_df: pd.DataFrame,
    fitted_models: dict[str, Any],
    x_test: np.ndarray,
    y_test: np.ndarray,
) -> None:
    FIGURE_DIR.mkdir(parents=True, exist_ok=True)
    colors = {
        "baseline": "#7B8794",
        "ablation_bagging": "#D8942F",
        "proposed": "#176B87",
    }
    labels_cn = {
        "baseline": "单棵决策树",
        "ablation_bagging": "全特征 Bagging",
        "proposed": "随机森林（本文）",
    }

    # Figure 1: framework diagram.
    fig, ax = plt.subplots(figsize=(11.0, 4.5))
    ax.set_xlim(0, 11)
    ax.set_ylim(0, 4.5)
    ax.axis("off")
    boxes = [
        (0.25, 1.55, 1.7, 1.3, "8×8\n手写数字", "#E8EEF5"),
        (2.35, 1.55, 1.9, 1.3, "像素向量化\n64 维", "#D9E8F2"),
        (4.65, 2.20, 2.0, 1.25, "Bootstrap\n重采样", "#F5EAD8"),
        (4.65, 0.65, 2.0, 1.05, "随机像素\n子空间", "#E7F0E8"),
        (7.15, 1.55, 1.7, 1.3, "200 棵\n决策树", "#DCEAF4"),
        (9.20, 1.55, 1.45, 1.3, "概率平均\n分类", "#176B87"),
    ]
    for x, y, width, height, text, fill in boxes:
        face = fill if fill != "#176B87" else "#176B87"
        text_color = "white" if fill == "#176B87" else "#172B4D"
        rect = plt.Rectangle(
            (x, y), width, height, facecolor=face, edgecolor="#C5D0DC", linewidth=1.2
        )
        ax.add_patch(rect)
        ax.text(
            x + width / 2,
            y + height / 2,
            text,
            ha="center",
            va="center",
            fontsize=11,
            color=text_color,
            fontweight="bold",
        )
    arrows = [
        ((1.95, 2.20), (2.35, 2.20)),
        ((4.25, 2.20), (4.65, 2.82)),
        ((4.25, 2.20), (4.65, 1.18)),
        ((6.65, 2.82), (7.15, 2.20)),
        ((6.65, 1.18), (7.15, 2.20)),
        ((8.85, 2.20), (9.20, 2.20)),
    ]
    for start, end in arrows:
        ax.annotate(
            "",
            xy=end,
            xytext=start,
            arrowprops=dict(arrowstyle="->", color="#60758A", lw=1.6),
        )
    ax.text(
        5.5,
        0.18,
        "集成多样性：样本扰动 + 特征扰动",
        ha="center",
        fontsize=10,
        color="#52616B",
    )
    fig.tight_layout()
    fig.savefig(FIGURE_DIR / "fig01_framework.png", dpi=220, bbox_inches="tight")
    plt.close(fig)

    # Figure 2: dataset examples and class distribution.
    fig, axes = plt.subplots(2, 10, figsize=(11.0, 3.2))
    for label in LABELS:
        index = int(np.flatnonzero(digits.target == label)[0])
        ax = axes[0, label]
        ax.imshow(digits.images[index], cmap="gray_r", interpolation="nearest")
        ax.set_title(str(label), fontsize=10)
        ax.axis("off")
        ax = axes[1, label]
        index = int(np.flatnonzero(digits.target == label)[12])
        ax.imshow(digits.images[index], cmap="gray_r", interpolation="nearest")
        ax.axis("off")
    fig.suptitle("数据集样例：每列对应一个数字类别", fontsize=12, y=1.02)
    fig.tight_layout()
    fig.savefig(FIGURE_DIR / "fig02_digit_samples.png", dpi=220, bbox_inches="tight")
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(7.2, 3.8))
    counts = np.bincount(digits.target, minlength=10)
    bars = ax.bar(LABELS, counts, color="#6FA8C4", edgecolor="white")
    ax.set_xlabel("数字类别")
    ax.set_ylabel("样本数")
    ax.set_xticks(LABELS)
    ax.set_ylim(0, 205)
    ax.spines[["top", "right"]].set_visible(False)
    for bar, count in zip(bars, counts):
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            count + 3,
            str(count),
            ha="center",
            fontsize=8.5,
        )
    fig.tight_layout()
    fig.savefig(FIGURE_DIR / "fig03_class_distribution.png", dpi=220, bbox_inches="tight")
    plt.close(fig)

    # Figure 4: main metrics.
    summary = (
        metrics_df.groupby("method")
        .agg(
            accuracy_mean=("accuracy", "mean"),
            accuracy_std=("accuracy", "std"),
            macro_f1_mean=("macro_f1", "mean"),
            macro_f1_std=("macro_f1", "std"),
        )
        .reset_index()
    )
    method_order = ["baseline", "ablation_bagging", "proposed"]
    summary = summary.set_index("method").loc[method_order].reset_index()
    x_pos = np.arange(len(summary))
    width = 0.34
    fig, ax = plt.subplots(figsize=(8.0, 4.2))
    ax.bar(
        x_pos - width / 2,
        summary["accuracy_mean"],
        width,
        yerr=summary["accuracy_std"],
        color=[colors[m] for m in summary["method"]],
        capsize=4,
        label="准确率",
    )
    ax.bar(
        x_pos + width / 2,
        summary["macro_f1_mean"],
        width,
        yerr=summary["macro_f1_std"],
        color=[colors[m] for m in summary["method"]],
        alpha=0.58,
        capsize=4,
        hatch="//",
        label="宏平均 F1",
    )
    ax.set_xticks(x_pos, [labels_cn[m] for m in summary["method"]])
    ax.set_ylim(0.70, 1.01)
    ax.set_ylabel("得分")
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(frameon=False, ncol=2, loc="upper left")
    for index, row in summary.iterrows():
        ax.text(
            index - width / 2,
            row["accuracy_mean"] + 0.008,
            f'{row["accuracy_mean"]:.3f}',
            ha="center",
            fontsize=8.5,
        )
        ax.text(
            index + width / 2,
            row["macro_f1_mean"] + 0.008,
            f'{row["macro_f1_mean"]:.3f}',
            ha="center",
            fontsize=8.5,
        )
    fig.tight_layout()
    fig.savefig(FIGURE_DIR / "fig04_main_results.png", dpi=220, bbox_inches="tight")
    plt.close(fig)

    # Figure 5: robustness.
    fig, axes = plt.subplots(1, 2, figsize=(11.0, 4.1))
    for method in method_order:
        subset = noise_df[noise_df["method"] == method]
        stats = subset.groupby("sigma")["accuracy"].agg(["mean", "std"]).reset_index()
        axes[0].errorbar(
            stats["sigma"],
            stats["mean"],
            yerr=stats["std"],
            marker="o",
            linewidth=2,
            capsize=3,
            label=labels_cn[method],
            color=colors[method],
        )
        subset = occlusion_df[occlusion_df["method"] == method]
        stats = subset.groupby("block_size")["accuracy"].agg(["mean", "std"]).reset_index()
        axes[1].errorbar(
            stats["block_size"],
            stats["mean"],
            yerr=stats["std"],
            marker="o",
            linewidth=2,
            capsize=3,
            label=labels_cn[method],
            color=colors[method],
        )
    axes[0].set_xlabel(r"高斯噪声标准差 $\sigma$")
    axes[0].set_ylabel("测试准确率")
    axes[0].set_title("像素噪声鲁棒性")
    axes[1].set_xlabel("中心遮挡块边长")
    axes[1].set_ylabel("测试准确率")
    axes[1].set_title("局部遮挡鲁棒性")
    for ax in axes:
        ax.grid(alpha=0.18)
        ax.spines[["top", "right"]].set_visible(False)
    handles, labels = axes[0].get_legend_handles_labels()
    fig.legend(handles, labels, loc="lower center", ncol=3, frameon=False)
    fig.tight_layout(rect=(0, 0.08, 1, 1))
    fig.savefig(FIGURE_DIR / "fig05_robustness.png", dpi=220, bbox_inches="tight")
    plt.close(fig)

    # Figure 6: learning curve.
    fig, ax = plt.subplots(figsize=(7.4, 4.1))
    for method in ["baseline", "proposed"]:
        subset = learning_df[learning_df["method"] == method]
        stats = (
            subset.groupby("train_samples")["accuracy"]
            .agg(["mean", "std"])
            .reset_index()
        )
        ax.errorbar(
            stats["train_samples"],
            stats["mean"],
            yerr=stats["std"],
            marker="o",
            linewidth=2,
            capsize=3,
            label=labels_cn[method],
            color=colors[method],
        )
    ax.set_xlabel("训练样本数")
    ax.set_ylabel("测试准确率")
    ax.grid(alpha=0.18)
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(frameon=False)
    fig.tight_layout()
    fig.savefig(FIGURE_DIR / "fig06_learning_curve.png", dpi=220, bbox_inches="tight")
    plt.close(fig)

    # Figure 7: confusion matrix.
    prediction = fitted_models["proposed"].predict(x_test)
    cm = confusion_matrix(y_test, prediction, labels=LABELS)
    fig, ax = plt.subplots(figsize=(6.2, 5.2))
    image = ax.imshow(cm, cmap="Blues", vmin=0, vmax=cm.max())
    for row in range(10):
        for col in range(10):
            color = "white" if cm[row, col] > cm.max() * 0.55 else "#172B4D"
            ax.text(
                col,
                row,
                str(cm[row, col]),
                ha="center",
                va="center",
                fontsize=8.5,
                color=color,
            )
    ax.set_xticks(LABELS)
    ax.set_yticks(LABELS)
    ax.set_xlabel("预测类别")
    ax.set_ylabel("真实类别")
    fig.colorbar(image, ax=ax, fraction=0.046, pad=0.04)
    fig.tight_layout()
    fig.savefig(FIGURE_DIR / "fig07_confusion_matrix.png", dpi=220, bbox_inches="tight")
    plt.close(fig)

    # Figure 8: feature importance and class F1.
    importance = fitted_models["proposed"].feature_importances_.reshape(8, 8)
    fig, axes = plt.subplots(1, 2, figsize=(10.6, 4.0))
    image = axes[0].imshow(importance, cmap="YlGnBu")
    axes[0].set_title("随机森林像素重要性")
    axes[0].set_xticks([])
    axes[0].set_yticks([])
    fig.colorbar(image, ax=axes[0], fraction=0.046, pad=0.04)

    f1_baseline = f1_score(y_test, fitted_models["baseline"].predict(x_test), average=None, labels=LABELS)
    f1_proposed = f1_score(y_test, fitted_models["proposed"].predict(x_test), average=None, labels=LABELS)
    x_pos = np.arange(10)
    axes[1].bar(
        x_pos - 0.19,
        f1_baseline,
        0.38,
        color=colors["baseline"],
        label="单棵决策树",
    )
    axes[1].bar(
        x_pos + 0.19,
        f1_proposed,
        0.38,
        color=colors["proposed"],
        label="随机森林",
    )
    axes[1].set_xticks(x_pos)
    axes[1].set_xlabel("数字类别")
    axes[1].set_ylabel("F1")
    axes[1].set_ylim(0.60, 1.02)
    axes[1].set_title("逐类 F1 对照")
    axes[1].spines[["top", "right"]].set_visible(False)
    axes[1].legend(frameon=False, fontsize=8.5, loc="lower right")
    fig.tight_layout()
    fig.savefig(FIGURE_DIR / "fig08_importance_per_class.png", dpi=220, bbox_inches="tight")
    plt.close(fig)

    # Figure 9: tree-count curve.
    fig, ax = plt.subplots(figsize=(7.2, 4.0))
    for method in ["ablation_bagging", "proposed"]:
        subset = tree_curve_df[tree_curve_df["method"] == method]
        ax.plot(
            subset["tree_count"],
            subset["accuracy"],
            marker="o",
            linewidth=2,
            label=labels_cn[method],
            color=colors[method],
        )
    ax.set_xscale("log")
    ax.set_xticks([1, 5, 10, 25, 50, 100, 200], ["1", "5", "10", "25", "50", "100", "200"])
    ax.set_xlabel("集成中的树数量")
    ax.set_ylabel("测试准确率")
    ax.grid(alpha=0.18)
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(frameon=False)
    fig.tight_layout()
    fig.savefig(FIGURE_DIR / "fig09_tree_count.png", dpi=220, bbox_inches="tight")
    plt.close(fig)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--quick",
        action="store_true",
        help="Run a reduced smoke test (not suitable for paper results).",
    )
    args = parser.parse_args()
    configure_logging()
    for directory in [CHECKPOINT_DIR, FIGURE_DIR, RESULT_DIR]:
        directory.mkdir(parents=True, exist_ok=True)

    global EVAL_SEEDS, N_ESTIMATORS
    if args.quick:
        EVAL_SEEDS = EVAL_SEEDS[:2]
        N_ESTIMATORS = 30
        logging.warning("QUICK MODE: numbers are for smoke testing only.")

    digits = load_digits()
    x = digits.data.astype(np.float64) / 16.0
    y = digits.target.astype(int)
    x_train, x_test, y_train, y_test = train_test_split(
        x,
        y,
        test_size=0.20,
        random_state=MAIN_SPLIT_SEED,
        stratify=y,
    )
    logging.info("Dataset: x=%s, train=%s, test=%s", x.shape, x_train.shape, x_test.shape)
    logging.info("Train counts: %s", np.bincount(y_train, minlength=10).tolist())
    logging.info("Test counts: %s", np.bincount(y_test, minlength=10).tolist())

    all_models: dict[int, dict[str, Any]] = {}
    metric_rows: list[dict[str, Any]] = []
    fit_seconds: dict[tuple[int, str], float] = {}
    for seed in EVAL_SEEDS:
        models = make_models(seed)
        all_models[seed] = models
        for method, model in models.items():
            start = time.perf_counter()
            model.fit(x_train, y_train)
            elapsed = time.perf_counter() - start
            fit_seconds[(seed, method)] = elapsed
            metrics = evaluate_model(model, x_test, y_test)
            row = {
                "seed": seed,
                "method": method,
                "fit_seconds": elapsed,
                "model_bytes_joblib": measure_model_size(model),
                "oob_score": float(getattr(model, "oob_score_", np.nan)),
                **metrics,
            }
            metric_rows.append(row)
            logging.info(
                "seed=%s method=%-18s acc=%.4f macroF1=%.4f fit=%.3fs",
                seed,
                method,
                metrics["accuracy"],
                metrics["macro_f1"],
                elapsed,
            )
            if method == "proposed":
                logging.info(
                    "seed=%s method=%-18s oob=%.4f",
                    seed,
                    method,
                    model.oob_score_,
                )

    metrics_df = pd.DataFrame(metric_rows)
    metrics_df.to_csv(RESULT_DIR / "per_seed_metrics.csv", index=False)
    summary = (
        metrics_df.groupby("method")
        .agg(
            accuracy_mean=("accuracy", "mean"),
            accuracy_std=("accuracy", "std"),
            balanced_accuracy_mean=("balanced_accuracy", "mean"),
            macro_f1_mean=("macro_f1", "mean"),
            macro_f1_std=("macro_f1", "std"),
            log_loss_mean=("log_loss", "mean"),
            brier_mean=("multiclass_brier", "mean"),
            ece_mean=("ece_10bin", "mean"),
            oob_mean=("oob_score", "mean"),
            oob_std=("oob_score", "std"),
            fit_seconds_mean=("fit_seconds", "mean"),
            fit_seconds_median=("fit_seconds", "median"),
            inference_ms_mean=("inference_ms_per_sample", "mean"),
            inference_ms_median=("inference_ms_per_sample", "median"),
            model_bytes_mean=("model_bytes_joblib", "mean"),
        )
        .reset_index()
    )
    summary.to_csv(RESULT_DIR / "main_results.csv", index=False)

    checkpoint_models = all_models[CHECKPOINT_SEED]
    prediction_map: dict[str, np.ndarray] = {}
    probability_map: dict[str, np.ndarray] = {}
    checkpoint_metrics: dict[str, Any] = {}
    for method, model in checkpoint_models.items():
        prediction_map[method] = model.predict(x_test)
        probability_map[method] = model.predict_proba(x_test)
        checkpoint_metrics[method] = evaluate_model(model, x_test, y_test)
        checkpoint_metrics[method].update(
            {
                "save_accuracy": float(accuracy_score(y_test, prediction_map[method])),
                "n_estimators": int(N_ESTIMATORS) if hasattr(model, "estimators_") else 1,
                "model_bytes_joblib": int(
                    next(
                        row["model_bytes_joblib"]
                        for row in metric_rows
                        if row["seed"] == CHECKPOINT_SEED and row["method"] == method
                    )
                ),
                "oob_score": float(getattr(model, "oob_score_", np.nan)),
            }
        )
        artifact = {
            "model": model,
            "method": method,
            "input": "8x8 grayscale; flatten in C order; divide values by 16",
            "input_shape": [8, 8],
            "feature_dimension": 64,
            "split_seed": MAIN_SPLIT_SEED,
            "model_seed": CHECKPOINT_SEED,
            "dataset": "sklearn.datasets.load_digits",
            "training_samples": int(len(y_train)),
            "test_samples": int(len(y_test)),
            "metrics": checkpoint_metrics[method],
            "sklearn_version": sklearn.__version__,
            "numpy_version": np.__version__,
        }
        filename = {
            "baseline": "baseline_tree.joblib",
            "ablation_bagging": "bagging_all_features.joblib",
            "proposed": "proposed_random_forest.joblib",
        }[method]
        joblib.dump(artifact, CHECKPOINT_DIR / filename, compress=3)

    np.savez_compressed(
        RESULT_DIR / "predictions.npz",
        y_test=y_test,
        **{f"pred_{key}": value for key, value in prediction_map.items()},
        **{f"proba_{key}": value for key, value in probability_map.items()},
    )

    ci = {
        method: bootstrap_accuracy_ci(
            y_test, prediction_map[method], MAIN_SPLIT_SEED + index
        )
        for index, method in enumerate(["baseline", "ablation_bagging", "proposed"])
    }
    mcnemar_baseline = mcnemar_exact(
        prediction_map["baseline"], prediction_map["proposed"], y_test
    )
    mcnemar_bagging = mcnemar_exact(
        prediction_map["ablation_bagging"], prediction_map["proposed"], y_test
    )
    per_class = {}
    for method in ["baseline", "ablation_bagging", "proposed"]:
        f1_values = f1_score(
            y_test, prediction_map[method], average=None, labels=LABELS
        )
        per_class[method] = {
            str(int(label)): float(value)
            for label, value in zip(LABELS, f1_values)
        }

    noise_df = run_noise_robustness(all_models, x_test, y_test)
    noise_df.to_csv(RESULT_DIR / "noise_robustness_raw.csv", index=False)
    noise_summary = (
        noise_df.groupby(["method", "sigma"])["accuracy"]
        .agg(["mean", "std"])
        .reset_index()
    )
    noise_summary.to_csv(RESULT_DIR / "noise_robustness.csv", index=False)

    occlusion_df = run_occlusion_robustness(all_models, x_test, y_test)
    occlusion_df.to_csv(RESULT_DIR / "occlusion_robustness_raw.csv", index=False)
    occlusion_summary = (
        occlusion_df.groupby(["method", "block_size"])["accuracy"]
        .agg(["mean", "std"])
        .reset_index()
    )
    occlusion_summary.to_csv(RESULT_DIR / "occlusion_robustness.csv", index=False)

    learning_df = run_learning_curves(x_train, y_train, x_test, y_test)
    learning_df.to_csv(RESULT_DIR / "learning_curves_raw.csv", index=False)
    learning_summary = (
        learning_df.groupby(["method", "train_samples"])["accuracy"]
        .agg(["mean", "std"])
        .reset_index()
    )
    learning_summary.to_csv(RESULT_DIR / "learning_curves.csv", index=False)

    tree_curve_df = staged_ensemble_curve(
        checkpoint_models, x_test, y_test
    )
    tree_curve_df.to_csv(RESULT_DIR / "tree_count_curve.csv", index=False)

    generate_figures_flag = not args.quick
    if generate_figures_flag:
        save_figures(
            digits,
            metrics_df,
            noise_df,
            occlusion_df,
            learning_df,
            tree_curve_df,
            checkpoint_models,
            x_test,
            y_test,
        )

    train_counts = np.bincount(y_train, minlength=10).tolist()
    test_counts = np.bincount(y_test, minlength=10).tolist()
    output = {
        "experiment": "Random Forest for Handwritten Digit Recognition",
        "is_quick_mode": bool(args.quick),
        "dataset": {
            "name": "sklearn.datasets.load_digits",
            "n_samples": int(len(y)),
            "n_features": int(x.shape[1]),
            "class_count": 10,
            "train_samples": int(len(y_train)),
            "test_samples": int(len(y_test)),
            "train_counts": train_counts,
            "test_counts": test_counts,
            "split_seed": MAIN_SPLIT_SEED,
        },
        "configuration": {
            "eval_seeds": EVAL_SEEDS,
            "checkpoint_seed": CHECKPOINT_SEED,
            "n_estimators": N_ESTIMATORS,
            "max_features": "sqrt",
            "class_weight": "balanced_subsample",
            "oob_score_enabled_proposed": True,
        },
        "environment": {
            "python": platform.python_version(),
            "platform": platform.platform(),
            "numpy": np.__version__,
            "scipy": scipy.__version__,
            "sklearn": sklearn.__version__,
            "matplotlib": matplotlib.__version__,
            "pandas": pd.__version__,
        },
        "main_summary": summary.to_dict(orient="records"),
        "checkpoint_metrics": checkpoint_metrics,
        "bootstrap_95_ci": ci,
        "mcnemar": {
            "baseline_vs_proposed": mcnemar_baseline,
            "bagging_vs_proposed": mcnemar_bagging,
        },
        "per_class_f1": per_class,
        "oob_per_seed": (
            metrics_df.loc[metrics_df["method"] == "proposed", ["seed", "oob_score"]]
            .to_dict(orient="records")
        ),
        "noise_summary": noise_summary.to_dict(orient="records"),
        "occlusion_summary": occlusion_summary.to_dict(orient="records"),
        "tree_count_curve": tree_curve_df.to_dict(orient="records"),
        "artifacts": {
            "baseline_checkpoint": "checkpoints/baseline_tree.joblib",
            "bagging_checkpoint": "checkpoints/bagging_all_features.joblib",
            "proposed_checkpoint": "checkpoints/proposed_random_forest.joblib",
            "predictions": "results/predictions.npz",
        },
    }
    with (RESULT_DIR / "metrics.json").open("w", encoding="utf-8") as handle:
        json.dump(json_safe(output), handle, ensure_ascii=False, indent=2, allow_nan=False)

    logging.info("Saved checkpoints to %s", CHECKPOINT_DIR)
    logging.info("Saved results to %s", RESULT_DIR)
    logging.info("Saved figures to %s", FIGURE_DIR)


if __name__ == "__main__":
    main()
