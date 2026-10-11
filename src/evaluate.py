"""
evaluate.py - Đánh giá mô hình & phân tích lỗi

Chức năng:
  1. Đánh giá trên tập validation (chi tiết)
  2. Confusion matrix
  3. ROC Curve so sánh
  4. Phân tích lỗi chi tiết (TN4)
  5. Đánh giá cuối trên Test set (CHỈ 1 LẦN!)
  6. Lưu final model cho web app
"""

import os
import sys

# Khắc phục lỗi font Tiếng Việt / Unicode trên Windows Console
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import numpy as np
import pandas as pd
import joblib
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns

from sklearn.metrics import (
    accuracy_score, f1_score, roc_auc_score,
    classification_report, confusion_matrix,
    ConfusionMatrixDisplay, RocCurveDisplay,
)

# ============================================================
# CẤU HÌNH
# ============================================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROCESSED_DIR = os.path.join(BASE_DIR, "data", "processed")
MODELS_DIR = os.path.join(BASE_DIR, "models")
FIGURES_DIR = os.path.join(BASE_DIR, "report", "figures")

FEATURE_COLS = [
    "Area", "Perimeter", "Major_Axis_Length", "Minor_Axis_Length",
    "Eccentricity", "Convex_Area", "Extent"
]
TARGET_COL = "Class"


def load_data_and_models():
    """Đọc dữ liệu val, test và các model đã lưu."""
    print("[0] Đọc dữ liệu và model...")

    val_df = pd.read_csv(os.path.join(PROCESSED_DIR, "val.csv"))
    test_df = pd.read_csv(os.path.join(PROCESSED_DIR, "test.csv"))

    X_val, y_val = val_df[FEATURE_COLS], val_df[TARGET_COL]
    X_test, y_test = test_df[FEATURE_COLS], test_df[TARGET_COL]

    models = {}
    for fname in ["decision_tree_pruned.pkl", "random_forest_100_trees.pkl"]:
        path = os.path.join(MODELS_DIR, fname)
        if os.path.exists(path):
            name = fname.replace(".pkl", "").replace("_", " ").title()
            models[name] = joblib.load(path)
            print(f"  ✓ Loaded: {fname}")

    return X_val, y_val, X_test, y_test, models


def plot_confusion_matrices(models, X_val, y_val):
    """Vẽ confusion matrix cho tất cả model."""
    print("\n[1] Confusion matrices")

    n = len(models)
    fig, axes = plt.subplots(1, n, figsize=(7 * n, 5))
    if n == 1:
        axes = [axes]

    for ax, (name, model) in zip(axes, models.items()):
        ConfusionMatrixDisplay.from_estimator(
            model, X_val, y_val, ax=ax,
            display_labels=model.classes_,
            cmap="Blues",
        )
        ax.set_title(name, fontsize=12, fontweight="bold")

    plt.suptitle("Confusion Matrix trên tập Validation", fontsize=14, fontweight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, "10_confusion_matrices.png"), dpi=150)
    plt.close()
    print("  ✓ Lưu: 10_confusion_matrices.png")


def plot_roc_curves(models, X_val, y_val):
    """Vẽ ROC curve so sánh các model."""
    print("\n[2] ROC Curves")

    fig, ax = plt.subplots(figsize=(8, 6))
    colors = ["steelblue", "forestgreen", "tomato"]

    for (name, model), color in zip(models.items(), colors):
        RocCurveDisplay.from_estimator(
            model, X_val, y_val,
            ax=ax, name=name,
            plot_chance_level=False,
        )

    ax.plot([0, 1], [0, 1], "k--", alpha=0.5, label="Random (AUC=0.5)")
    ax.set_title("ROC Curves — So sánh mô hình", fontsize=14, fontweight="bold")
    ax.legend(fontsize=9)
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, "11_roc_curves.png"), dpi=150)
    plt.close()
    print("  ✓ Lưu: 11_roc_curves.png")


def error_analysis(models, X_val, y_val):
    """TN4: Phân tích vùng đặc trưng của các lỗi."""
    print("\n[3] THÍ NGHIỆM 4: Phân tích lỗi chi tiết")

    # Dùng model tốt nhất (Random Forest)
    model_name = list(models.keys())[-1]  # RF thường là model cuối
    model = models[model_name]
    y_pred = model.predict(X_val)

    # Tách đúng vs sai
    mask_correct = (y_pred == y_val.values)
    mask_wrong = ~mask_correct

    n_errors = mask_wrong.sum()
    n_total = len(y_val)
    print(f"\n  Mô hình: {model_name}")
    print(f"  Tổng lỗi: {n_errors}/{n_total} ({n_errors/n_total*100:.1f}%)")

    # Chi tiết lỗi
    errors_df = X_val[mask_wrong].copy()
    errors_df["True_Label"] = y_val.values[mask_wrong]
    errors_df["Predicted"] = y_pred[mask_wrong]

    # Loại lỗi
    cammeo_as_osmancik = errors_df[
        (errors_df["True_Label"] == "Cammeo") & (errors_df["Predicted"] == "Osmancik")
    ]
    osmancik_as_cammeo = errors_df[
        (errors_df["True_Label"] == "Osmancik") & (errors_df["Predicted"] == "Cammeo")
    ]
    print(f"  - Cammeo → Osmancik (sai): {len(cammeo_as_osmancik)}")
    print(f"  - Osmancik → Cammeo (sai): {len(osmancik_as_cammeo)}")

    # So sánh thống kê: mẫu đúng vs mẫu sai
    print(f"\n  📊 So sánh trung bình đặc trưng (Đúng vs Sai):")
    print(f"  {'Feature':<25s} {'Đúng':>10s} {'Sai':>10s} {'Δ':>10s}")
    print(f"  {'-'*25} {'-'*10} {'-'*10} {'-'*10}")
    for col in FEATURE_COLS:
        correct_mean = X_val[mask_correct][col].mean()
        wrong_mean = X_val[mask_wrong][col].mean()
        delta = wrong_mean - correct_mean
        print(f"  {col:<25s} {correct_mean:10.2f} {wrong_mean:10.2f} {delta:+10.2f}")

    # Vẽ biểu đồ phân tích lỗi
    fig, axes = plt.subplots(2, 4, figsize=(18, 9))
    axes = axes.flatten()
    for i, col in enumerate(FEATURE_COLS):
        ax = axes[i]
        ax.hist(X_val[mask_correct][col], bins=25, alpha=0.6, label="Đúng", color="steelblue")
        ax.hist(X_val[mask_wrong][col], bins=25, alpha=0.6, label="Sai", color="tomato")
        ax.set_title(col, fontsize=11, fontweight="bold")
        ax.legend(fontsize=8)
    axes[-1].axis("off")
    fig.suptitle(f"Phân phối đặc trưng: Dự đoán Đúng vs Sai ({model_name})",
                 fontsize=14, fontweight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, "12_error_analysis.png"), dpi=150)
    plt.close()
    print("  ✓ Lưu: 12_error_analysis.png")

    # Xác suất của các mẫu bị lỗi
    y_proba = model.predict_proba(X_val)
    max_proba = y_proba.max(axis=1)
    avg_conf_correct = max_proba[mask_correct].mean()
    avg_conf_wrong = max_proba[mask_wrong].mean()
    print(f"\n  Xác suất trung bình:")
    print(f"  - Mẫu đúng: {avg_conf_correct:.4f}")
    print(f"  - Mẫu sai:  {avg_conf_wrong:.4f}")
    print(f"  → Mẫu bị sai có xác suất thấp hơn → mô hình 'không chắc chắn' ở các mẫu khó.")

    return errors_df


def final_test_evaluation(models, X_test, y_test):
    """
    Đánh giá CUỐI CÙNG trên tập Test.

    ⚠ CHỈ CHẠY 1 LẦN! Không quay lại thay đổi mô hình sau bước này!
    """
    print("\n" + "=" * 60)
    print("  ⚠  ĐÁNH GIÁ CUỐI CÙNG TRÊN TẬP TEST (CHỈ 1 LẦN!)")
    print("=" * 60)

    final_results = {}

    for name, model in models.items():
        y_pred = model.predict(X_test)
        y_proba = model.predict_proba(X_test)
        classes = list(model.classes_)
        # Dùng Osmancik (lớp thứ 2 theo alphabet) làm positive cho ROC-AUC
        osmancik_idx = classes.index("Osmancik")

        acc = accuracy_score(y_test, y_pred)
        f1 = f1_score(y_test, y_pred, pos_label="Cammeo")
        auc = roc_auc_score(y_test, y_proba[:, osmancik_idx])

        final_results[name] = {
            "accuracy": round(acc, 4),
            "f1": round(f1, 4),
            "roc_auc": round(auc, 4),
        }

        print(f"\n  === {name} ===")
        print(f"  Accuracy: {acc:.4f}")
        print(f"  F1:       {f1:.4f}")
        print(f"  ROC-AUC:  {auc:.4f}")
        print(classification_report(y_test, y_pred))

    # Confusion matrix trên test
    n = len(models)
    fig, axes = plt.subplots(1, n, figsize=(7 * n, 5))
    if n == 1:
        axes = [axes]
    for ax, (name, model) in zip(axes, models.items()):
        ConfusionMatrixDisplay.from_estimator(
            model, X_test, y_test, ax=ax,
            display_labels=model.classes_, cmap="Greens",
        )
        ax.set_title(f"{name}\n(TEST SET)", fontsize=12, fontweight="bold")
    plt.suptitle("Confusion Matrix trên tập TEST (chỉ 1 lần)", fontsize=14, fontweight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, "13_test_confusion_matrices.png"), dpi=150)
    plt.close()
    print("  ✓ Lưu: 13_test_confusion_matrices.png")

    return final_results


def save_final_model(models):
    """Lưu model cuối cùng cho web app."""
    print("\n[5] Lưu final model cho web app")

    # Chọn model tốt nhất (thường là Random Forest)
    best_name = list(models.keys())[-1]
    best_model = models[best_name]

    path = os.path.join(MODELS_DIR, "final_model.pkl")
    joblib.dump(best_model, path)
    print(f"  ✓ Lưu: final_model.pkl ({best_name})")
    print(f"  → Dùng cho web app: app/app.py")


def main():
    print("=" * 60)
    print("  PIPELINE ĐÁNH GIÁ - Rice Classification")
    print("=" * 60)

    # 0. Đọc
    X_val, y_val, X_test, y_test, models = load_data_and_models()
    os.makedirs(FIGURES_DIR, exist_ok=True)

    # 1. Confusion matrices (val)
    plot_confusion_matrices(models, X_val, y_val)

    # 2. ROC curves
    plot_roc_curves(models, X_val, y_val)

    # 3. TN4: Phân tích lỗi
    errors_df = error_analysis(models, X_val, y_val)

    # 4. Test cuối cùng
    final_results = final_test_evaluation(models, X_test, y_test)

    # 5. Lưu final model
    save_final_model(models)

    print("\n" + "=" * 60)
    print("  ✅ HOÀN THÀNH ĐÁNH GIÁ!")
    print("  → Tiếp theo: python app/app.py (khởi động web)")
    print("=" * 60)


if __name__ == "__main__":
    main()
