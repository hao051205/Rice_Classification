"""
train.py - Huấn luyện mô hình phân loại gạo

Các mô hình:
  1. Baseline (DummyClassifier - most_frequent)
  2. Decision Tree (với pruning bằng ccp_alpha)
  3. Random Forest

Thí nghiệm bắt buộc:
  TN1: Đường hiệu năng theo max_depth
  TN2: Pruning bằng ccp_alpha trên Cross-Validation
  TN3: So sánh cây đơn vs rừng ngẫu nhiên

⚠ Mọi model chỉ fit trên train, đánh giá trên val.
  Test set KHÔNG ĐƯỢC DÙNG ở bước này!
"""

import os
import sys
import time
import json

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

from sklearn.dummy import DummyClassifier
from sklearn.tree import DecisionTreeClassifier, plot_tree
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import cross_val_score
from sklearn.metrics import (
    accuracy_score, f1_score, roc_auc_score,
    classification_report, confusion_matrix
)

# ============================================================
# CẤU HÌNH
# ============================================================
RANDOM_SEED = 42
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROCESSED_DIR = os.path.join(BASE_DIR, "data", "processed")
MODELS_DIR = os.path.join(BASE_DIR, "models")
FIGURES_DIR = os.path.join(BASE_DIR, "report", "figures")

FEATURE_COLS = [
    "Area", "Perimeter", "Major_Axis_Length", "Minor_Axis_Length",
    "Eccentricity", "Convex_Area", "Extent"
]
TARGET_COL = "Class"


def load_splits():
    """Đọc dữ liệu đã chia."""
    print("[0/6] Đọc dữ liệu đã chia...")
    splits = {}
    for name in ["train", "val"]:
        df = pd.read_csv(os.path.join(PROCESSED_DIR, f"{name}.csv"))
        splits[name] = (df[FEATURE_COLS], df[TARGET_COL])
        print(f"  ✓ {name}: {len(df)} mẫu")
    return splits


def train_baseline(X_train, y_train, X_val, y_val):
    """Baseline: Dự đoán lớp phổ biến nhất."""
    print("\n[1/6] Huấn luyện BASELINE (Most Frequent)")
    print("  → Đây là mức tối thiểu mà mô hình phải vượt qua.")

    baseline = DummyClassifier(strategy="most_frequent", random_state=RANDOM_SEED)
    baseline.fit(X_train, y_train)
    y_pred = baseline.predict(X_val)

    acc = accuracy_score(y_val, y_pred)
    f1 = f1_score(y_val, y_pred, pos_label="Cammeo", zero_division=0)

    print(f"  ✓ Baseline Accuracy: {acc:.4f}")
    print(f"  ✓ Baseline F1:       {f1:.4f}")
    print(f"  → Mô hình chính phải vượt Accuracy > {acc:.4f}")

    return baseline, {"accuracy": acc, "f1": f1}


def experiment_depth_curve(X_train, y_train, X_val, y_val):
    """TN1: Đường hiệu năng theo max_depth."""
    print("\n[2/6] THÍ NGHIỆM 1: Hiệu năng theo max_depth")

    depths = list(range(1, 21))
    train_accs, val_accs = [], []

    for d in depths:
        dt = DecisionTreeClassifier(max_depth=d, random_state=RANDOM_SEED)
        dt.fit(X_train, y_train)
        train_accs.append(accuracy_score(y_train, dt.predict(X_train)))
        val_accs.append(accuracy_score(y_val, dt.predict(X_val)))

    # Tìm depth tốt nhất trên val
    best_depth = depths[np.argmax(val_accs)]
    best_val_acc = max(val_accs)
    print(f"  ✓ Best max_depth = {best_depth} (Val Accuracy = {best_val_acc:.4f})")

    # Vẽ biểu đồ
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.plot(depths, train_accs, "b-o", label="Train Accuracy", markersize=5)
    ax.plot(depths, val_accs, "r-o", label="Validation Accuracy", markersize=5)
    ax.axvline(x=best_depth, color="green", linestyle="--",
               label=f"Best depth = {best_depth}", alpha=0.7)
    ax.set_xlabel("max_depth", fontsize=12)
    ax.set_ylabel("Accuracy", fontsize=12)
    ax.set_title("TN1: Hiệu năng theo max_depth", fontsize=14, fontweight="bold")
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, "05_depth_curve.png"), dpi=150)
    plt.close()
    print("  ✓ Lưu: 05_depth_curve.png")

    return best_depth


def experiment_pruning(X_train, y_train):
    """TN2: Pruning bằng ccp_alpha trên 5-fold Cross-Validation."""
    print("\n[3/6] THÍ NGHIỆM 2: Pruning bằng ccp_alpha (5-fold CV)")

    # Tìm tất cả effective alpha
    dt_full = DecisionTreeClassifier(random_state=RANDOM_SEED)
    dt_full.fit(X_train, y_train)
    path = dt_full.cost_complexity_pruning_path(X_train, y_train)
    ccp_alphas = path.ccp_alphas
    print(f"  → Tìm được {len(ccp_alphas)} giá trị alpha")

    # CV cho mỗi alpha
    cv_means, cv_stds = [], []
    for alpha in ccp_alphas:
        dt = DecisionTreeClassifier(ccp_alpha=alpha, random_state=RANDOM_SEED)
        scores = cross_val_score(dt, X_train, y_train, cv=5, scoring="accuracy")
        cv_means.append(scores.mean())
        cv_stds.append(scores.std())

    cv_means = np.array(cv_means)
    cv_stds = np.array(cv_stds)

    # Chọn alpha tốt nhất
    best_idx = np.argmax(cv_means)
    best_alpha = ccp_alphas[best_idx]
    best_cv_acc = cv_means[best_idx]
    print(f"  ✓ Best ccp_alpha = {best_alpha:.6f}")
    print(f"  ✓ Best CV Accuracy = {best_cv_acc:.4f} ± {cv_stds[best_idx]:.4f}")

    # Vẽ biểu đồ
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.errorbar(ccp_alphas, cv_means, yerr=cv_stds, fmt="g-o", markersize=3,
                alpha=0.7, capsize=2, label="CV Accuracy ± std")
    ax.axvline(x=best_alpha, color="red", linestyle="--",
               label=f"Best alpha = {best_alpha:.6f}", alpha=0.7)
    ax.set_xlabel("ccp_alpha", fontsize=12)
    ax.set_ylabel("CV Accuracy", fontsize=12)
    ax.set_title("TN2: Pruning — CV Accuracy theo ccp_alpha", fontsize=14, fontweight="bold")
    ax.legend(fontsize=10)
    ax.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, "06_pruning_curve.png"), dpi=150)
    plt.close()
    print("  ✓ Lưu: 06_pruning_curve.png")

    return best_alpha


def experiment_compare(X_train, y_train, X_val, y_val, best_alpha, baseline_metrics):
    """TN3: So sánh Decision Tree (pruned) vs Random Forest."""
    print("\n[4/6] THÍ NGHIỆM 3: So sánh Cây quyết định vs Rừng ngẫu nhiên")

    models = {
        "Decision Tree (pruned)": DecisionTreeClassifier(
            ccp_alpha=best_alpha, random_state=RANDOM_SEED
        ),
        "Random Forest (100 trees)": RandomForestClassifier(
            n_estimators=100, random_state=RANDOM_SEED, n_jobs=-1
        ),
    }

    results = {}

    for name, model in models.items():
        print(f"\n  --- {name} ---")
        start = time.time()
        model.fit(X_train, y_train)
        train_time = time.time() - start

        y_pred = model.predict(X_val)
        y_proba = model.predict_proba(X_val)

        # Xác định index của lớp Osmancik (lớp dương mặc định theo thứ tự chữ cái của scikit-learn)
        classes = list(model.classes_)
        osmancik_idx = classes.index("Osmancik")

        acc = accuracy_score(y_val, y_pred)
        f1 = f1_score(y_val, y_pred, pos_label="Cammeo")
        auc = roc_auc_score(y_val, y_proba[:, osmancik_idx])

        results[name] = {
            "accuracy": round(acc, 4),
            "f1": round(f1, 4),
            "roc_auc": round(auc, 4),
            "train_time_s": round(train_time, 4),
        }

        print(f"  ✓ Accuracy:   {acc:.4f}")
        print(f"  ✓ F1:         {f1:.4f}")
        print(f"  ✓ ROC-AUC:    {auc:.4f}")
        print(f"  ✓ Train time: {train_time:.4f}s")
        print(f"\n{classification_report(y_val, y_pred)}")

    # Bảng so sánh
    print("\n  📊 BẢNG SO SÁNH:")
    print(f"  {'Model':<30s} {'Accuracy':>9s} {'F1':>8s} {'ROC-AUC':>9s} {'Time':>8s}")
    print(f"  {'-'*30} {'-'*9} {'-'*8} {'-'*9} {'-'*8}")
    print(f"  {'Baseline':30s} {baseline_metrics['accuracy']:9.4f} {baseline_metrics['f1']:8.4f} {'N/A':>9s} {'—':>8s}")
    for name, m in results.items():
        print(f"  {name:30s} {m['accuracy']:9.4f} {m['f1']:8.4f} {m['roc_auc']:9.4f} {m['train_time_s']:7.4f}s")

    return models, results


def visualize_tree(model, feature_names):
    """Trực quan hóa cây quyết định đã pruned."""
    print("\n[5/6] Trực quan hóa cây quyết định")

    fig, ax = plt.subplots(figsize=(24, 12))
    plot_tree(
        model,
        feature_names=feature_names,
        class_names=list(model.classes_),
        filled=True,
        rounded=True,
        max_depth=4,
        fontsize=9,
        ax=ax,
    )
    ax.set_title("Cây quyết định (pruned) — hiển thị 4 tầng đầu", fontsize=16, fontweight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, "07_decision_tree.png"), dpi=150, bbox_inches="tight")
    plt.close()
    print("  ✓ Lưu: 07_decision_tree.png")

    # Feature importance
    fi = pd.Series(model.feature_importances_, index=feature_names).sort_values(ascending=True)
    fig, ax = plt.subplots(figsize=(8, 5))
    fi.plot(kind="barh", ax=ax, color="steelblue", edgecolor="black")
    ax.set_title("Feature Importance — Decision Tree", fontsize=13, fontweight="bold")
    ax.set_xlabel("Importance")
    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, "08_feature_importance_tree.png"), dpi=150)
    plt.close()
    print("  ✓ Lưu: 08_feature_importance_tree.png")


def save_models(models, results, best_alpha):
    """Lưu model và metadata."""
    print("\n[6/6] Lưu mô hình")
    os.makedirs(MODELS_DIR, exist_ok=True)

    for name, model in models.items():
        fname = name.lower().replace(" ", "_").replace("(", "").replace(")", "")
        path = os.path.join(MODELS_DIR, f"{fname}.pkl")
        joblib.dump(model, path)
        print(f"  ✓ Lưu: {os.path.basename(path)}")

    # Lưu metadata
    meta = {
        "random_seed": RANDOM_SEED,
        "best_ccp_alpha": float(best_alpha),
        "results": results,
    }
    meta_path = os.path.join(MODELS_DIR, "training_metadata.json")
    with open(meta_path, "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2, ensure_ascii=False)
    print(f"  ✓ Lưu: training_metadata.json")


def main():
    print("=" * 60)
    print("  PIPELINE HUẤN LUYỆN - Rice Classification")
    print("=" * 60)

    # 0. Đọc dữ liệu
    splits = load_splits()
    X_train, y_train = splits["train"]
    X_val, y_val = splits["val"]

    os.makedirs(FIGURES_DIR, exist_ok=True)

    # 1. Baseline
    baseline_model, baseline_metrics = train_baseline(X_train, y_train, X_val, y_val)

    # 2. TN1: max_depth
    best_depth = experiment_depth_curve(X_train, y_train, X_val, y_val)

    # 3. TN2: pruning
    best_alpha = experiment_pruning(X_train, y_train)

    # 4. TN3: So sánh
    models, results = experiment_compare(X_train, y_train, X_val, y_val, best_alpha, baseline_metrics)

    # 5. Trực quan hóa cây
    dt_model = models["Decision Tree (pruned)"]
    visualize_tree(dt_model, FEATURE_COLS)

    # Feature importance cho Random Forest
    rf_model = models["Random Forest (100 trees)"]
    fi = pd.Series(rf_model.feature_importances_, index=FEATURE_COLS).sort_values(ascending=True)
    fig, ax = plt.subplots(figsize=(8, 5))
    fi.plot(kind="barh", ax=ax, color="forestgreen", edgecolor="black")
    ax.set_title("Feature Importance — Random Forest", fontsize=13, fontweight="bold")
    ax.set_xlabel("Importance")
    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, "09_feature_importance_rf.png"), dpi=150)
    plt.close()
    print("  ✓ Lưu: 09_feature_importance_rf.png")

    # 6. Lưu
    save_models(models, results, best_alpha)

    print("\n" + "=" * 60)
    print("  ✅ HOÀN THÀNH HUẤN LUYỆN!")
    print("  → Tiếp theo: python src/evaluate.py")
    print("=" * 60)


if __name__ == "__main__":
    main()
