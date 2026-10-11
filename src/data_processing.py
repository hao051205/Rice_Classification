"""
data_processing.py - Tiền xử lý và chia dữ liệu Rice (Cammeo & Osmancik)

Chức năng:
  1. Đọc file ARFF gốc
  2. Kiểm tra chất lượng dữ liệu (missing, trùng lặp, ngoại lệ)
  3. EDA cơ bản: phân phối, tương quan, boxplot
  4. Chia stratified train/val/test (60/20/20)
  5. Lưu các tập dữ liệu ra CSV
"""

import os
import sys
import numpy as np
import pandas as pd
from scipy.io import arff
from sklearn.model_selection import train_test_split
import matplotlib
matplotlib.use('Agg')  # Non-interactive backend
import matplotlib.pyplot as plt
import seaborn as sns

# ============================================================
# CẤU HÌNH
# ============================================================
RANDOM_SEED = 42
TEST_SIZE = 0.2
VAL_SIZE = 0.25  # 25% của phần còn lại = 20% tổng

# Đường dẫn tương đối (không dùng đường dẫn tuyệt đối!)
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
RAW_DATA_PATH = os.path.join(BASE_DIR, "data", "raw", "Rice_Cammeo_Osmancik.arff")
PROCESSED_DIR = os.path.join(BASE_DIR, "data", "processed")
FIGURES_DIR = os.path.join(BASE_DIR, "report", "figures")

FEATURE_COLS = [
    "Area", "Perimeter", "Major_Axis_Length", "Minor_Axis_Length",
    "Eccentricity", "Convex_Area", "Extent"
]
TARGET_COL = "Class"


def load_data(path: str) -> pd.DataFrame:
    """Đọc file ARFF và trả về DataFrame."""
    print(f"[1/5] Đọc dữ liệu từ: {os.path.basename(path)}")
    data, meta = arff.loadarff(path)
    df = pd.DataFrame(data)
    # Decode bytes → string cho cột Class
    df[TARGET_COL] = df[TARGET_COL].str.decode("utf-8")
    print(f"  ✓ Shape: {df.shape}")
    print(f"  ✓ Columns: {list(df.columns)}")
    return df


def check_quality(df: pd.DataFrame) -> None:
    """Kiểm tra chất lượng dữ liệu."""
    print(f"\n[2/5] Kiểm tra chất lượng dữ liệu")

    # Missing values
    missing = df.isnull().sum()
    print(f"  ✓ Missing values:\n{missing.to_string()}")

    # Trùng lặp
    n_dup = df.duplicated().sum()
    print(f"  ✓ Bản ghi trùng lặp: {n_dup}")
    if n_dup > 0:
        print(f"    ⚠ Phát hiện {n_dup} dòng trùng — giữ nguyên (đặt cùng split nếu cần)")

    # Phân phối lớp
    class_counts = df[TARGET_COL].value_counts()
    class_pct = df[TARGET_COL].value_counts(normalize=True) * 100
    print(f"  ✓ Phân phối lớp:")
    for cls in class_counts.index:
        print(f"    - {cls}: {class_counts[cls]} ({class_pct[cls]:.1f}%)")

    # Thống kê mô tả
    print(f"\n  ✓ Thống kê mô tả:")
    print(df[FEATURE_COLS].describe().round(2).to_string())


def run_eda(df: pd.DataFrame) -> None:
    """EDA: tạo biểu đồ phân phối, boxplot, tương quan."""
    print(f"\n[3/5] Tạo biểu đồ EDA")
    os.makedirs(FIGURES_DIR, exist_ok=True)

    # --- 3a. Phân phối đặc trưng theo lớp ---
    fig, axes = plt.subplots(2, 4, figsize=(18, 9))
    axes = axes.flatten()
    for i, col in enumerate(FEATURE_COLS):
        ax = axes[i]
        for cls in ["Cammeo", "Osmancik"]:
            subset = df[df[TARGET_COL] == cls][col]
            ax.hist(subset, bins=30, alpha=0.6, label=cls, edgecolor="black")
        ax.set_title(col, fontsize=11, fontweight="bold")
        ax.set_xlabel("")
        ax.legend(fontsize=8)
    axes[-1].axis("off")  # Ẩn ô thừa
    fig.suptitle("Phân phối đặc trưng theo giống gạo", fontsize=14, fontweight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, "01_distributions.png"), dpi=150)
    plt.close()
    print("  ✓ Lưu: 01_distributions.png")

    # --- 3b. Boxplot phát hiện ngoại lệ ---
    fig, axes = plt.subplots(2, 4, figsize=(18, 9))
    axes = axes.flatten()
    for i, col in enumerate(FEATURE_COLS):
        ax = axes[i]
        df.boxplot(column=col, by=TARGET_COL, ax=ax)
        ax.set_title(col, fontsize=11, fontweight="bold")
        ax.set_xlabel("")
    axes[-1].axis("off")
    fig.suptitle("Boxplot đặc trưng theo giống gạo", fontsize=14, fontweight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, "02_boxplots.png"), dpi=150)
    plt.close()
    print("  ✓ Lưu: 02_boxplots.png")

    # --- 3c. Ma trận tương quan ---
    fig, ax = plt.subplots(figsize=(10, 8))
    corr = df[FEATURE_COLS].corr()
    sns.heatmap(corr, annot=True, fmt=".2f", cmap="coolwarm",
                center=0, square=True, ax=ax)
    ax.set_title("Ma trận tương quan giữa các đặc trưng", fontsize=13, fontweight="bold")
    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, "03_correlation_matrix.png"), dpi=150)
    plt.close()
    print("  ✓ Lưu: 03_correlation_matrix.png")

    # --- 3d. Pairplot (2 đặc trưng quan trọng nhất) ---
    fig, ax = plt.subplots(figsize=(8, 6))
    for cls, color in [("Cammeo", "royalblue"), ("Osmancik", "tomato")]:
        subset = df[df[TARGET_COL] == cls]
        ax.scatter(subset["Major_Axis_Length"], subset["Minor_Axis_Length"],
                   alpha=0.4, label=cls, color=color, s=15)
    ax.set_xlabel("Major_Axis_Length")
    ax.set_ylabel("Minor_Axis_Length")
    ax.set_title("Scatter: Major vs Minor Axis Length", fontsize=13, fontweight="bold")
    ax.legend()
    plt.tight_layout()
    plt.savefig(os.path.join(FIGURES_DIR, "04_scatter_major_minor.png"), dpi=150)
    plt.close()
    print("  ✓ Lưu: 04_scatter_major_minor.png")


def split_data(df: pd.DataFrame) -> dict:
    """
    Chia dữ liệu: 60% train, 20% validation, 20% test (stratified).

    ⚠ QUAN TRỌNG:
    - Chia TRƯỚC mọi bước xử lý/học từ dữ liệu
    - Dùng stratify để giữ tỷ lệ lớp
    - Cố định random_state để tái lập
    """
    print(f"\n[4/5] Chia dữ liệu (stratified: 60/20/20)")

    X = df[FEATURE_COLS]
    y = df[TARGET_COL]

    # Bước 1: Tách test (20%)
    X_temp, X_test, y_temp, y_test = train_test_split(
        X, y, test_size=TEST_SIZE, random_state=RANDOM_SEED, stratify=y
    )

    # Bước 2: Tách val (25% của 80% = 20% tổng)
    X_train, X_val, y_train, y_val = train_test_split(
        X_temp, y_temp, test_size=VAL_SIZE, random_state=RANDOM_SEED, stratify=y_temp
    )

    splits = {
        "train": (X_train, y_train),
        "val": (X_val, y_val),
        "test": (X_test, y_test),
    }

    # In thông tin
    for name, (X_split, y_split) in splits.items():
        pct = y_split.value_counts(normalize=True)
        print(f"  ✓ {name:5s}: {len(X_split):5d} mẫu | "
              f"Cammeo {pct.get('Cammeo', 0)*100:.1f}% | "
              f"Osmancik {pct.get('Osmancik', 0)*100:.1f}%")

    return splits


def save_splits(splits: dict) -> None:
    """Lưu các tập dữ liệu ra CSV."""
    print(f"\n[5/5] Lưu dữ liệu đã chia")
    os.makedirs(PROCESSED_DIR, exist_ok=True)

    for name, (X, y) in splits.items():
        df_out = X.copy()
        df_out[TARGET_COL] = y
        path = os.path.join(PROCESSED_DIR, f"{name}.csv")
        df_out.to_csv(path, index=False)
        print(f"  ✓ Lưu: {name}.csv ({len(df_out)} dòng)")


def main():
    """Chạy toàn bộ pipeline xử lý dữ liệu."""
    print("=" * 60)
    print("  PIPELINE XỬ LÝ DỮ LIỆU - Rice Classification")
    print("=" * 60)

    # 1. Đọc dữ liệu
    df = load_data(RAW_DATA_PATH)

    # 2. Kiểm tra chất lượng
    check_quality(df)

    # 3. EDA
    run_eda(df)

    # 4. Chia dữ liệu
    splits = split_data(df)

    # 5. Lưu
    save_splits(splits)

    print("\n" + "=" * 60)
    print("  ✅ HOÀN THÀNH XỬ LÝ DỮ LIỆU!")
    print("  → Tiếp theo: python src/train.py")
    print("=" * 60)


if __name__ == "__main__":
    main()
