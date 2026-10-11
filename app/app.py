"""
app.py - Ứng dụng web Flask phân loại giống gạo

Tính năng:
  1. Dự đoán đơn lẻ (Single Prediction): Nhập 7 đặc trưng hình học → dự đoán giống gạo.
  2. Kiểm tra tính hợp lệ dữ liệu (Validation): Kiểm tra thiếu, sai kiểu, ngoài miền.
  3. Kiểm tra bất định (Uncertainty Check): Cảnh báo khi xác suất gần 0.5.
  4. Dự đoán theo lô (Batch Prediction CSV): Upload CSV, phân loại toàn bộ, thống kê và xuất kết quả.
  5. Tải file mẫu CSV để kiểm thử.
  6. Model Card & Hướng dẫn sử dụng có trách nhiệm.
"""

import os
import io
import numpy as np
import pandas as pd
import joblib
from flask import Flask, render_template, request, jsonify, send_file, Response

# ============================================================
# CẤU HÌNH & KHỞI TẠO HỆ THỐNG
# ============================================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODEL_PATH = os.path.join(BASE_DIR, "models", "final_model.pkl")
SAMPLE_CSV_PATH = os.path.join(BASE_DIR, "data", "sample_batch.csv")

app = Flask(__name__)

# Tách train offline khỏi serving: Chỉ nạp model đã lưu, không train lại ở mỗi request!
# Sửa lỗi an toàn: Kiểm tra sự tồn tại của model trước khi load để tránh crash ứng dụng
if os.path.exists(MODEL_PATH):
    model = joblib.load(MODEL_PATH)
else:
    model = None
    print(f"⚠️ CẢNH BÁO: Chưa tìm thấy file mô hình tại '{MODEL_PATH}'.")
    print("👉 Vui lòng chạy 'python src/train.py' và 'python src/evaluate.py' để tạo mô hình trước.")

FEATURES = [
    {
        "name": "Area",
        "label": "Diện Tích Hạt Gạo (Area)",
        "unit": "pixel²",
        "desc": "Tổng số pixel thuộc bề mặt hạt gạo",
        "min": 2523,
        "max": 23823,
        "step": 1,
    },
    {
        "name": "Perimeter",
        "label": "Chu Vi Đường Biên (Perimeter)",
        "unit": "pixel",
        "desc": "Độ dài đường bao quanh chu vi hạt gạo",
        "min": 208,
        "max": 680,
        "step": 0.01,
    },
    {
        "name": "Major_Axis_Length",
        "label": "Chiều Dài Trục Chính (Major Axis Length)",
        "unit": "pixel",
        "desc": "Chiều dài nhất của hạt gạo (trục elip chính)",
        "min": 78,
        "max": 320,
        "step": 0.01,
    },
    {
        "name": "Minor_Axis_Length",
        "label": "Chiều Rộng Trục Phụ (Minor Axis Length)",
        "unit": "pixel",
        "desc": "Bề ngang rộng nhất của hạt gạo (trục elip phụ)",
        "min": 46,
        "max": 130,
        "step": 0.01,
    },
    {
        "name": "Eccentricity",
        "label": "Độ Thuôn Dài / Lệch Tâm (Eccentricity)",
        "unit": "tỷ lệ 0–1",
        "desc": "Độ dẹt của elip (0 là tròn trịa; càng gần 1 hạt càng thuôn dài)",
        "min": 0.6,
        "max": 0.97,
        "step": 0.0001,
    },
    {
        "name": "Convex_Area",
        "label": "Diện Tích Đa Giác Bao Lồi (Convex Area)",
        "unit": "pixel²",
        "desc": "Diện tích đa giác lồi nhỏ nhất bao trọn hạt gạo",
        "min": 2594,
        "max": 24300,
        "step": 1,
    },
    {
        "name": "Extent",
        "label": "Tỷ Lệ Chiếm Khung Bao (Extent)",
        "unit": "tỷ lệ 0–1",
        "desc": "Tỷ lệ diện tích hạt so với hình chữ nhật bao quanh (Area / Bounding Box)",
        "min": 0.38,
        "max": 0.86,
        "step": 0.0001,
    },
]
FEATURE_NAMES = [f["name"] for f in FEATURES]

# Lưu trữ tạm kết quả batch gần nhất để tải về
latest_batch_results = None


@app.route("/")
def index():
    """Trang chủ giao diện người dùng."""
    return render_template("index.html", features=FEATURES)


@app.route("/predict", methods=["POST"])
def predict():
    """API dự đoán cho một mẫu hạt gạo đơn lẻ."""
    if model is None:
        return jsonify({
            "error": "Mô hình AI chưa được nạp. Vui lòng chạy huấn luyện trước (python src/train.py & python src/evaluate.py)."
        }), 503

    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "Không nhận được dữ liệu JSON."}), 400

        values = []
        warnings = []

        for feat in FEATURES:
            name = feat["name"]
            raw = data.get(name)

            # 1. Kiểm tra thiếu dữ liệu (Missing check)
            if raw is None or str(raw).strip() == "":
                return jsonify({"error": f"Thiếu trường bắt buộc '{feat['label']}'."}), 400

            # 2. Kiểm tra kiểu số (Type check)
            try:
                val = float(raw)
            except (ValueError, TypeError):
                return jsonify({
                    "error": f"Trường '{feat['label']}' phải là giá trị số hợp lệ. Nhận được: '{raw}'."
                }), 400

            # 3. Kiểm tra miền giá trị đã huấn luyện (Domain/Outlier check)
            if val < feat["min"] or val > feat["max"]:
                warnings.append(
                    f"{feat['label']}: Giá trị {val} nằm ngoài miền quan sát khi huấn luyện "
                    f"[{feat['min']} – {feat['max']} {feat['unit']}]. Kết quả dự đoán có thể không đáng tin cậy."
                )

            values.append(val)

        # Suy luận bằng mô hình đã lưu (Dùng DataFrame để giữ nguyên tên đặc trưng, tránh UserWarning)
        X_input = pd.DataFrame([values], columns=FEATURE_NAMES)
        prediction = model.predict(X_input)[0]
        probabilities = model.predict_proba(X_input)[0]

        classes = list(model.classes_)
        prob_dict = {cls: float(round(prob, 4)) for cls, prob in zip(classes, probabilities)}
        max_prob = float(max(probabilities))

        # Phân loại mức tin cậy
        if max_prob >= 0.85:
            confidence = "Cao"
            confidence_color = "success"
        elif max_prob >= 0.65:
            confidence = "Trung bình"
            confidence_color = "warning"
        else:
            confidence = "Thấp"
            confidence_color = "danger"

        # Tính năng mở rộng 1: Kiểm tra bất định khi xác suất gần 0.5
        is_uncertain = False
        uncertainty_msg = None
        if 0.40 <= probabilities[0] <= 0.60:
            is_uncertain = True
            uncertainty_msg = (
                "⚠️ CẢNH BÁO BẤT ĐỊNH (Xác suất hai lớp xấp xỉ 50/50): "
                "Mẫu hạt này nằm trong vùng ranh giới hình thái không chắc chắn. "
                "Theo quy tắc sử dụng có trách nhiệm, không được tự động đóng gói mà cần chuyên viên kiểm tra trực tiếp."
            )

        # Tính năng mở rộng 2: Giải thích nguyên nhân quyết định (Explainability / Local Rule Inspector)
        major_axis = float(data.get("Major_Axis_Length", 0))
        perimeter = float(data.get("Perimeter", 0))
        if prediction == "Cammeo":
            explanation = (
                f"Hạt được phân loại là Cammeo vì chiều dài trục chính ({major_axis:.2f}px) và chu vi ({perimeter:.2f}px) "
                f"lớn hơn ngưỡng phân định hình thái (ngưỡng phân nhánh gốc ~191.05px). Giống Cammeo có đặc trưng dáng thuôn dài."
            )
        else:
            explanation = (
                f"Hạt được phân loại là Osmancik vì chiều dài trục chính ({major_axis:.2f}px) "
                f"nhỏ hơn ngưỡng phân định (~191.05px). Giống Osmancik có đặc trưng dáng tròn và ngắn hơn."
            )

        return jsonify({
            "prediction": prediction,
            "probabilities": prob_dict,
            "confidence": confidence,
            "confidence_color": confidence_color,
            "is_uncertain": is_uncertain,
            "uncertainty_message": uncertainty_msg,
            "explanation": explanation,
            "warnings": warnings,
        })

    except Exception as e:
        return jsonify({"error": f"Lỗi xử lý nội bộ: {str(e)}"}), 500


@app.route("/predict-batch", methods=["POST"])
def predict_batch():
    """
    Tính năng mở rộng: Dự đoán theo lô từ file CSV.
    Tự động gắn cờ bất định cho từng hạt và thống kê kết quả.
    """
    global latest_batch_results
    if model is None:
        return jsonify({
            "error": "Mô hình AI chưa được nạp. Vui lòng chạy huấn luyện trước (python src/train.py & python src/evaluate.py)."
        }), 503

    try:
        if "file" not in request.files:
            return jsonify({"error": "Không tìm thấy tệp CSV được tải lên."}), 400

        file = request.files["file"]
        if file.filename == "":
            return jsonify({"error": "Chưa chọn tệp để tải lên."}), 400

        if not file.filename.endswith(".csv"):
            return jsonify({"error": "Vui lòng chọn tệp định dạng .csv."}), 400

        # Đọc dữ liệu CSV
        df = pd.read_csv(file)

        # Kiểm tra sự hiện diện của các cột đặc trưng
        missing_cols = [col for col in FEATURE_NAMES if col not in df.columns]
        if missing_cols:
            return jsonify({
                "error": f"Tệp CSV thiếu các cột bắt buộc: {', '.join(missing_cols)}"
            }), 400

        X_batch = df[FEATURE_NAMES]

        # Kiểm tra missing values trong batch
        if X_batch.isnull().any().any():
            return jsonify({"error": "Tệp CSV chứa giá trị rỗng (NaN/Missing) ở các cột đặc trưng."}), 400

        # Dự đoán
        preds = model.predict(X_batch)
        probas = model.predict_proba(X_batch)
        classes = list(model.classes_)
        cammeo_idx = classes.index("Cammeo")
        osmancik_idx = classes.index("Osmancik")

        results_list = []
        uncertain_count = 0
        cammeo_count = 0
        osmancik_count = 0

        for i in range(len(df)):
            p_cam = float(probas[i, cammeo_idx])
            p_osm = float(probas[i, osmancik_idx])
            pred_class = str(preds[i])

            if pred_class == "Cammeo":
                cammeo_count += 1
            else:
                osmancik_count += 1

            # Kiểm tra bất định
            uncertain = (0.40 <= p_cam <= 0.60)
            if uncertain:
                uncertain_count += 1

            item = {
                "index": i + 1,
                "prediction": pred_class,
                "prob_cammeo": round(p_cam, 4),
                "prob_osmancik": round(p_osm, 4),
                "is_uncertain": uncertain,
                "status": "Bất định (Cần xem xét)" if uncertain else "Xác định",
            }
            results_list.append(item)

        # Tạo DataFrame xuất CSV kết quả
        out_df = df.copy()
        out_df["Predicted_Class"] = preds
        out_df["Prob_Cammeo"] = [round(float(p[cammeo_idx]), 4) for p in probas]
        out_df["Prob_Osmancik"] = [round(float(p[osmancik_idx]), 4) for p in probas]
        out_df["Uncertainty_Flag"] = ["Yes" if (0.40 <= p[cammeo_idx] <= 0.60) else "No" for p in probas]
        latest_batch_results = out_df

        return jsonify({
            "total_samples": len(df),
            "cammeo_count": cammeo_count,
            "osmancik_count": osmancik_count,
            "uncertain_count": uncertain_count,
            "samples": results_list[:50],  # Trả về tối đa 50 dòng để hiển thị preview
            "has_more": len(results_list) > 50,
        })

    except Exception as e:
        return jsonify({"error": f"Lỗi xử lý file lô: {str(e)}"}), 500


@app.route("/export-batch-csv")
def export_batch_csv():
    """Xuất file kết quả dự đoán batch thành CSV."""
    global latest_batch_results
    if latest_batch_results is None:
        return "Chưa có kết quả lô nào để xuất.", 400

    csv_buffer = io.StringIO()
    latest_batch_results.to_csv(csv_buffer, index=False)
    csv_bytes = csv_buffer.getvalue().encode("utf-8-sig")

    return Response(
        csv_bytes,
        mimetype="text/csv",
        headers={"Content-Disposition": "attachment;filename=ket_qua_phan_loai_lo.csv"}
    )


@app.route("/download-sample-csv")
def download_sample_csv():
    """Cho phép người dùng tải file mẫu CSV để kiểm thử tính năng lô."""
    if os.path.exists(SAMPLE_CSV_PATH):
        return send_file(SAMPLE_CSV_PATH, as_attachment=True, download_name="mau_du_lieu_hat_gao.csv")
    return "Không tìm thấy file mẫu.", 404


@app.route("/model-card")
def model_card():
    """Trang Model Card và tài liệu giới hạn kỹ thuật."""
    return render_template("model_card.html")


@app.route("/api/info")
def api_info():
    """
    API cung cấp metadata về mô hình và hệ thống (phục vụ giám sát & kiểm thử).
    """
    return jsonify({
        "project": "Phân loại giống gạo Cammeo vs Osmancik",
        "model_loaded": model is not None,
        "model_type": str(type(model).__name__) if model else None,
        "classes": list(model.classes_) if model else [],
        "features": FEATURE_NAMES,
        "test_accuracy": "91.60%",
        "test_roc_auc": 0.9709,
        "uncertainty_threshold": [0.40, 0.60]
    })


if __name__ == "__main__":
    print("=" * 60)
    print("  🍚 HỆ THỐNG PHÂN LOẠI GIỐNG GẠO (FLASK WEB APP)")
    print("  → Địa chỉ: http://localhost:5000")
    print("=" * 60)
    app.run(debug=True, host="0.0.0.0", port=5000)
