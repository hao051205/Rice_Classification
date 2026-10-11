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
# CẤU HÌNH & KHỞI TẠO
# ============================================================
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MODEL_PATH = os.path.join(BASE_DIR, "models", "final_model.pkl")
SAMPLE_CSV_PATH = os.path.join(BASE_DIR, "data", "sample_batch.csv")

app = Flask(__name__)

# Tách train offline khỏi serving: Chỉ nạp model đã lưu, không train lại ở mỗi request!
model = joblib.load(MODEL_PATH)

FEATURES = [
    {"name": "Area",              "label": "Area (pixel²)",          "min": 2523,  "max": 23823, "step": 1,     "demo": 15231},
    {"name": "Perimeter",         "label": "Perimeter (pixel)",      "min": 208,   "max": 680,   "step": 0.01,  "demo": 525.58},
    {"name": "Major_Axis_Length", "label": "Major Axis Length (px)",  "min": 78,    "max": 320,   "step": 0.01,  "demo": 229.75},
    {"name": "Minor_Axis_Length", "label": "Minor Axis Length (px)",  "min": 46,    "max": 130,   "step": 0.01,  "demo": 85.09},
    {"name": "Eccentricity",     "label": "Eccentricity [0-1]",      "min": 0.6,   "max": 0.97,  "step": 0.0001,"demo": 0.9289},
    {"name": "Convex_Area",      "label": "Convex Area (pixel²)",    "min": 2594,  "max": 24300, "step": 1,     "demo": 15617},
    {"name": "Extent",           "label": "Extent [0-1]",            "min": 0.38,  "max": 0.86,  "step": 0.0001,"demo": 0.5729},
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
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "Không nhận được dữ liệu JSON."}), 400

        values = []
        warnings = []

        for feat in FEATURES:
            name = feat["name"]
            raw = data.get(name)

            # 1. Kiểm tra thiếu
            if raw is None or str(raw).strip() == "":
                return jsonify({"error": f"Thiếu trường bắt buộc '{feat['label']}'."}), 400

            # 2. Kiểm tra kiểu số
            try:
                val = float(raw)
            except (ValueError, TypeError):
                return jsonify({
                    "error": f"Trường '{feat['label']}' phải là giá trị số hợp lệ. Nhận được: '{raw}'."
                }), 400

            # 3. Kiểm tra miền giá trị đã huấn luyện
            if val < feat["min"] or val > feat["max"]:
                warnings.append(
                    f"{feat['label']}: Giá trị {val} nằm ngoài miền quan sát khi huấn luyện "
                    f"[{feat['min']}, {feat['max']}]. Kết quả dự đoán có thể không đáng tin cậy."
                )

            values.append(val)

        # Suy luận bằng mô hình đã lưu
        X_input = np.array(values).reshape(1, -1)
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

        # Tính năng mở rộng: Kiểm tra bất định khi xác suất gần 0.5
        is_uncertain = False
        uncertainty_msg = None
        if 0.40 <= probabilities[0] <= 0.60:
            is_uncertain = True
            uncertainty_msg = (
                "⚠️ CẢNH BÁO BẤT ĐỊNH (Xác suất hai lớp xấp xỉ 50/50): "
                "Mẫu hạt này nằm trong vùng ranh giới hình thái không chắc chắn. "
                "Theo quy tắc sử dụng có trách nhiệm, không được tự động đóng gói mà cần chuyên viên kiểm tra trực tiếp."
            )

        return jsonify({
            "prediction": prediction,
            "probabilities": prob_dict,
            "confidence": confidence,
            "confidence_color": confidence_color,
            "is_uncertain": is_uncertain,
            "uncertainty_message": uncertainty_msg,
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


if __name__ == "__main__":
    print("=" * 60)
    print("  🍚 HỆ THỐNG PHÂN LOẠI GIỐNG GẠO (FLASK WEB APP)")
    print("  → Địa chỉ: http://localhost:5000")
    print("=" * 60)
    app.run(debug=True, host="0.0.0.0", port=5000)
