"""
tests_endpoint.py - Bộ kiểm thử tự động toàn diện cho hệ thống Flask & Mô hình AI.

Chức năng:
  - Tự động kiểm thử API trực tiếp qua test_client (không cần khởi động server thủ công).
  - Kiểm tra dự đoán đơn lẻ (Test 1: Single Prediction)
  - Kiểm tra tải file mẫu CSV (Test 2: Download Sample CSV)
  - Kiểm tra dự đoán theo lô (Test 3: Batch Prediction)
  - Kiểm tra xác thực đầu vào & cảnh báo ngoài miền (Test 4: Input Validation & Boundary Checks)
  - Kiểm tra API Metadata (Test 5: System & Model Info)
"""

import os
import sys
import io
import json

# Khắc phục lỗi font Unicode trên Windows console
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

# Đảm bảo import được module app từ thư mục project
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(CURRENT_DIR, "app"))

try:
    from app import app, model
except ImportError:
    # Trường hợp chạy từ thư mục app
    from app import app, model


def run_tests():
    print("=" * 60)
    print("🚀 BẮT ĐẦU CHẠY KIỂM THỬ TỰ ĐỘNG (AUTOMATED TEST SUITE)")
    print("=" * 60)

    # Khởi tạo Test Client nội bộ của Flask
    client = app.test_client()

    # -------------------------------------------------------------
    # TEST 1: DỰ ĐOÁN ĐƠN LẺ (SINGLE PREDICTION)
    # -------------------------------------------------------------
    print("\n[TEST 1] Kiểm tra API dự đoán 1 hạt gạo (/predict)...")
    payload = {
        "Area": 13000,
        "Perimeter": 460.0,
        "Major_Axis_Length": 190.0,
        "Minor_Axis_Length": 87.0,
        "Eccentricity": 0.89,
        "Convex_Area": 13300,
        "Extent": 0.65
    }
    resp = client.post("/predict", json=payload)
    assert resp.status_code == 200, f"Lỗi HTTP {resp.status_code}: {resp.data}"
    data = json.loads(resp.data.decode("utf-8"))

    print(f"  ✓ Kết quả dự đoán : {data['prediction']}")
    print(f"  ✓ Xác suất        : {data['probabilities']}")
    print(f"  ✓ Mức độ tin cậy  : {data['confidence']}")
    print(f"  ✓ Giải thích      : {data.get('explanation', '')[:60]}...")
    assert data["prediction"] in ["Cammeo", "Osmancik"]
    assert "confidence" in data

    # -------------------------------------------------------------
    # TEST 2: TẢI FILE MẪU CSV (/download-sample-csv)
    # -------------------------------------------------------------
    print("\n[TEST 2] Kiểm tra tải tệp CSV mẫu (/download-sample-csv)...")
    resp = client.get("/download-sample-csv")
    assert resp.status_code == 200, f"Lỗi HTTP {resp.status_code}"
    lines = resp.data.decode("utf-8").splitlines()
    print(f"  ✓ Nhận được {len(lines)} dòng dữ liệu từ file mẫu.")
    assert len(lines) > 1, "File mẫu rỗng!"

    # -------------------------------------------------------------
    # TEST 3: DỰ ĐOÁN THEO LÔ (BATCH PREDICTION CSV)
    # -------------------------------------------------------------
    print("\n[TEST 3] Kiểm tra xử lý hàng loạt theo lô (/predict-batch)...")
    sample_csv_path = os.path.join(CURRENT_DIR, "data", "sample_batch.csv")
    with open(sample_csv_path, "rb") as f:
        file_bytes = f.read()

    data_form = {
        "file": (io.BytesIO(file_bytes), "sample_batch.csv")
    }
    resp = client.post("/predict-batch", data=data_form, content_type="multipart/form-data")
    assert resp.status_code == 200, f"Lỗi HTTP {resp.status_code}: {resp.data}"
    res_batch = json.loads(resp.data.decode("utf-8"))

    print(f"  ✓ Tổng số mẫu đã phân loại: {res_batch['total_samples']}")
    print(f"  ✓ Số hạt Cammeo          : {res_batch['cammeo_count']}")
    print(f"  ✓ Số hạt Osmancik        : {res_batch['osmancik_count']}")
    print(f"  ✓ Số hạt bất định (gần ranh giới): {res_batch['uncertain_count']}")
    assert res_batch["total_samples"] > 0

    # -------------------------------------------------------------
    # TEST 4: KIỂM TRA VALIDATION ĐẦU VÀO (LỖI THIẾU DỮ LIỆU)
    # -------------------------------------------------------------
    print("\n[TEST 4] Kiểm tra phòng vệ dữ liệu sai/thiếu (Input Validation)...")
    # Thiếu trường Area
    invalid_payload = {
        "Perimeter": 460.0,
        "Major_Axis_Length": 190.0,
        "Minor_Axis_Length": 87.0
    }
    resp_invalid = client.post("/predict", json=invalid_payload)
    assert resp_invalid.status_code == 400, "Phải trả về lỗi 400 khi thiếu dữ liệu!"
    err_data = json.loads(resp_invalid.data.decode("utf-8"))
    print(f"  ✓ Hệ thống chặn thành công: '{err_data['error']}'")

    # -------------------------------------------------------------
    # TEST 5: METADATA & THÔNG TIN HỆ THỐNG (/api/info)
    # -------------------------------------------------------------
    print("\n[TEST 5] Kiểm tra API Metadata hệ thống (/api/info)...")
    resp_info = client.get("/api/info")
    assert resp_info.status_code == 200
    info_data = json.loads(resp_info.data.decode("utf-8"))
    print(f"  ✓ Dự án             : {info_data['project']}")
    print(f"  ✓ Mô hình đã load   : {info_data['model_loaded']}")
    print(f"  ✓ Độ chính xác test : {info_data['test_accuracy']}")

    print("\n" + "=" * 60)
    print("🎉 TẤT CẢ 5 BÀI TEST ĐỀU ĐẠT CHUẨN 100% (ALL TESTS PASSED)!")
    print("=" * 60)


if __name__ == "__main__":
    run_tests()
