@echo off
chcp 65001 > nul
title Huấn Luyện & Đánh Giá Toàn Bộ Pipeline Machine Learning
echo ============================================================
echo   🌾 CHẠY TOÀN BỘ QUY TRÌNH HỌC MÁY (ML PIPELINE)
echo ============================================================

cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo [X] Lỗi: Chưa tìm thấy môi trường ảo .venv!
    pause
    exit /b
)

echo.
echo [1/4] TIỀN XỬ LÝ VÀ CHIA DỮ LIỆU STRATIFIED...
".venv\Scripts\python.exe" "src\data_processing.py"
if %errorlevel% neq 0 ( echo [X] Lỗi bước 1! & pause & exit /b )

echo.
echo [2/4] HUẤN LUYỆN MÔ HÌNH VÀ CẮT TỈA CCP_ALPHA...
".venv\Scripts\python.exe" "src\train.py"
if %errorlevel% neq 0 ( echo [X] Lỗi bước 2! & pause & exit /b )

echo.
echo [3/4] ĐÁNH GIÁ MÔ HÌNH VÀ PHÂN TÍCH LỖI TRÊN TẬP TEST...
".venv\Scripts\python.exe" "src\evaluate.py"
if %errorlevel% neq 0 ( echo [X] Lỗi bước 3! & pause & exit /b )

echo.
echo [4/4] CHẠY KIỂM THỬ TỰ ĐỘNG CÁC ĐẦU MÚT API...
".venv\Scripts\python.exe" "tests_endpoint.py"

echo.
echo ============================================================
echo   🎉 HOÀN TẤT TOÀN BỘ PIPELINE THÀNH CÔNG RỰC RỠ!
echo ============================================================
pause
