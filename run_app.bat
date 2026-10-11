@echo off
chcp 65001 > nul
title Hệ Thống Phân Loại Giống Gạo - Flask Web App
echo ============================================================
echo   🌾 KHỞI ĐỘNG HỆ THỐNG PHÂN LOẠI GIỐNG GẠO (FLASK WEB APP)
echo ============================================================
echo Đang nạp môi trường ảo Python .venv...

cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    echo [✓] Đã tìm thấy môi trường ảo .venv!
    echo [✓] Khởi chạy máy chủ tại: http://localhost:5000
    start http://localhost:5000
    ".venv\Scripts\python.exe" "app\app.py"
) else (
    echo [X] Chưa tìm thấy thư mục .venv!
    pause
)
