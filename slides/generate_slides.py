"""
generate_slides.py - Tạo file presentation.pptx (11 slides) cho buổi bảo vệ kết thúc học phần
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SLIDES_PATH = os.path.join(BASE_DIR, "slides", "presentation.pptx")
FIGURES_DIR = os.path.join(BASE_DIR, "report", "figures")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Colors
C_PRIMARY = RGBColor(24, 76, 120)     # Navy Blue
C_SECONDARY = RGBColor(46, 117, 182)  # Medium Blue
C_DARK = RGBColor(40, 40, 40)         # Charcoal text
C_LIGHT_BG = RGBColor(245, 247, 250)
C_ACCENT = RGBColor(217, 83, 79)      # Warning red

blank_layout = prs.slide_layouts[6]


def add_slide_header(slide, title_text, category_text="HỌC MÁY CƠ BẢN — PROJECT 17"):
    """Thêm thanh header tiêu đề chuẩn."""
    # Category badge
    tb_cat = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.5), Inches(0.4))
    tf_cat = tb_cat.text_frame
    tf_cat.word_wrap = True
    p_cat = tf_cat.paragraphs[0]
    p_cat.text = category_text.upper()
    p_cat.font.size = Pt(11)
    p_cat.font.bold = True
    p_cat.font.color.rgb = C_SECONDARY

    # Title
    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.5), Inches(0.8))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.size = Pt(24)
    p_title.font.bold = True
    p_title.font.color.rgb = C_PRIMARY


# ============================================================
# SLIDE 1: TRANG BÌA
# ============================================================
s1 = prs.slides.add_slide(blank_layout)

tb1 = s1.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.3), Inches(3.5))
tf1 = tb1.text_frame
tf1.word_wrap = True

p1_1 = tf1.paragraphs[0]
p1_1.text = "PROJECT 17: PHÂN LOẠI HAI GIỐNG GẠO"
p1_1.font.size = Pt(36)
p1_1.font.bold = True
p1_1.font.color.rgb = C_PRIMARY

p1_2 = tf1.add_paragraph()
p1_2.text = "TỪ ĐẶC TRƯNG HÌNH HỌC HẠT (CAMMEO VS OSMANCIK)"
p1_2.font.size = Pt(26)
p1_2.font.bold = True
p1_2.font.color.rgb = C_SECONDARY
p1_2.space_before = Pt(8)

p1_3 = tf1.add_paragraph()
p1_3.text = "Chủ đề: Cây quyết định, Cắt tỉa (Pruning) và Rừng ngẫu nhiên (Random Forest)"
p1_3.font.size = Pt(14)
p1_3.font.color.rgb = RGBColor(100, 100, 100)
p1_3.space_before = Pt(16)

p1_4 = tf1.add_paragraph()
p1_4.text = "Nhóm thực hiện: Nhóm 17 ([Thành viên 1] & [Thành viên 2])"
p1_4.font.size = Pt(15)
p1_4.font.color.rgb = C_DARK
p1_4.space_before = Pt(28)


# ============================================================
# SLIDE 2: BỐI CẢNH & CÂU HỎI NGHIÊN CỨU
# ============================================================
s2 = prs.slides.add_slide(blank_layout)
add_slide_header(s2, "1. Bối Cảnh Thực Tiễn & Câu Hỏi Nghiên Cứu")

tb2 = s2.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.2))
tf2 = tb2.text_frame
tf2.word_wrap = True

p = tf2.paragraphs[0]
p.text = "Bối Cảnh Bài Toán:"
p.font.size = Pt(18)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items2 = [
    "• Cơ sở đóng gói nông sản cần sàng lọc sơ bộ hai giống gạo chất lượng cao: Cammeo và Osmancik.",
    "• Phương pháp truyền thống: Giám định bằng mắt thường gây mỏi mệt, độ tin cậy không ổn định.",
    "• Giải pháp đề xuất: Chụp ảnh hạt gạo, trích xuất 7 đặc trưng hình học tự động và ứng dụng Machine Learning."
]
for item in items2:
    p = tf2.add_paragraph()
    p.text = item
    p.font.size = Pt(14)
    p.space_before = Pt(6)

p = tf2.add_paragraph()
p.text = "Câu Hỏi Nghiên Cứu Trọng Tâm:"
p.font.size = Pt(18)
p.font.bold = True
p.font.color.rgb = C_PRIMARY
p.space_before = Pt(18)

items2_q = [
    "❓ 'Cây quyết định tối giản (pruned) có đủ chính xác và dễ giải thích, hay bắt buộc phải dùng Rừng ngẫu nhiên?'",
    "🎯 Tiêu chí thành công:",
    "   1. Đạt độ chính xác (Accuracy) và F1 > 90% (vượt xa Baseline 57.2%).",
    "   2. Pipeline có khả năng tái lập 100%, bảo vệ tập test nghiêm ngặt (không data leakage).",
    "   3. Đóng gói Web App hoàn chỉnh có API, validation và cảnh báo bất định cho mẫu khó."
]
for item in items2_q:
    p = tf2.add_paragraph()
    p.text = item
    p.font.size = Pt(14)
    p.space_before = Pt(6)


# ============================================================
# SLIDE 3: DỮ LIỆU & ĐẶC TẢ ĐẶC TRƯNG
# ============================================================
s3 = prs.slides.add_slide(blank_layout)
add_slide_header(s3, "2. Dữ Liệu Thực Tế: Rice (Cammeo and Osmancik)")

tb3 = s3.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(5.8), Inches(5.2))
tf3 = tb3.text_frame
tf3.word_wrap = True

p = tf3.paragraphs[0]
p.text = "Thông Tin Dữ Liệu:"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items3_left = [
    "• Nguồn: UCI Machine Learning Repository (DOI: 10.24432/C5MW4Z).",
    "• Giấy phép: CC BY 4.0 (Được phép sử dụng học thuật & nghiên cứu).",
    "• Quy mô: 3,810 mẫu hạt gạo (Osmancik: 2,180 - 57.2%, Cammeo: 1,630 - 42.8%).",
    "• Kiểm định chất lượng: 0 missing values, 0 bản ghi trùng lặp."
]
for item in items3_left:
    p = tf3.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(4)

p = tf3.add_paragraph()
p.text = "7 Đặc Trưng Hình Học (Features):"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY
p.space_before = Pt(12)

features_desc = [
    "1. Area: Diện tích hạt (pixel²)",
    "2. Perimeter: Chu vi hạt gạo (pixel)",
    "3. Major_Axis_Length: Trục chính (chiều dài)",
    "4. Minor_Axis_Length: Trục phụ (chiều rộng)",
    "5. Eccentricity: Độ lệch tâm hình elip [0, 1]",
    "6. Convex_Area: Diện tích vỏ lồi nhỏ nhất",
    "7. Extent: Tỷ lệ diện tích / bounding box"
]
for f in features_desc:
    p = tf3.add_paragraph()
    p.text = f"• {f}"
    p.font.size = Pt(12)
    p.space_before = Pt(2)

# Insert figure on right
img_path_corr = os.path.join(FIGURES_DIR, "03_correlation_matrix.png")
if os.path.exists(img_path_corr):
    s3.shapes.add_picture(img_path_corr, Inches(6.9), Inches(1.6), width=Inches(5.7))


# ============================================================
# SLIDE 4: THIẾT KẾ PHÂN CHIA & BASELINE
# ============================================================
s4 = prs.slides.add_slide(blank_layout)
add_slide_header(s4, "3. Thiết Kế Phân Chia Dữ Liệu & Baseline")

tb4 = s4.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.2))
tf4 = tb4.text_frame
tf4.word_wrap = True

p = tf4.paragraphs[0]
p.text = "Chiến Lược Phân Chia Dữ Liệu (Stratified 60/20/20):"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items4 = [
    "• Train Set (60% = 2,286 mẫu): Dùng huấn luyện cây, tính toán ccp_alpha và 5-fold CV.",
    "• Validation Set (20% = 762 mẫu): Dùng chọn độ sâu, chọn mô hình tối ưu, phân tích lỗi.",
    "• Test Set (20% = 762 mẫu): BẢO VỆ TUYỆT ĐỐI — chỉ đánh giá duy nhất 1 lần lúc nghiệm thu.",
    "• Stratification: Đảm bảo tỷ lệ 57.2% Osmancik / 42.8% Cammeo đồng nhất trên cả 3 tập."
]
for item in items4:
    p = tf4.add_paragraph()
    p.text = item
    p.font.size = Pt(14)
    p.space_before = Pt(4)

p = tf4.add_paragraph()
p.text = "Quy Tắc Chống Rò Rỉ Dữ Liệu (Data Leakage):"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY
p.space_before = Pt(14)

items4_leak = [
    "✓ Mọi biến đổi và ngưỡng quyết định được học DUY NHẤT từ Train Set.",
    "✓ Cố định random_state = 42 xuyên suốt toàn bộ pipeline để đảm bảo tính tái lập 100%."
]
for item in items4_leak:
    p = tf4.add_paragraph()
    p.text = item
    p.font.size = Pt(14)
    p.space_before = Pt(4)

p = tf4.add_paragraph()
p.text = "Baseline Đối Chứng (DummyClassifier - Most Frequent):"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY
p.space_before = Pt(14)

p = tf4.add_paragraph()
p.text = "• Luôn dự đoán lớp chiếm đa số (Osmancik) → Accuracy = 57.22%, F1-score = 0.0000\n• Mọi mô hình máy học phải vượt mốc 57.22% này một cách thuyết phục."
p.font.size = Pt(14)
p.space_before = Pt(4)


# ============================================================
# SLIDE 5: THÍ NGHIỆM 1: ĐỘ SÂU MAX_DEPTH & OVERFITTING
# ============================================================
s5 = prs.slides.add_slide(blank_layout)
add_slide_header(s5, "4. Thí Nghiệm 1: Hiệu Năng Theo max_depth (Overfitting)")

tb5 = s5.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(5.8), Inches(5.2))
tf5 = tb5.text_frame
tf5.word_wrap = True

p = tf5.paragraphs[0]
p.text = "Hiện Tượng Quan Sát Được:"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items5 = [
    "• Train Accuracy: Tăng liên tục từ 92.9% lên 100.0% khi max_depth tăng từ 1 lên 20.",
    "• Validation Accuracy: Đạt đỉnh 93.7% ở độ sâu rất nhỏ (depth=1, 2) rồi giảm dần xuống ~89.1% khi cây quá sâu.",
    "• Kết luận thí nghiệm:",
    "   → Cây quyết định không giới hạn bị Overfitting nghiêm trọng, học cả nhiễu của tập huấn luyện.",
    "   → Bắt buộc phải áp dụng Cắt tỉa (Pruning) để tối ưu hóa khả năng tổng quát hóa."
]
for item in items5:
    p = tf5.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(6)

img_path_depth = os.path.join(FIGURES_DIR, "05_depth_curve.png")
if os.path.exists(img_path_depth):
    s5.shapes.add_picture(img_path_depth, Inches(6.8), Inches(1.6), width=Inches(5.8))


# ============================================================
# SLIDE 6: THÍ NGHIỆM 2: PRUNING BẰNG CCP_ALPHA
# ============================================================
s6 = prs.slides.add_slide(blank_layout)
add_slide_header(s6, "5. Thí Nghiệm 2: Cost-Complexity Pruning (ccp_alpha)")

tb6 = s6.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(5.8), Inches(5.2))
tf6 = tb6.text_frame
tf6.word_wrap = True

p = tf6.paragraphs[0]
p.text = "Cơ Chế Cắt Tỉa Cost-Complexity:"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items6 = [
    "• Hàm chi phí: R_alpha(T) = R(T) + alpha * |T|",
    "  (Trong đó R(T) là sai số phân lớp, |T| là số lượng lá, alpha phạt độ phức tạp).",
    "• Tìm 70 giá trị alpha khả dĩ trên cây đầy đủ.",
    "• Đánh giá qua 5-Fold Stratified Cross-Validation:",
    "  → Giá trị tối ưu: ccp_alpha = 0.008024",
    "  → Đạt CV Accuracy = 92.48% ± 1.59%",
    "• Cây sau khi cắt tỉa trở nên cực kỳ gọn gàng, loại bỏ các nhánh nhiễu nhưng giữ nguyên độ chính xác cao."
]
for item in items6:
    p = tf6.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(6)

img_path_pruning = os.path.join(FIGURES_DIR, "06_pruning_curve.png")
if os.path.exists(img_path_pruning):
    s6.shapes.add_picture(img_path_pruning, Inches(6.8), Inches(1.6), width=Inches(5.8))


# ============================================================
# SLIDE 7: THÍ NGHIỆM 3: SO SÁNH CÂY ĐƠN VS RỪNG NGẪU NHIÊN
# ============================================================
s7 = prs.slides.add_slide(blank_layout)
add_slide_header(s7, "6. Thí Nghiệm 3: Cây Quyết Định vs Rừng Ngẫu Nhiên")

tb7 = s7.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.2))
tf7 = tb7.text_frame
tf7.word_wrap = True

p = tf7.paragraphs[0]
p.text = "Bảng So Sánh Hiệu Năng Chi Tiết (Tập Validation & Tập Test):"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

p = tf7.add_paragraph()
p.text = (
    "| Tiêu chí | Baseline (Dummy) | Decision Tree (Pruned) | Random Forest (100 Trees) |\n"
    "|---|---|---|---|\n"
    "| Val Accuracy | 57.22% | 93.70% | 93.44% |\n"
    "| Val F1-score | 0.00% | 92.77% | 92.28% |\n"
    "| Val ROC-AUC | — | 0.9596 | 0.9741 |\n"
    "| Test Accuracy | — | 91.47% | 91.60% |\n"
    "| Test ROC-AUC | — | 0.9344 | 0.9709 |\n"
    "| Thời gian train | 0.001s | 0.036s | 0.320s |\n"
    "| Khả năng diễn giải | Không có | Rất cao (quy tắc if-else rõ ràng) | Hộp đen (phức tạp) |"
)
p.font.size = Pt(13)
p.space_before = Pt(8)

p = tf7.add_paragraph()
p.text = "Trả Lời Câu Hỏi Nghiên Cứu:"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = C_PRIMARY
p.space_before = Pt(14)

items7_ans = [
    "✓ Cây quyết định cắt tỉa (pruned) đạt Accuracy tương đương Rừng ngẫu nhiên (~91.5% - 93.7%).",
    "✓ Tuy nhiên, Rừng ngẫu nhiên có ROC-AUC vượt trội (0.9709 vs 0.9344) và cung cấp xác suất mượt mà hơn.",
    "✓ Nếu ưu tiên giải thích cho nông dân/chuyên gia: Chọn Cây quyết định.",
    "✓ Nếu ưu tiên độ tin cậy xác suất và hệ thống tự động: Chọn Rừng ngẫu nhiên."
]
for item in items7_ans:
    p = tf7.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(4)


# ============================================================
# SLIDE 8: THÍ NGHIỆM 4: PHÂN TÍCH LỖI
# ============================================================
s8 = prs.slides.add_slide(blank_layout)
add_slide_header(s8, "7. Thí Nghiệm 4: Phân Tích Lỗi Sai (Error Analysis)")

tb8 = s8.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(5.8), Inches(5.2))
tf8 = tb8.text_frame
tf8.word_wrap = True

p = tf8.paragraphs[0]
p.text = "Đặc Điểm Các Trường Hợp Dự Đoán Sai (50/762 = 6.6%):"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items8 = [
    "• Sai số hai chiều cân bằng: 27 hạt Cammeo nhầm sang Osmancik, 23 hạt Osmancik nhầm sang Cammeo.",
    "• Nguyên nhân hình thái học:",
    "  → Các hạt bị phân loại sai có kích thước nằm ở vùng giao thoa (Area ~ 12,773, Major_Axis ~ 191.7px).",
    "• Độ tin cậy dự đoán:",
    "  → Mẫu đúng: Độ tự tin trung bình 94.4%.",
    "  → Mẫu sai: Độ tự tin trung bình chỉ đạt 78.3%.",
    "• Bài học kinh nghiệm: Khi mô hình không chắc chắn, xác suất dự đoán sẽ giảm gần ngưỡng 0.5 → Cần cơ chế cờ báo động."
]
for item in items8:
    p = tf8.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(6)

img_path_err = os.path.join(FIGURES_DIR, "12_error_analysis.png")
if os.path.exists(img_path_err):
    s8.shapes.add_picture(img_path_err, Inches(6.8), Inches(1.6), width=Inches(5.8))


# ============================================================
# SLIDE 9: ỨNG DỤNG WEB & TÍNH NĂNG MỞ RỘNG
# ============================================================
s9 = prs.slides.add_slide(blank_layout)
add_slide_header(s9, "8. Ứng Dụng Web Flask & Tính Năng Mở Rộng")

tb9 = s9.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.2))
tf9 = tb9.text_frame
tf9.word_wrap = True

p = tf9.paragraphs[0]
p.text = "Kiến Trúc & Thiết Kế Web App:"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items9 = [
    "• Backend: Python Flask nạp model Offline (final_model.pkl), không huấn luyện lại ở mỗi request.",
    "• Frontend: Giao diện Bootstrap 5 hiện đại, thân thiện, tương tác thời gian thực qua REST API.",
    "• Kiểm tra đầu vào nghiêm ngặt (Input Validation):",
    "  - Báo lỗi khi thiếu dữ liệu hoặc nhập sai định dạng.",
    "  - Cảnh báo rõ ràng khi thông số nằm ngoài miền quan sát [min, max] của tập huấn luyện."
]
for item in items9:
    p = tf9.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(4)

p = tf9.add_paragraph()
p.text = "Tính Năng Mở Rộng (Hoàn Thành Vượt Trội):"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY
p.space_before = Pt(14)

items9_ext = [
    "⭐ 1. Cảnh báo bất định (Uncertainty Detection): Khi xác suất nằm trong khoảng [0.40, 0.60], hệ thống tự động gắn nhãn cảnh báo đỏ và khuyến cáo chuyển sang chuyên gia.",
    "⭐ 2. Phân loại theo lô (Batch CSV Processing): Người dùng có thể upload tệp CSV hàng trăm hạt gạo, hệ thống tính toán tức thì, gắn cờ mẫu nghi ngờ và cho phép xuất file kết quả."
]
for item in items9_ext:
    p = tf9.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(6)


# ============================================================
# SLIDE 10: ĐẠO ĐỨC & SỬ DỤNG CÓ TRÁCH NHIỆM
# ============================================================
s10 = prs.slides.add_slide(blank_layout)
add_slide_header(s10, "9. Đạo Đức, Giới Hạn & Sử Dụng Có Trách Nhiệm")

tb10 = s10.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.2))
tf10 = tb10.text_frame
tf10.word_wrap = True

p = tf10.paragraphs[0]
p.text = "Điều Kiện & Giới Hạn Sử Dụng:"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items10 = [
    "• Không thay thế kiểm nghiệm giống chính thức: Mô hình chỉ đóng vai trò công cụ sàng lọc sơ bộ tại nhà máy.",
    "• Phụ thuộc quy trình trích xuất ảnh: Đầu vào bắt buộc phải cùng độ phân giải và phương pháp tiền xử lý như nghiên cứu gốc.",
    "• Phạm vi giống lúa: Chỉ phân loại giữa Cammeo và Osmancik; không dùng cho các giống lúa khác."
]
for item in items10:
    p = tf10.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(4)

p = tf10.paragraphs[0]
p = tf10.add_paragraph()
p.text = "Quyền Riêng Tư & Minh Bạch:"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY
p.space_before = Pt(14)

items10_eth = [
    "• Dữ liệu công khai, giấy phép CC BY 4.0, không chứa thông tin cá nhân hay nhạy cảm.",
    "• Web app minh bạch: Cung cấp đầy đủ trang Model Card giải thích cấu trúc thuật toán và các rủi ro vận hành."
]
for item in items10_eth:
    p = tf10.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(4)


# ============================================================
# SLIDE 11: KẾT LUẬN & ĐÓNG GÓP NHÓM
# ============================================================
s11 = prs.slides.add_slide(blank_layout)
add_slide_header(s11, "10. Kết Luận & Phân Công Nhóm")

tb11 = s11.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.7), Inches(5.2))
tf11 = tb11.text_frame
tf11.word_wrap = True

p = tf11.paragraphs[0]
p.text = "Tổng Kết Đạt Được:"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items11 = [
    "✓ Xây dựng trọn vẹn quy trình Machine Learning không rò rỉ dữ liệu, tái lập 100%.",
    "✓ Giải quyết triệt để câu hỏi nghiên cứu về sự cân bằng giữa Cây đơn và Rừng ngẫu nhiên.",
    "✓ Triển khai hoàn thiện Web App với đầy đủ tính năng mở rộng (Batch CSV & Uncertainty Check)."
]
for item in items11:
    p = tf11.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(4)

p = tf11.add_paragraph()
p.text = "Phân Công Trách Nhiệm Hai Thành Viên (Cân bằng 50 - 50):"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_PRIMARY
p.space_before = Pt(14)

items11_team = [
    "• Thành viên 1: Quản lý dữ liệu, kiểm soát rò rỉ, xây dựng cây quyết định & thuật toán pruning, lập trình API backend Flask.",
    "• Thành viên 2: Huấn luyện Rừng ngẫu nhiên, thực hiện 4 thí nghiệm đánh giá & phân tích lỗi, thiết kế giao diện frontend & tính năng batch CSV.",
    "• Cả hai thành viên cùng tham gia kiểm thử chéo và soạn thảo báo cáo học thuật."
]
for item in items11_team:
    p = tf11.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(4)

p = tf11.add_paragraph()
p.text = "CẢM ƠN THẦY VÀ CÁC BẠN ĐÃ LẮNG NGHE! XIN MỜI ĐẶT CÂU HỎI (Q&A)"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = C_SECONDARY
p.space_before = Pt(20)

prs.save(SLIDES_PATH)
print(f"Presentation saved successfully to: {SLIDES_PATH}")
