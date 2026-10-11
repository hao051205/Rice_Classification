"""
build_structured_slides.py - Tạo file PowerPoint (12 slides) chuẩn cấu trúc thuyết trình học thuật
Tự động gắn Lời thuyết trình (Speaker Notes), nhúng hình ảnh thực nghiệm và căn chỉnh layout 16:9 hiện đại.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SLIDES_PATH = os.path.join(BASE_DIR, "slides", "presentation.pptx")
ALT_PATH = os.path.join(BASE_DIR, "slides", "presentation_final.pptx")
FIGURES_DIR = os.path.join(BASE_DIR, "report", "figures")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
blank_layout = prs.slide_layouts[6]

# Palettes
C_NAVY_DARK = RGBColor(16, 42, 67)      # #102A43
C_PRIMARY = RGBColor(20, 52, 98)        # #143462
C_SECONDARY = RGBColor(36, 59, 83)      # #243B53
C_ACCENT = RGBColor(217, 119, 6)        # #D97706
C_BLUE_ACCENT = RGBColor(37, 99, 165)   # #2563A5
C_DARK = RGBColor(33, 37, 41)
C_MUTED = RGBColor(100, 116, 139)
C_CARD_BG = RGBColor(248, 250, 252)
C_CARD_BORDER = RGBColor(226, 232, 240)
C_WHITE = RGBColor(255, 255, 255)
C_GREEN = RGBColor(22, 101, 52)
C_GREEN_BG = RGBColor(240, 253, 244)
C_RED = RGBColor(185, 28, 28)
C_RED_BG = RGBColor(254, 242, 242)


def add_slide_header(slide, title_text, section_text="PROJECT 17: PHÂN LOẠI HAI GIỐNG LÚA GẠO CAMMEO VÀ OSMANCIK"):
    """Header chuẩn học thuật."""
    tb_sec = slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11.733), Inches(0.35))
    tf_sec = tb_sec.text_frame
    tf_sec.word_wrap = True
    p_sec = tf_sec.paragraphs[0]
    p_sec.text = section_text.upper()
    p_sec.font.size = Pt(11)
    p_sec.font.bold = True
    p_sec.font.color.rgb = C_BLUE_ACCENT

    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.62), Inches(11.733), Inches(0.72))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.size = Pt(22)
    p_title.font.bold = True
    p_title.font.color.rgb = C_NAVY_DARK

    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.36), Inches(11.733), Inches(0.025))
    line.fill.solid()
    line.fill.fore_color.rgb = C_BLUE_ACCENT
    line.line.color.rgb = C_BLUE_ACCENT


def add_slide_footer(slide, current_page, total_pages=12):
    """Footer với số trang và thông tin học phần."""
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.92), Inches(11.733), Inches(0.015))
    line.fill.solid()
    line.fill.fore_color.rgb = C_CARD_BORDER
    line.line.color.rgb = C_CARD_BORDER

    tb_l = slide.shapes.add_textbox(Inches(0.8), Inches(6.98), Inches(6.0), Inches(0.35))
    tf_l = tb_l.text_frame
    p_l = tf_l.paragraphs[0]
    p_l.text = "Học máy cơ bản — Báo cáo bảo vệ Project 17 | PGS.TS. Nguyễn Văn Hậu"
    p_l.font.size = Pt(10)
    p_l.font.color.rgb = C_MUTED

    tb_r = slide.shapes.add_textbox(Inches(10.0), Inches(6.98), Inches(2.533), Inches(0.35))
    tf_r = tb_r.text_frame
    p_r = tf_r.paragraphs[0]
    p_r.alignment = PP_ALIGN.RIGHT
    p_r.text = f"Slide {current_page} / {total_pages}"
    p_r.font.size = Pt(10.5)
    p_r.font.bold = True
    p_r.font.color.rgb = C_BLUE_ACCENT


def add_card(slide, left, top, width, height, bg_color=C_CARD_BG, border_color=C_CARD_BORDER):
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1)
    return card


def set_speaker_notes(slide, notes_text):
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    tf.text = notes_text.strip()


# ============================================================
# SLIDE 1: TIÊU ĐỀ
# ============================================================
s1 = prs.slides.add_slide(blank_layout)

tb_inst = s1.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.95))
tf_inst = tb_inst.text_frame
p = tf_inst.paragraphs[0]
p.text = "BỘ GIÁO DỤC VÀ ĐÀO TẠO — TRƯỜNG ĐẠI HỌC SƯ PHẠM KỸ THUẬT HƯNG YÊN"
p.font.size = Pt(13.5)
p.font.bold = True
p.font.color.rgb = C_PRIMARY
p.alignment = PP_ALIGN.CENTER

p2 = tf_inst.add_paragraph()
p2.text = "KHOA CÔNG NGHỆ THÔNG TIN — BỘ MÔN HỆ THỐNG THÔNG TIN"
p2.font.size = Pt(12)
p2.font.bold = True
p2.font.color.rgb = C_BLUE_ACCENT
p2.alignment = PP_ALIGN.CENTER
p2.space_before = Pt(3)

add_card(s1, Inches(0.8), Inches(1.45), Inches(11.733), Inches(3.25), bg_color=C_CARD_BG, border_color=C_BLUE_ACCENT)
tb_title = s1.shapes.add_textbox(Inches(1.1), Inches(1.6), Inches(11.133), Inches(2.9))
tf_title = tb_title.text_frame
tf_title.word_wrap = True

p = tf_title.paragraphs[0]
p.text = "BÁO CÁO BẢO VỆ ĐỒ ÁN HỌC MÁY CƠ BẢN (PROJECT 17)"
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = C_ACCENT
p.alignment = PP_ALIGN.CENTER

p = tf_title.add_paragraph()
p.text = "PHÂN LOẠI HAI GIỐNG LÚA GẠO CAMMEO VÀ OSMANCIK\nBẰNG CÂY QUYẾT ĐỊNH VÀ RỪNG NGẪU NHIÊN"
p.font.size = Pt(26)
p.font.bold = True
p.font.color.rgb = C_NAVY_DARK
p.alignment = PP_ALIGN.CENTER
p.space_before = Pt(8)

p = tf_title.add_paragraph()
p.text = "Đánh giá thực nghiệm Cost-Complexity Pruning (ccp_alpha), Random Forest và triển khai Web Service"
p.font.size = Pt(13)
p.font.italic = True
p.font.color.rgb = C_MUTED
p.alignment = PP_ALIGN.CENTER
p.space_before = Pt(8)

# Left info: GVHD
add_card(s1, Inches(0.8), Inches(4.85), Inches(5.7), Inches(1.8), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_sup = s1.shapes.add_textbox(Inches(1.0), Inches(5.0), Inches(5.3), Inches(1.5))
tf_sup = tb_sup.text_frame
p = tf_sup.paragraphs[0]
p.text = "GIẢNG VIÊN HƯỚNG DẪN:"
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = C_BLUE_ACCENT

p = tf_sup.add_paragraph()
p.text = "PGS. TS. Nguyễn Văn Hậu"
p.font.size = Pt(17)
p.font.bold = True
p.font.color.rgb = C_NAVY_DARK
p.space_before = Pt(3)

p = tf_sup.add_paragraph()
p.text = "Bộ môn Hệ thống thông tin — Khoa CNTT"
p.font.size = Pt(11)
p.font.color.rgb = C_MUTED
p.space_before = Pt(2)

# Right info: SVTH
add_card(s1, Inches(6.833), Inches(4.85), Inches(5.7), Inches(1.8), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_stu = s1.shapes.add_textbox(Inches(7.033), Inches(5.0), Inches(5.3), Inches(1.5))
tf_stu = tb_stu.text_frame
p = tf_stu.paragraphs[0]
p.text = "SINH VIÊN THỰC HIỆN (NHÓM 17):"
p.font.size = Pt(11)
p.font.bold = True
p.font.color.rgb = C_BLUE_ACCENT

p = tf_stu.add_paragraph()
p.text = "1. Chử Quốc Hào  —  MSV: 10123112"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = C_NAVY_DARK
p.space_before = Pt(4)

p = tf_stu.add_paragraph()
p.text = "2. [Thành viên 2]  —  MSV: [Mã số SV 2]"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = C_NAVY_DARK
p.space_before = Pt(2)

p = tf_stu.add_paragraph()
p.text = "Lớp: 10123O.1 — Ngành Công nghệ thông tin"
p.font.size = Pt(11)
p.font.color.rgb = C_MUTED
p.space_before = Pt(3)

set_speaker_notes(s1, """
Kính thưa Thầy giáo và toàn thể các bạn! Hôm nay, nhóm 17 xin phép được báo cáo đề tài: Phân loại hai giống lúa gạo Cammeo và Osmancik bằng Cây quyết định, Cắt tỉa ccp_alpha và Rừng ngẫu nhiên. Dự án của chúng em không chỉ dừng lại ở mô hình lý thuyết, mà hướng tới việc giải quyết bài toán tự động hóa thực tế trong dây chuyền đóng gói nông sản.
""")


# ============================================================
# SLIDE 2: ĐẶT VẤN ĐỀ / ĐỘNG LỰC
# ============================================================
s2 = prs.slides.add_slide(blank_layout)
add_slide_header(s2, "Giám Định Thủ Công Bộc Lộ Nhiều Sai Sót Và Làm Chậm Tốc Độ Đóng Gói", "BỐI CẢNH THỰC TIỄN & ĐỘNG LỰC ĐỀ TÀI")
add_slide_footer(s2, 2)

# Left Column (Content)
add_card(s2, Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.15))
tb_c2 = s2.shapes.add_textbox(Inches(1.0), Inches(1.75), Inches(5.3), Inches(4.75))
tf_c2 = tb_c2.text_frame
tf_c2.word_wrap = True

bullets_s2 = [
    ("Bối cảnh thương phẩm:", "Cammeo và Osmancik có giá thành và thị trường khác biệt."),
    ("Hạn chế giám định mắt:", "Công nhân dễ mỏi mắt, phân loại chậm và sai lệch sau ca dài."),
    ("Hậu quả lẫn tạp:", "Lẫn tạp giống làm giảm sút chất lượng và uy tín thương hiệu."),
    ("Giải pháp số hóa:", "Máy ảnh chụp hạt gạo trên băng chuyền và đo đạc hình học tự động."),
    ("Mục tiêu tự động hóa:", "Máy học phân loại tức thì trong vài mili-giây, bảo đảm độ thuần chủng.")
]
for i, (b_title, b_desc) in enumerate(bullets_s2):
    p = tf_c2.paragraphs[0] if i == 0 else tf_c2.add_paragraph()
    p.text = f"- {b_title} "
    p.font.bold = True
    p.font.size = Pt(13.5)
    p.font.color.rgb = C_NAVY_DARK
    p.space_before = Pt(8) if i > 0 else Pt(0)
    
    run = p.add_run()
    run.text = b_desc
    run.font.bold = False
    run.font.size = Pt(13)
    run.font.color.rgb = C_DARK

# Right Column (Figure)
img_p = os.path.join(FIGURES_DIR, "01_distributions.png")
if os.path.exists(img_p):
    s2.shapes.add_picture(img_p, Inches(6.75), Inches(1.55), width=Inches(5.78))
    # Caption box
    tb_cap = s2.shapes.add_textbox(Inches(6.75), Inches(6.35), Inches(5.78), Inches(0.35))
    tb_cap.text_frame.paragraphs[0].text = "Hình 1: Phân bố 7 đặc trưng hình học giữa Cammeo (xanh) và Osmancik (cam)"
    tb_cap.text_frame.paragraphs[0].font.size = Pt(9.5)
    tb_cap.text_frame.paragraphs[0].font.italic = True
    tb_cap.text_frame.paragraphs[0].font.color.rgb = C_MUTED
    tb_cap.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

set_speaker_notes(s2, """
Trong ngành xuất khẩu gạo, Cammeo và Osmancik là hai giống lúa chất lượng cao nhưng có giá thành khác nhau. Nếu phân loại thủ công bằng mắt, công nhân nhanh chóng mệt mỏi dẫn đến nhầm lẫn, gây lẫn tạp lô hàng. Hệ thống chụp ảnh trên băng chuyền kết hợp máy học giúp đo đạc các kích thước vi mô và ra quyết định chính xác chỉ trong vài mili-giây.
""")


# ============================================================
# SLIDE 3: PHÁT BIỂU BÀI TOÁN
# ============================================================
s3 = prs.slides.add_slide(blank_layout)
add_slide_header(s3, "Bài Toán Phân Loại Nhị Phân Dựa Trên 7 Đặc Trưng Hình Học", "ĐẶC TẢ BÀI TOÁN HỌC MÁY")
add_slide_footer(s3, 3)

# 4 Key Cards Layout
coords_s3 = [
    (Inches(0.8), Inches(1.55), Inches(5.7), Inches(2.4), "LOẠI BÀI TOÁN & QUY MÔ DỮ LIỆU", [
        "- Phân loại nhị phân có giám sát (Binary Classification).",
        "- Đơn vị quan sát: Đúng 1 hạt gạo đơn lẻ duy nhất.",
        "- Bộ dữ liệu: 3.810 mẫu chuẩn từ UCI Repository (CC BY 4.0).",
        "- Phân bố: Osmancik 57,2% (2.180 hạt) | Cammeo 42,8% (1.630 hạt)."
    ]),
    (Inches(6.833), Inches(1.55), Inches(5.7), Inches(2.4), "KHÔNG GIAN ĐẦU VÀO & ĐẦU RA", [
        "- Đầu vào: Véc-tơ 7 chiều số thực: X in R^7.",
        "- 7 chỉ số: Diện tích, Chu vi, Trục chính, Trục phụ, Độ thon...",
        "- Đầu ra: Nhãn y in {Cammeo, Osmancik}.",
        "- Kèm theo: Phân phối xác suất và độ tin cậy của mô hình."
    ]),
    (Inches(0.8), Inches(4.2), Inches(5.7), Inches(2.5), "Ý NGHĨA HÌNH THÁI HỌC 7 ĐẶC TRƯNG", [
        "- Area, Perimeter, Major_Axis: Kích thước tổng thể hạt gạo.",
        "- Minor_Axis, Convex_Area: Bề ngang và diện tích bao lồi hạt.",
        "- Eccentricity: Đo độ thuôn dài (hạt thon hay hạt tròn).",
        "- Extent: Tỷ lệ diện tích hạt chiếm khung chữ nhật bao ngoài."
    ]),
    (Inches(6.833), Inches(4.2), Inches(5.7), Inches(2.5), "TIÊU CHÍ THÀNH CÔNG TỐI THIỂU", [
        "- Vượt trội hoàn toàn mốc Dummy Baseline (57,22%).",
        "- Đạt Accuracy & F1-Score > 90% trên tập Test độc lập.",
        "- Quy trình không rò rỉ dữ liệu, có khả năng tái lập 100%.",
        "- Đóng gói thành Web Service có kiểm soát dữ liệu đầu vào."
    ])
]

for left, top, w, h, title, items in coords_s3:
    add_card(s3, left, top, w, h, bg_color=C_WHITE, border_color=C_CARD_BORDER)
    tb = s3.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), w - Inches(0.4), h - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_BLUE_ACCENT
    for item in items:
        p_item = tf.add_paragraph()
        p_item.text = item
        p_item.font.size = Pt(11.5)
        p_item.font.color.rgb = C_DARK
        p_item.space_before = Pt(3)

set_speaker_notes(s3, """
Bài toán của đề tài là phân loại nhị phân có giám sát với đơn vị quan sát là đúng 1 hạt gạo. Đầu vào là véc-tơ 7 chiều gồm các đặc trưng hình học 2D trích xuất từ 3.810 ảnh mẫu thu thập từ UCI. Mục tiêu là dự đoán chính xác giống lúa với độ chính xác trên 90%, đồng thời đảm bảo quy trình thực nghiệm độc lập và chống rò rỉ dữ liệu.
""")


# ============================================================
# SLIDE 4: TIẾP CẬN HIỆN CÓ VÀ HẠN CHẾ
# ============================================================
s4 = prs.slides.add_slide(blank_layout)
add_slide_header(s4, "Đa Cộng Tuyến Mạnh Và Rò Rỉ Dữ Liệu Là Thách Thức Cốt Lõi", "CÁC TIẾP CẬN HIỆN CÓ & HẠN CHẾ")
add_slide_footer(s4, 4)

add_card(s4, Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.15))
tb_c4 = s4.shapes.add_textbox(Inches(1.0), Inches(1.75), Inches(5.3), Inches(4.75))
tf_c4 = tb_c4.text_frame
tf_c4.word_wrap = True

bullets_s4 = [
    ("Hiện tượng đa cộng tuyến:", "Nhóm 4 biến kích cỡ tương quan rất cao (r > 0.95)."),
    ("Hạn chế Hồi quy Logistic:", "Ma trận nghịch đảo bị suy biến, trọng số bị méo mó."),
    ("Hạn chế Cây tự do:", "Dễ học vẹt toàn bộ mẫu nhiễu, phương sai cao (Overfitting)."),
    ("Hạn chế Rừng ngẫu nhiên:", "Mô hình hộp đen phức tạp, khó giải thích cho công nhân."),
    ("Câu hỏi nghiên cứu:", "Cây cắt tỉa có đủ chính xác và dễ hiểu, hay cần Rừng ngẫu nhiên?")
]
for i, (b_title, b_desc) in enumerate(bullets_s4):
    p = tf_c4.paragraphs[0] if i == 0 else tf_c4.add_paragraph()
    p.text = f"- {b_title} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_NAVY_DARK
    p.space_before = Pt(8) if i > 0 else Pt(0)
    
    run = p.add_run()
    run.text = b_desc
    run.font.bold = False
    run.font.size = Pt(12.5)
    run.font.color.rgb = C_DARK

img_p = os.path.join(FIGURES_DIR, "03_correlation_matrix.png")
if os.path.exists(img_p):
    s4.shapes.add_picture(img_p, Inches(6.75), Inches(1.55), width=Inches(5.78))
    tb_cap = s4.shapes.add_textbox(Inches(6.75), Inches(6.35), Inches(5.78), Inches(0.35))
    tb_cap.text_frame.paragraphs[0].text = "Hình 2: Ma trận tương quan Pearson khẳng định hiện tượng đa cộng tuyến (r > 0.95)"
    tb_cap.text_frame.paragraphs[0].font.size = Pt(9.5)
    tb_cap.text_frame.paragraphs[0].font.italic = True
    tb_cap.text_frame.paragraphs[0].font.color.rgb = C_MUTED
    tb_cap.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

set_speaker_notes(s4, """
Khảo sát tương quan cho thấy nhóm 4 biến kích cỡ có hệ số r trên 0.95. Nếu áp dụng mô hình hồi quy tuyến tính, đa cộng tuyến sẽ làm trọng số bị bất ổn định. Mặt khác, Cây quyết định tự do lại dễ bị quá khớp, còn Rừng ngẫu nhiên thì khó giải thích. Điều này đặt ra câu hỏi: Cây tối giản có đủ tốt để thay thế Rừng ngẫu nhiên hay không?
""")


# ============================================================
# SLIDE 5: PHƯƠNG PHÁP ĐỀ XUẤT - CHIA DỮ LIỆU & BASELINE
# ============================================================
s5 = prs.slides.add_slide(blank_layout)
add_slide_header(s5, "Thiết Lập Quy Trình Stratified 60/20/20 Và Mốc Dummy Baseline", "PHƯƠNG PHÁP ĐỀ XUẤT: QUY TRÌNH CHỐNG RÒ RỈ")
add_slide_footer(s5, 5)

cards_s5 = [
    (Inches(0.8), Inches(1.55), Inches(3.64), Inches(5.15), "1. PHÂN CHIA TẦNG (STRATIFIED)", [
        ("- Chia tầng 60/20/20:", "Giữ nguyên tỷ lệ 57,2% Osmancik và 42,8% Cammeo."),
        ("- Train Set (2.286 mẫu):", "Học cây, lập bảng ccp_alpha và 5-Fold Cross Validation."),
        ("- Validation Set (762 mẫu):", "Dò tìm siêu tham số và so sánh đối chứng mô hình."),
        ("- Test Set (762 mẫu):", "Niêm phong tuyệt đối, chỉ nạp 1 lần để nghiệm thu."),
        ("- Chống Leakage:", "Mọi thống kê min, max chỉ tính toán trên tập Train.")
    ]),
    (Inches(4.84), Inches(1.55), Inches(3.64), Inches(5.15), "2. MÔ HÌNH CƠ SỞ (BASELINE)", [
        ("- Ý nghĩa khoa học:", "Không bao giờ đánh giá mô hình trong sự cô lập."),
        ("- DummyClassifier:", "Chiến lược 'most_frequent' luôn luôn đoán nhãn Osmancik."),
        ("- Baseline Accuracy:", "57,22% (đúng bằng tỷ lệ của lớp đa số)."),
        ("- Baseline F1-Score:", "0,0000 cho lớp Cammeo; ROC-AUC = 0.5000."),
        ("- Cột mốc đối chứng:", "Mọi thuật toán đề xuất bắt buộc phải vượt xa mốc 57,22%.")
    ]),
    (Inches(8.88), Inches(1.55), Inches(3.64), Inches(5.15), "3. QUY TRÌNH 4 THÍ NGHIỆM", [
        ("- Thí nghiệm 1:", "Khảo sát max_depth chứng minh hiện tượng Overfitting."),
        ("- Thí nghiệm 2:", "Cắt tỉa Cost-Complexity Pruning bằng 5-Fold CV."),
        ("- Thí nghiệm 3:", "So sánh toàn diện Cây đơn cắt tỉa vs Rừng ngẫu nhiên."),
        ("- Thí nghiệm 4:", "Phân tích 50 ca đoán sai tại vùng giao thoa hình thái."),
        ("- Đóng gói sản phẩm:", "Triển khai mô hình vào Web Service Flask phục vụ thực tế.")
    ])
]

for left, top, w, h, title, items in cards_s5:
    add_card(s5, left, top, w, h, bg_color=C_WHITE, border_color=C_CARD_BORDER)
    tb = s5.shapes.add_textbox(left + Inches(0.2), top + Inches(0.2), w - Inches(0.4), h - Inches(0.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_BLUE_ACCENT
    for b_title, b_desc in items:
        p_item = tf.add_paragraph()
        p_item.text = f"{b_title} "
        p_item.font.bold = True
        p_item.font.size = Pt(11.5)
        p_item.font.color.rgb = C_NAVY_DARK
        p_item.space_before = Pt(6)
        
        run = p_item.add_run()
        run.text = b_desc
        run.font.bold = False
        run.font.size = Pt(11)
        run.font.color.rgb = C_DARK

set_speaker_notes(s5, """
Để chống rò rỉ dữ liệu, nhóm phân chia phân tầng Stratified 60% Train, 20% Val và 20% Test, trong đó tập Test được niêm phong tuyệt đối. Chúng em cũng xây dựng Dummy Baseline luôn đoán nhãn đa số đạt sẵn 57,22% làm cột mốc sàn. Toàn bộ nghiên cứu được tổ chức tuần tự qua 4 thí nghiệm thực chứng chặt chẽ.
""")


# ============================================================
# SLIDE 6: THÍ NGHIỆM 1 - KHẢO SÁT OVERFITTING CÂY
# ============================================================
s6 = prs.slides.add_slide(blank_layout)
add_slide_header(s6, "Cây Quyết Định Tự Do Bị Quá Khớp Nghiêm Trọng Khi Tăng Độ Sâu", "THÍ NGHIỆM 1: KHẢO SÁT OVERFITTING THEO CHIỀU SÂU CÂY")
add_slide_footer(s6, 6)

add_card(s6, Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.15))
tb_c6 = s6.shapes.add_textbox(Inches(1.0), Inches(1.75), Inches(5.3), Inches(4.75))
tf_c6 = tb_c6.text_frame
tf_c6.word_wrap = True

bullets_s6 = [
    ("Thiết kế thí nghiệm:", "Cho max_depth tăng dần từ 1 đến 20 trên thuật toán CART."),
    ("Tập Train bùng nổ:", "Accuracy tăng vọt lên 100,0% tuyệt đối khi depth >= 16."),
    ("Tập Validation suy giảm:", "Đạt đỉnh 93,70% ở depth=1, sau đó giảm sút về 89,10%."),
    ("Bản chất toán học:", "Cây học thuộc lòng cả các mẫu dị biệt và nhiễu (High Variance)."),
    ("Kết luận Thí nghiệm 1:", "Bắt buộc phải áp dụng kỹ thuật Cắt tỉa (Pruning) để kiểm soát cây.")
]
for i, (b_title, b_desc) in enumerate(bullets_s6):
    p = tf_c6.paragraphs[0] if i == 0 else tf_c6.add_paragraph()
    p.text = f"- {b_title} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_NAVY_DARK
    p.space_before = Pt(8) if i > 0 else Pt(0)
    
    run = p.add_run()
    run.text = b_desc
    run.font.bold = False
    run.font.size = Pt(12.5)
    run.font.color.rgb = C_DARK

img_p = os.path.join(FIGURES_DIR, "05_depth_curve.png")
if os.path.exists(img_p):
    s6.shapes.add_picture(img_p, Inches(6.75), Inches(1.55), width=Inches(5.78))
    tb_cap = s6.shapes.add_textbox(Inches(6.75), Inches(6.35), Inches(5.78), Inches(0.35))
    tb_cap.text_frame.paragraphs[0].text = "Hình 3: Đường cong hiệu năng chứng minh hiện tượng Overfitting khi tăng max_depth"
    tb_cap.text_frame.paragraphs[0].font.size = Pt(9.5)
    tb_cap.text_frame.paragraphs[0].font.italic = True
    tb_cap.text_frame.paragraphs[0].font.color.rgb = C_MUTED
    tb_cap.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

set_speaker_notes(s6, """
Ở Thí nghiệm 1, khi tăng độ sâu cây, điểm số trên tập Train nhanh chóng chạm mốc 100% học vẹt toàn bộ dữ liệu, nhưng trên tập Validation lại giảm dần từ 93,7% xuống 89,1%. Sự phân kỳ rõ rệt giữa hai đường này chứng minh cây quyết định tự do có phương sai rất lớn, khẳng định bắt buộc phải cắt tỉa mô hình.
""")


# ============================================================
# SLIDE 7: THÍ NGHIỆM 2 - CẮT TỈA CCP_ALPHA
# ============================================================
s7 = prs.slides.add_slide(blank_layout)
add_slide_header(s7, "Cắt Tỉa ccp_alpha Giảm 97,9% Kích Thước Nhưng Giữ Vững Hiệu Năng", "THÍ NGHIỆM 2: CẮT TỈA CHI PHÍ ĐỘ PHỨC TẠP (COST-COMPLEXITY PRUNING)")
add_slide_footer(s7, 7)

# Left info
add_card(s7, Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.15))
tb_c7 = s7.shapes.add_textbox(Inches(1.0), Inches(1.75), Inches(5.3), Inches(4.75))
tf_c7 = tb_c7.text_frame
tf_c7.word_wrap = True

bullets_s7 = [
    ("Hàm chi phí độ phức tạp:", "R_alpha(T) = R(T) + alpha * |T| phạt số lượng nút lá |T|."),
    ("Tối ưu bằng 5-Fold CV:", "Tìm được alpha* = 0.008024 với CV Accuracy đạt 92,48% ± 1,59%."),
    ("Đột phá về cấu trúc:", "Cắt bỏ 97,9% nút lá: từ 187 lá cồng kềnh xuống chỉ còn 4 lá."),
    ("Hiệu năng thực tế:", "Validation Accuracy vẫn đạt 93,70% và F1-Score đạt 92,58%."),
    ("Khả năng giải thích:", "Chỉ cần 2 biến Perimeter và Eccentricity là giải thích trọn vẹn luật.")
]
for i, (b_title, b_desc) in enumerate(bullets_s7):
    p = tf_c7.paragraphs[0] if i == 0 else tf_c7.add_paragraph()
    p.text = f"- {b_title} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_NAVY_DARK
    p.space_before = Pt(8) if i > 0 else Pt(0)
    
    run = p.add_run()
    run.text = b_desc
    run.font.bold = False
    run.font.size = Pt(12.5)
    run.font.color.rgb = C_DARK

# Right side: Decision tree diagram
img_p = os.path.join(FIGURES_DIR, "07_decision_tree.png")
if os.path.exists(img_p):
    s7.shapes.add_picture(img_p, Inches(6.75), Inches(1.55), width=Inches(5.78))
    tb_cap = s7.shapes.add_textbox(Inches(6.75), Inches(6.35), Inches(5.78), Inches(0.35))
    tb_cap.text_frame.paragraphs[0].text = "Hình 4: Cấu trúc Cây quyết định tối giản chỉ gồm 4 nút lá sau khi cắt tỉa ccp_alpha"
    tb_cap.text_frame.paragraphs[0].font.size = Pt(9.5)
    tb_cap.text_frame.paragraphs[0].font.italic = True
    tb_cap.text_frame.paragraphs[0].font.color.rgb = C_MUTED
    tb_cap.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

set_speaker_notes(s7, """
Tại Thí nghiệm 2, bằng kỹ thuật Cost-Complexity Pruning qua 5-Fold Cross Validation, nhóm tìm ra alpha tối ưu bằng 0.008024. Cây ban đầu cồng kềnh với 187 lá đã được rút gọn triệt để chỉ còn 4 lá. Mô hình chỉ cần dựa vào chu vi và độ lệch tâm mà vẫn giữ nguyên độ chính xác 93,7%, mang lại khả năng giải thích tuyệt đối cho người vận hành.
""")


# ============================================================
# SLIDE 8: THÍ NGHIỆM 3 - SO SÁNH CÂY ĐƠN VS RỪNG NGẪU NHIÊN
# ============================================================
s8 = prs.slides.add_slide(blank_layout)
add_slide_header(s8, "Cây Tối Giản Vượt Trội Về Minh Bạch, Rừng Ngẫu Nhiên Tối Ưu ROC-AUC", "THÍ NGHIỆM 3: SO SÁNH ĐỐI CHỨNG CÂY ĐƠN VS RỪNG NGẪU NHIÊN")
add_slide_footer(s8, 8)

# Left Column (Table)
add_card(s8, Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.15))
tb_c8 = s8.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.3), Inches(4.85))
tf_c8 = tb_c8.text_frame
tf_c8.word_wrap = True

p = tf_c8.paragraphs[0]
p.text = "BẢNG ĐỐI CHIẾU HIỆU SUẤT TRÊN VALIDATION:"
p.font.size = Pt(12)
p.font.bold = True
p.font.color.rgb = C_BLUE_ACCENT

# Comparison rows
comp_rows = [
    ("Dummy Baseline:", "Acc: 57,22% | F1: 0,00% | ROC-AUC: 0.5000"),
    ("Decision Tree (Chưa tỉa):", "Acc: 89,24% | F1: 87,41% | ROC-AUC: 0.8901"),
    ("Decision Tree (Đã tỉa):", "Acc: 93,70% | F1: 92,58% | ROC-AUC: 0.9344"),
    ("Random Forest (100 cây):", "Acc: 93,44% | F1: 92,24% | ROC-AUC: 0.9709"),
    ("Thời gian huấn luyện:", "Cây đơn: 0.036s | Rừng ngẫu nhiên: 0.320s (gấp 9 lần)"),
    ("Khả năng giải thích:", "Cây đơn: Minh bạch 100% | Rừng ngẫu nhiên: Hộp đen")
]
for title, val in comp_rows:
    p = tf_c8.add_paragraph()
    p.text = f"- {title} "
    p.font.bold = True
    p.font.size = Pt(11.5)
    p.font.color.rgb = C_NAVY_DARK
    p.space_before = Pt(5)
    run = p.add_run()
    run.text = val
    run.font.bold = False
    run.font.size = Pt(11)
    run.font.color.rgb = C_DARK

# Right side: ROC curves
img_p = os.path.join(FIGURES_DIR, "11_roc_curves.png")
if os.path.exists(img_p):
    s8.shapes.add_picture(img_p, Inches(6.75), Inches(1.55), width=Inches(5.78))
    tb_cap = s8.shapes.add_textbox(Inches(6.75), Inches(6.35), Inches(5.78), Inches(0.35))
    tb_cap.text_frame.paragraphs[0].text = "Hình 5: Đường cong ROC so sánh năng lực phân tách giữa Cây đơn và Rừng ngẫu nhiên"
    tb_cap.text_frame.paragraphs[0].font.size = Pt(9.5)
    tb_cap.text_frame.paragraphs[0].font.italic = True
    tb_cap.text_frame.paragraphs[0].font.color.rgb = C_MUTED
    tb_cap.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

set_speaker_notes(s8, """
So sánh đối chứng ở Thí nghiệm 3 cho thấy: Cây quyết định cắt tỉa đạt độ chính xác tương đương Rừng ngẫu nhiên (93,7% so với 93,4%). Tuy nhiên, Rừng ngẫu nhiên vượt trội ở ROC-AUC đạt 0.9709 nhờ đường cong xác suất liên tục. Do đó, Cây đơn phù hợp cho khâu kiểm toán nghiệp vụ, còn Rừng ngẫu nhiên tối ưu để điều khiển van khí nén tự động.
""")


# ============================================================
# SLIDE 9: THÍ NGHIỆM 4 & NGHIỆM THU TEST SET
# ============================================================
s9 = prs.slides.add_slide(blank_layout)
add_slide_header(s9, "Nghiệm Thu Tập Test Đạt 91,60% Và Bản Chất Sai Số Do Ranh Giới Tự Nhiên", "NGHIỆM THU TEST SET & PHÂN TÍCH LỖI (ERROR ANALYSIS)")
add_slide_footer(s9, 9)

add_card(s9, Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.15))
tb_c9 = s9.shapes.add_textbox(Inches(1.0), Inches(1.75), Inches(5.3), Inches(4.75))
tf_c9 = tb_c9.text_frame
tf_c9.word_wrap = True

bullets_s9 = [
    ("Nghiệm thu tập Test (1 lần):", "Random Forest đạt Test Accuracy 91,60%, ROC-AUC 0.9709."),
    ("Khả năng tổng quát hóa:", "Độ suy giảm giữa Val và Test chỉ 1,84%, không có rò rỉ dữ liệu."),
    ("Phân tích 50 ca đoán sai:", "Rơi vào vùng giao thoa hình thái tự nhiên (Major_Axis 210-230 px)."),
    ("Bản chất sinh học:", "Hạt Osmancik to đột biến đan xen hạt Cammeo bị còi cọc nhỏ."),
    ("Năng lực tự nhận thức:", "Độ tự tin mẫu đúng đạt 94,37%, mẫu đoán sai sụt giảm còn 78,28%.")
]
for i, (b_title, b_desc) in enumerate(bullets_s9):
    p = tf_c9.paragraphs[0] if i == 0 else tf_c9.add_paragraph()
    p.text = f"- {b_title} "
    p.font.bold = True
    p.font.size = Pt(13)
    p.font.color.rgb = C_NAVY_DARK
    p.space_before = Pt(8) if i > 0 else Pt(0)
    
    run = p.add_run()
    run.text = b_desc
    run.font.bold = False
    run.font.size = Pt(12.5)
    run.font.color.rgb = C_DARK

img_p = os.path.join(FIGURES_DIR, "12_error_analysis.png")
if os.path.exists(img_p):
    s9.shapes.add_picture(img_p, Inches(6.75), Inches(1.55), width=Inches(5.78))
    tb_cap = s9.shapes.add_textbox(Inches(6.75), Inches(6.35), Inches(5.78), Inches(0.35))
    tb_cap.text_frame.paragraphs[0].text = "Hình 6: Phân tích vùng đặc trưng sai số chỉ ra vùng giao thoa ranh giới tự nhiên"
    tb_cap.text_frame.paragraphs[0].font.size = Pt(9.5)
    tb_cap.text_frame.paragraphs[0].font.italic = True
    tb_cap.text_frame.paragraphs[0].font.color.rgb = C_MUTED
    tb_cap.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER

set_speaker_notes(s9, """
Trên tập Test độc lập chạy đúng một lần duy nhất, mô hình đạt độ chính xác 91,60%, độ lệch so với Val chưa đầy 2%. Khi mổ xẻ 50 ca đoán sai, nhóm nhận thấy các hạt này đều nằm ở vùng ranh giới tự nhiên, nơi hạt Osmancik to bất thường gặp hạt Cammeo còi cọc. Đặc biệt, mô hình thể hiện năng lực tự nhận thức rất tốt: khi đoán sai, độ tự tin tự động sụt giảm từ 94% xuống 78%.
""")


# ============================================================
# SLIDE 10: ỨNG DỤNG WEB THỰC TẾ
# ============================================================
s10 = prs.slides.add_slide(blank_layout)
add_slide_header(s10, "Đóng Gói Web App Flask Với Kiểm Soát 3 Tầng Và Nhận Diện Bất Định", "TRIỂN KHAI ỨNG DỤNG & PHỤC VỤ DÂY CHUYỀN")
add_slide_footer(s10, 10)

cards_s10 = [
    (Inches(0.8), Inches(1.55), Inches(3.64), Inches(5.15), "1. TÁCH BIỆT SERVING OFFLINE", [
        ("- Nạp mô hình tĩnh:", "final_model.pkl được load vào RAM 1 lần khi khởi động."),
        ("- Suy luận tức thì:", "Thời gian dự đoán < 1 mili-giây, không train lại."),
        ("- Giao diện tiếng Việt:", "7 chỉ số hiển thị rõ ràng kèm đơn vị pixel, pixel²."),
        ("- Điền nhanh mẫu thử:", "Hỗ trợ 2 nút bấm mẫu Cammeo và Osmancik hợp lệ.")
    ]),
    (Inches(4.84), Inches(1.55), Inches(3.64), Inches(5.15), "2. KIỂM TRA ĐẦU VÀO 3 TẦNG", [
        ("- Tầng 1 (Thiếu trường):", "Chặn đứng các trường rỗng, trả lỗi HTTP 400."),
        ("- Tầng 2 (Sai kiểu số):", "Bắt lỗi ký tự lạ, chữ cái hoặc giá trị không hợp lệ."),
        ("- Tầng 3 (Ngoài miền):", "Phát cảnh báo vàng nếu dữ liệu vượt khoảng [min, max]."),
        ("- Trách nhiệm dữ liệu:", "Không âm thầm thay thế dữ liệu ngoại lai.")
    ]),
    (Inches(8.88), Inches(1.55), Inches(3.64), Inches(5.15), "3. TÍNH NĂNG ĐẠT ĐIỂM THƯỞNG", [
        ("- Cảnh báo bất định:", "Xác suất trong dải [0.40, 0.60] tự động báo cờ đỏ."),
        ("- Chuyển soi kính:", "Yêu cầu chuyển chuyên viên soi mắt, không đóng gói nhầm."),
        ("- Xử lý CSV theo lô:", "Upload file hàng trăm mẫu, xuất kết quả kèm cờ cảnh báo."),
        ("- Model Card:", "Công bố minh bạch phiên bản scikit-learn và giới hạn đạo đức.")
    ])
]

for left, top, w, h, title, items in cards_s10:
    add_card(s10, left, top, w, h, bg_color=C_WHITE, border_color=C_CARD_BORDER)
    tb = s10.shapes.add_textbox(left + Inches(0.2), top + Inches(0.2), w - Inches(0.4), h - Inches(0.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = C_BLUE_ACCENT
    for b_title, b_desc in items:
        p_item = tf.add_paragraph()
        p_item.text = f"{b_title} "
        p_item.font.bold = True
        p_item.font.size = Pt(11.5)
        p_item.font.color.rgb = C_NAVY_DARK
        p_item.space_before = Pt(6)
        
        run = p_item.add_run()
        run.text = b_desc
        run.font.bold = False
        run.font.size = Pt(11)
        run.font.color.rgb = C_DARK

set_speaker_notes(s10, """
Nhóm đã đóng gói mô hình thành Web App Flask chạy cục bộ với cơ chế kiểm soát đầu vào 3 tầng chặt chẽ. Hệ thống tích hợp tính năng cảnh báo bất định: khi xác suất phân loại rơi vào dải phân vân 40 đến 60%, phần mềm sẽ cảnh báo yêu cầu chuyên viên soi kính hiển vi thủ công, đồng thời hỗ trợ tải lên file CSV để phân loại hàng nghìn hạt cùng lúc.
""")


# ============================================================
# SLIDE 11: THẢO LUẬN - HẠN CHẾ VÀ HƯỚNG CẢI TIẾN
# ============================================================
s11 = prs.slides.add_slide(blank_layout)
add_slide_header(s11, "Hoàn Thiện Tính Năng Thời Gian Thực Và Mở Rộng Đa Giống Lúa", "THẢO LUẬN: HẠN CHẾ & HƯỚNG PHÁT TRIỂN")
add_slide_footer(s11, 11)

cards_s11 = [
    (Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.15), "CÁC HẠN CHẾ HIỆN TẠI CỦA ĐỀ TÀI", [
        ("Giới hạn về số lượng giống lúa:", "Hiện mô hình mới chỉ phân loại được 2 giống Cammeo và Osmancik, chưa mở rộng sang các giống gạo đặc sản Việt Nam (ST24, ST25, Jasmine)."),
        ("Phụ thuộc trích xuất đặc trưng:", "Hệ thống phụ thuộc vào chất lượng module tách nền và đo đạc hình học 2D từ ảnh ban đầu."),
        ("Giới hạn đặc trưng 2D:", "Chưa khai thác được các thông tin đa chiều như kết cấu vân hạt (Texture) hay màu sắc quang phổ (Color spectrum).")
    ], C_RED_BG, C_RED),
    (Inches(6.833), Inches(1.55), Inches(5.7), Inches(5.15), "HƯỚNG PHÁT TRIỂN & CẢI TIẾN TƯƠNG LAI", [
        ("Ứng dụng Deep Learning (CNNs):", "Thử nghiệm mạng tích chập ResNet / EfficientNet phân loại trực tiếp từ ảnh chụp khay hạt gạo của camera điện thoại."),
        ("Cơ chế Active Learning:", "Tự động gom các mẫu bị gắn cờ bất định (Uncertainty Flag) chuyển chuyên gia dán nhãn lại, tái huấn luyện mô hình liên tục."),
        ("Nhúng phần cứng IoT thời gian thực:", "Tích hợp mô hình Cây tối giản vào vi điều khiển điều khiển van xịt khí nén phân loại hạt gạo trên băng chuyền.")
    ], C_GREEN_BG, C_GREEN)
]

for left, top, w, h, title, items, bg_c, border_c in cards_s11:
    add_card(s11, left, top, w, h, bg_color=C_WHITE, border_color=border_c)
    tb = s11.shapes.add_textbox(left + Inches(0.2), top + Inches(0.2), w - Inches(0.4), h - Inches(0.4))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = border_c
    for b_title, b_desc in items:
        p_item = tf.add_paragraph()
        p_item.text = f"- {b_title} "
        p_item.font.bold = True
        p_item.font.size = Pt(12)
        p_item.font.color.rgb = C_NAVY_DARK
        p_item.space_before = Pt(8)
        
        run = p_item.add_run()
        run.text = b_desc
        run.font.bold = False
        run.font.size = Pt(11.5)
        run.font.color.rgb = C_DARK

set_speaker_notes(s11, """
Về mặt hạn chế, mô hình hiện mới chỉ phân biệt được 2 giống Cammeo và Osmancik trên không gian đặc trưng 2D. Trong tương lai, nhóm định hướng mở rộng sang các giống lúa đặc sản của Việt Nam như ST25, áp dụng mạng tích chập CNN để đọc trực tiếp ảnh chụp từ camera điện thoại và xây dựng cơ chế Active Learning tự cải thiện từ các mẫu bất định.
""")


# ============================================================
# SLIDE 12: KẾT LUẬN & HỎI ĐÁP (Q&A)
# ============================================================
s12 = prs.slides.add_slide(blank_layout)
add_slide_header(s12, "Ba Bài Học Cốt Lõi Từ Quy Trình, Sự Tối Giản Và Trách Nhiệm", "KẾT LUẬN & PHIÊN HỎI ĐÁP (Q&A)")
add_slide_footer(s12, 12)

# Top 3 Takeaways
coords_s12 = [
    (Inches(0.8), Inches(1.55), Inches(3.64), Inches(2.3), "1. QUY TRÌNH QUYẾT ĐỊNH ĐỘ TIN CẬY",
     "Chia tầng Stratified 60/20/20 và bảo vệ tập Test độc lập có giá trị khoa học vượt trội so với điểm số ảo bị rò rỉ dữ liệu."),
    (Inches(4.84), Inches(1.55), Inches(3.64), Inches(2.3), "2. ĐƠN GIẢN LÀ ĐỈNH CAO",
     "Cây 4 lá sau cắt tỉa ccp_alpha giải thích minh bạch 100% cho công nhân mà vẫn đạt độ chính xác tương đương Rừng 100 cây."),
    (Inches(8.88), Inches(1.55), Inches(3.64), Inches(2.3), "3. TRÁCH NHIỆM CÔNG NGHỆ",
     "Một hệ thống AI thực tế phải biết nghi ngờ: tính năng cảnh báo bất định dải 40–60% bảo đảm an toàn tuyệt đối cho nông sản.")
]

for left, top, w, h, title, text in coords_s12:
    add_card(s12, left, top, w, h, bg_color=C_WHITE, border_color=C_BLUE_ACCENT)
    tb = s12.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), w - Inches(0.4), h - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(11.5)
    p.font.bold = True
    p.font.color.rgb = C_BLUE_ACCENT
    p_desc = tf.add_paragraph()
    p_desc.text = text
    p_desc.font.size = Pt(11.5)
    p_desc.font.color.rgb = C_DARK
    p_desc.space_before = Pt(4)

# Bottom Thank you & QA Box
add_card(s12, Inches(0.8), Inches(4.05), Inches(11.733), Inches(2.65), bg_color=C_CARD_BG, border_color=C_SECONDARY)
tb_ty = s12.shapes.add_textbox(Inches(1.0), Inches(4.2), Inches(11.333), Inches(2.35))
tf_ty = tb_ty.text_frame
tf_ty.word_wrap = True

p = tf_ty.paragraphs[0]
p.text = "XIN TRÂN TRỌNG CẢM ƠN THẦY GIÁO VÀ HỘI ĐỒNG ĐÃ LẮNG NGHE!"
p.font.size = Pt(22)
p.font.bold = True
p.font.color.rgb = C_NAVY_DARK
p.alignment = PP_ALIGN.CENTER

p = tf_ty.add_paragraph()
p.text = "Kính mong nhận được những câu hỏi và ý kiến góp ý quý báu của Thầy để hoàn thiện đề tài!"
p.font.size = Pt(14)
p.font.italic = True
p.font.color.rgb = C_DARK
p.alignment = PP_ALIGN.CENTER
p.space_before = Pt(6)

p = tf_ty.add_paragraph()
p.text = "PHIÊN BẢO VỆ & HỎI ĐÁP (Q & A)"
p.font.size = Pt(18)
p.font.bold = True
p.font.color.rgb = C_ACCENT
p.alignment = PP_ALIGN.CENTER
p.space_before = Pt(8)

p = tf_ty.add_paragraph()
p.text = "Tài nguyên đề tài: GitHub: https://github.com/hao0979/Rice_Classification | Báo cáo chi tiết: BaoCao_BTL_PhanLoaiGiongLua_final.docx"
p.font.size = Pt(11)
p.font.color.rgb = C_MUTED
p.alignment = PP_ALIGN.CENTER
p.space_before = Pt(6)

set_speaker_notes(s12, """
Tổng kết lại, dự án đem lại cho chúng em 3 bài học lớn: Quy trình chống rò rỉ quan trọng hơn điểm số; Cây tối giản chứng minh giải pháp đơn giản vẫn mang lại hiệu quả vượt trội; và một hệ thống AI có trách nhiệm phải biết cảnh báo khi không chắc chắn. Nhóm xin trân trọng cảm ơn Thầy Hậu đã tận tình hướng dẫn. Chúng em rất mong nhận được câu hỏi góp ý từ Thầy ạ!
""")

# ============================================================
# SAVE PRESENTATION
# ============================================================
saved = False
try:
    prs.save(SLIDES_PATH)
    print(f"Presentation saved successfully to: {SLIDES_PATH}")
    saved = True
except PermissionError:
    print(f"File {SLIDES_PATH} is locked, trying alt path...")

try:
    prs.save(ALT_PATH)
    print(f"Presentation also saved successfully to: {ALT_PATH}")
    saved = True
except Exception as e:
    print(f"Alt save result: {e}")

if saved:
    print("SUCCESS: 12-slide presentation with embedded speaker notes created!")
else:
    print("FAILURE to save slides.")
