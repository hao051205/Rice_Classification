"""
generate_slides.py - Tạo file presentation.pptx (17 slides) chuẩn format Lễ Bảo vệ Đồ án / Khóa luận Tốt nghiệp
Trường Đại học Sư phạm Kỹ thuật Hưng Yên - Khoa Công nghệ thông tin - Bộ môn Hệ thống thông tin
Đề tài: Phân loại hai giống lúa gạo Cammeo và Osmancik bằng Cây quyết định, Cắt tỉa và Rừng ngẫu nhiên

Đã căn chỉnh:
1. Toàn bộ các dấu chấm tròn (•) chuyển thành dấu gạch đầu dòng (-) chuẩn mực.
2. Phóng to cỡ chữ toàn diện (Tiêu đề 24-28pt, Nội dung chính 13-16pt, Bảng 12.5-13pt) giúp nhìn rõ từ xa khi trình chiếu.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SLIDES_PATH = os.path.join(BASE_DIR, "slides", "presentation.pptx")
FIGURES_DIR = os.path.join(BASE_DIR, "report", "figures")

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

blank_layout = prs.slide_layouts[6]

# Palettes
C_NAVY_DARK = RGBColor(15, 34, 64)
C_PRIMARY = RGBColor(20, 52, 98)
C_SECONDARY = RGBColor(37, 99, 165)
C_ACCENT = RGBColor(220, 80, 50)
C_DARK = RGBColor(33, 37, 41)
C_MUTED = RGBColor(100, 116, 139)
C_CARD_BG = RGBColor(248, 250, 252)
C_CARD_BORDER = RGBColor(226, 232, 240)
C_WHITE = RGBColor(255, 255, 255)
C_GREEN = RGBColor(22, 101, 52)
C_GREEN_BG = RGBColor(240, 253, 244)
C_AMBER = RGBColor(180, 83, 9)
C_AMBER_BG = RGBColor(254, 243, 199)


def add_slide_header(slide, title_text, section_text="BÁO CÁO BẢO VỆ ĐỒ ÁN HỌC MÁY — PROJECT 17"):
    """Thêm thanh header chuẩn học thuật với cỡ chữ to rõ và đường kẻ phân cách."""
    # Top breadcrumb badge
    tb_sec = slide.shapes.add_textbox(Inches(0.8), Inches(0.32), Inches(11.733), Inches(0.35))
    tf_sec = tb_sec.text_frame
    tf_sec.word_wrap = True
    tf_sec.margin_left = tf_sec.margin_top = tf_sec.margin_right = tf_sec.margin_bottom = 0
    p_sec = tf_sec.paragraphs[0]
    p_sec.text = section_text.upper()
    p_sec.font.size = Pt(12)
    p_sec.font.bold = True
    p_sec.font.color.rgb = C_SECONDARY

    # Main Title (Phóng to 24pt)
    tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.65), Inches(11.733), Inches(0.7))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    tf_title.margin_left = tf_title.margin_top = tf_title.margin_right = tf_title.margin_bottom = 0
    p_title = tf_title.paragraphs[0]
    p_title.text = title_text
    p_title.font.size = Pt(24)
    p_title.font.bold = True
    p_title.font.color.rgb = C_NAVY_DARK

    # Accent line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.38), Inches(11.733), Inches(0.025))
    line.fill.solid()
    line.fill.fore_color.rgb = C_SECONDARY
    line.line.color.rgb = C_SECONDARY


def add_slide_footer(slide, current_page, total_pages=17):
    """Thêm footer chuẩn với cỡ chữ to rõ, hiển thị thông tin trường và số trang."""
    # Divider line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(6.92), Inches(11.733), Inches(0.015))
    line.fill.solid()
    line.fill.fore_color.rgb = C_CARD_BORDER
    line.line.color.rgb = C_CARD_BORDER

    # Footer left
    tb_l = slide.shapes.add_textbox(Inches(0.8), Inches(6.98), Inches(5.5), Inches(0.35))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True
    tf_l.margin_left = tf_l.margin_top = tf_l.margin_right = tf_l.margin_bottom = 0
    p_l = tf_l.paragraphs[0]
    p_l.text = "Bảo vệ Đồ án Học máy — Phân loại giống lúa gạo Cammeo & Osmancik"
    p_l.font.size = Pt(10.5)
    p_l.font.color.rgb = C_MUTED

    # Footer center
    tb_c = slide.shapes.add_textbox(Inches(5.0), Inches(6.98), Inches(4.5), Inches(0.35))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    tf_c.margin_left = tf_c.margin_top = tf_c.margin_right = tf_c.margin_bottom = 0
    p_c = tf_c.paragraphs[0]
    p_c.alignment = PP_ALIGN.CENTER
    p_c.text = "Trường ĐH Sư phạm Kỹ thuật Hưng Yên — Khoa CNTT"
    p_c.font.size = Pt(10.5)
    p_c.font.color.rgb = C_MUTED

    # Footer right
    tb_r = slide.shapes.add_textbox(Inches(10.0), Inches(6.98), Inches(2.533), Inches(0.35))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    tf_r.margin_left = tf_r.margin_top = tf_r.margin_right = tf_r.margin_bottom = 0
    p_r = tf_r.paragraphs[0]
    p_r.alignment = PP_ALIGN.RIGHT
    p_r.text = f"Slide {current_page} / {total_pages}"
    p_r.font.size = Pt(11)
    p_r.font.bold = True
    p_r.font.color.rgb = C_SECONDARY


def add_card(slide, left, top, width, height, bg_color=C_CARD_BG, border_color=C_CARD_BORDER):
    """Tạo một khối container thẻ viền bo sắc nét."""
    card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    card.fill.solid()
    card.fill.fore_color.rgb = bg_color
    card.line.color.rgb = border_color
    card.line.width = Pt(1)
    return card


# ============================================================
# SLIDE 1: TRANG BÌA (OFFICIAL TITLE SLIDE)
# ============================================================
s1 = prs.slides.add_slide(blank_layout)

# Top Institutional Header
tb_inst = s1.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.95))
tf_inst = tb_inst.text_frame
tf_inst.word_wrap = True
tf_inst.margin_left = tf_inst.margin_top = tf_inst.margin_right = tf_inst.margin_bottom = 0

p = tf_inst.paragraphs[0]
p.text = "BỘ GIÁO DỤC VÀ ĐÀO TẠO — TRƯỜNG ĐẠI HỌC SƯ PHẠM KỸ THUẬT HƯNG YÊN"
p.font.size = Pt(14.5)
p.font.bold = True
p.font.color.rgb = C_PRIMARY
p.alignment = PP_ALIGN.CENTER

p = tf_inst.add_paragraph()
p.text = "KHOA CÔNG NGHỆ THÔNG TIN — BỘ MÔN HỆ THỐNG THÔNG TIN"
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = C_SECONDARY
p.alignment = PP_ALIGN.CENTER
p.space_before = Pt(3)

# Center Project Card
c_main = add_card(s1, Inches(0.8), Inches(1.45), Inches(11.733), Inches(3.25), bg_color=C_CARD_BG, border_color=C_SECONDARY)

tb_title = s1.shapes.add_textbox(Inches(1.1), Inches(1.6), Inches(11.133), Inches(2.9))
tf_title = tb_title.text_frame
tf_title.word_wrap = True

p = tf_title.paragraphs[0]
p.text = "BÁO CÁO BẢO VỆ ĐỒ ÁN HỌC MÁY CƠ BẢN (PROJECT 17)"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = C_ACCENT
p.alignment = PP_ALIGN.CENTER

p = tf_title.add_paragraph()
p.text = "PHÂN LOẠI HAI GIỐNG LÚA GẠO CAMMEO VÀ OSMANCIK\nBẰNG CÂY QUYẾT ĐỊNH VÀ RỪNG NGẪU NHIÊN"
p.font.size = Pt(27)
p.font.bold = True
p.font.color.rgb = C_NAVY_DARK
p.alignment = PP_ALIGN.CENTER
p.space_before = Pt(8)

p = tf_title.add_paragraph()
p.text = "Classification of Rice Varieties Using Decision Tree, Cost-Complexity Pruning and Random Forest"
p.font.size = Pt(14)
p.font.italic = True
p.font.color.rgb = C_MUTED
p.alignment = PP_ALIGN.CENTER
p.space_before = Pt(8)

# Supervisor Card (Left)
add_card(s1, Inches(0.8), Inches(4.85), Inches(5.7), Inches(1.8), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_sup = s1.shapes.add_textbox(Inches(1.0), Inches(5.0), Inches(5.3), Inches(1.5))
tf_sup = tb_sup.text_frame
tf_sup.word_wrap = True

p = tf_sup.paragraphs[0]
p.text = "GIẢNG VIÊN HƯỚNG DẪN:"
p.font.size = Pt(12)
p.font.bold = True
p.font.color.rgb = C_SECONDARY

p = tf_sup.add_paragraph()
p.text = "PGS. TS. Nguyễn Văn Hậu"
p.font.size = Pt(18)
p.font.bold = True
p.font.color.rgb = C_NAVY_DARK
p.space_before = Pt(4)

p = tf_sup.add_paragraph()
p.text = "Khoa Công nghệ thông tin — Trường ĐH SPKT Hưng Yên"
p.font.size = Pt(12.5)
p.font.color.rgb = C_DARK
p.space_before = Pt(3)

# Students Card (Right)
add_card(s1, Inches(6.833), Inches(4.85), Inches(5.7), Inches(1.8), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_stu = s1.shapes.add_textbox(Inches(7.033), Inches(5.0), Inches(5.3), Inches(1.5))
tf_stu = tb_stu.text_frame
tf_stu.word_wrap = True

p = tf_stu.paragraphs[0]
p.text = "SINH VIÊN THỰC HIỆN:"
p.font.size = Pt(12)
p.font.bold = True
p.font.color.rgb = C_SECONDARY

p = tf_stu.add_paragraph()
p.text = "1. Chử Quốc Hào — MSV: 10123112 (Trưởng nhóm)\n2. [Thành viên 2] — MSV: [MSV2]"
p.font.size = Pt(14)
p.font.bold = True
p.font.color.rgb = C_NAVY_DARK
p.space_before = Pt(4)

p = tf_stu.add_paragraph()
p.text = "Lớp: 10123O.1 — Ngành Công nghệ thông tin"
p.font.size = Pt(12.5)
p.font.color.rgb = C_DARK
p.space_before = Pt(3)

# Bottom Location & Date
tb_date = s1.shapes.add_textbox(Inches(0.8), Inches(6.78), Inches(11.733), Inches(0.4))
tf_date = tb_date.text_frame
p = tf_date.paragraphs[0]
p.text = "Hưng Yên — Tháng 10 Năm 2026"
p.font.size = Pt(12)
p.font.bold = True
p.font.color.rgb = C_MUTED
p.alignment = PP_ALIGN.CENTER


# ============================================================
# SLIDE 2: NỘI DUNG BÁO CÁO (AGENDA)
# ============================================================
s2 = prs.slides.add_slide(blank_layout)
add_slide_header(s2, "Nội Dung Báo Cáo Bảo Vệ Đồ Án", "TỔNG QUAN HỌC THUẬT")
add_slide_footer(s2, 2)

agenda_items = [
    ("01", "Đặt Vấn Đề & Mục Tiêu Nghiên Cứu", "Bối cảnh thực tiễn sản xuất lúa gạo, tính cấp thiết và 3 câu hỏi nghiên cứu cốt lõi."),
    ("02", "Cơ Sở Dữ Liệu & Quy Trình Chống Rò Rỉ", "Khảo sát 3.810 mẫu UCI Rice, 7 đặc trưng hình thái, Stratified Split 60/20/20 và Baseline."),
    ("03", "Phương Pháp & Thuật Toán Đề Xuất", "Cơ chế CART, Gini Impurity, Minimal Cost-Complexity Pruning và Rừng ngẫu nhiên Bagging."),
    ("04", "Kết Quả Thực Nghiệm & Phân Tích Chuyên Sâu", "4 Thí nghiệm: Overfitting max_depth, Cắt tỉa ccp_alpha, So sánh đối kháng và Phân tích lỗi."),
    ("05", "Hiện Thực Hóa Ứng Dụng Web & An Toàn AI", "Kiến trúc Flask REST API, kiểm tra dữ liệu 3 tầng, module cảnh báo bất định và Model Card."),
    ("06", "Kết Luận, Đóng Góp & Hướng Phát Triển", "Tổng kết đóng góp đề tài, phân chia nhiệm vụ 50-50 minh bạch và mở rộng Computer Vision.")
]

for idx, (num, title, desc) in enumerate(agenda_items):
    col = idx % 2
    row = idx // 2
    x = Inches(0.8 + col * 5.95)
    y = Inches(1.58 + row * 1.7)
    
    add_card(s2, x, y, Inches(5.75), Inches(1.55), bg_color=C_WHITE, border_color=C_CARD_BORDER)
    
    # Badge number
    b_num = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, x + Inches(0.2), y + Inches(0.25), Inches(0.95), Inches(0.95))
    b_num.fill.solid()
    b_num.fill.fore_color.rgb = C_PRIMARY
    b_num.line.color.rgb = C_PRIMARY
    tf_b = b_num.text_frame
    tf_b.margin_top = Inches(0.18)
    p_b = tf_b.paragraphs[0]
    p_b.text = num
    p_b.font.size = Pt(22)
    p_b.font.bold = True
    p_b.font.color.rgb = C_WHITE
    p_b.alignment = PP_ALIGN.CENTER
    
    # Content (Phóng to 15pt & 12.5pt)
    tb_c = s2.shapes.add_textbox(x + Inches(1.3), y + Inches(0.18), Inches(4.25), Inches(1.2))
    tf_c = tb_c.text_frame
    tf_c.word_wrap = True
    tf_c.margin_left = tf_c.margin_top = tf_c.margin_right = tf_c.margin_bottom = 0
    
    p = tf_c.paragraphs[0]
    p.text = title
    p.font.size = Pt(15)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_DARK
    
    p = tf_c.add_paragraph()
    p.text = desc
    p.font.size = Pt(12.5)
    p.font.color.rgb = C_DARK
    p.space_before = Pt(4)


# ============================================================
# SLIDE 3: PHẦN 1 - ĐẶT VẤN ĐỀ & CÂU HỎI NGHIÊN CỨU
# ============================================================
s3 = prs.slides.add_slide(blank_layout)
add_slide_header(s3, "Bối Cảnh Thực Tiễn & Câu Hỏi Nghiên Cứu", "PHẦN 1: ĐẶT VẤN ĐỀ")
add_slide_footer(s3, 3)

# Left Column: Thực trạng & Nhu cầu
add_card(s3, Inches(0.8), Inches(1.55), Inches(5.7), Inches(4.3), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_l3 = s3.shapes.add_textbox(Inches(1.05), Inches(1.7), Inches(5.2), Inches(4.0))
tf_l3 = tb_l3.text_frame
tf_l3.word_wrap = True

p = tf_l3.paragraphs[0]
p.text = "1. Bối Cảnh & Thực Trạng Ngành Nông Sản:"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s3_1 = [
    ("Giá trị thương phẩm:", " Cammeo và Osmancik là 2 giống lúa chất lượng cao, giá thành và đặc tính khác nhau, rất dễ bị trà trộn."),
    ("Hạn chế thủ công:", " Kiểm định bằng mắt thường gây mệt mỏi thị giác, tốc độ chậm, phụ thuộc chủ quan, độ chính xác không đều."),
    ("Nhu cầu công nghệ:", " Cần một giải pháp AI tự động nhận dạng nhanh dựa trên 7 đặc trưng hình học 2D trích xuất từ ảnh chụp."),
    ("Mục tiêu ứng dụng:", " Hỗ trợ con người sàng lọc sơ bộ hiệu quả tại các nhà máy chế biến nông sản xuất khẩu.")
]
for title, content in items_s3_1:
    p = tf_l3.add_paragraph()
    p.text = f"- {title}{content}"
    p.font.size = Pt(13)
    p.space_before = Pt(6)

# Right Column: Câu hỏi nghiên cứu
add_card(s3, Inches(6.833), Inches(1.55), Inches(5.7), Inches(4.3), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_r3 = s3.shapes.add_textbox(Inches(7.083), Inches(1.7), Inches(5.2), Inches(4.0))
tf_r3 = tb_r3.text_frame
tf_r3.word_wrap = True

p = tf_r3.paragraphs[0]
p.text = "2. Ba Câu Hỏi Nghiên Cứu Trọng Tâm:"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s3_2 = [
    ("❓ Q1 (Cân bằng Hiệu năng - Độ phức tạp):", " Cây quyết định cắt tỉa tối giản (pruned) có đạt hiệu năng tương đương Rừng ngẫu nhiên (Random Forest) hay không?"),
    ("❓ Q2 (Tính diễn giải minh bạch):", " Mô hình có thể diễn giải bằng quy tắc logic if-else rõ ràng để chuyển giao cho chuyên gia nông nghiệp không chuyên AI?"),
    ("❓ Q3 (Độ an toàn & Xử lý bất định):", " Làm thế nào để hệ thống tự phát hiện và cảnh báo các hạt gạo nằm ở ranh giới giao thoa tự nhiên (Uncertainty) để tránh đóng gói sai sót?")
]
for title, content in items_s3_2:
    p = tf_r3.add_paragraph()
    p.text = f"{title}\n- {content}"
    p.font.size = Pt(13)
    p.space_before = Pt(7)

# Bottom KPI Card: Tiêu chí thành công
add_card(s3, Inches(0.8), Inches(6.0), Inches(11.733), Inches(0.75), bg_color=C_CARD_BG, border_color=C_SECONDARY)
tb_kpi = s3.shapes.add_textbox(Inches(1.0), Inches(6.12), Inches(11.333), Inches(0.55))
tf_kpi = tb_kpi.text_frame
tf_kpi.word_wrap = True
p = tf_kpi.paragraphs[0]
p.text = "🎯 TIÊU CHÍ THÀNH CÔNG: Vượt Baseline 57.22% | Accuracy & F1 > 90% | Pipeline chuẩn 100% không rò rỉ dữ liệu | Đóng gói Web App hoàn chỉnh."
p.font.size = Pt(13.5)
p.font.bold = True
p.font.color.rgb = C_NAVY_DARK
p.alignment = PP_ALIGN.CENTER


# ============================================================
# SLIDE 4: PHẦN 2 - DỮ LIỆU THỰC NGHIỆM (UCI RICE)
# ============================================================
s4 = prs.slides.add_slide(blank_layout)
add_slide_header(s4, "Khảo Sát Bộ Dữ Liệu Thực Nghiệm (UCI Rice)", "PHẦN 2: DỮ LIỆU & TIỀN XỬ LÝ")
add_slide_footer(s4, 4)

# Left Column: Data Description
add_card(s4, Inches(0.8), Inches(1.55), Inches(5.6), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_l4 = s4.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.2), Inches(4.9))
tf_l4 = tb_l4.text_frame
tf_l4.word_wrap = True

p = tf_l4.paragraphs[0]
p.text = "1. Xuất Xứ & Phân Bổ Nhãn Dữ Liệu:"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s4_data = [
    "- Nguồn: UCI Machine Learning Repository (Cinar & Koklu, 2019).",
    "- Giấy phép: CC BY 4.0 — Sử dụng nghiên cứu & học thuật.",
    "- Quy mô: 3.810 mẫu hạt gạo bóc vỏ.",
    "- Osmancik: 2.180 mẫu (57.22%) — Đa số, hạt bầu ngắn.",
    "- Cammeo: 1.630 mẫu (42.78%) — Thiểu số, hạt dài thon.",
    "- Kiểm định chất lượng: 0 missing value, 0 duplicate."
]
for item in items_s4_data:
    p = tf_l4.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(3)

p = tf_l4.add_paragraph()
p.text = "2. 7 Đặc Trưng Hình Thái Học 2D:"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = C_PRIMARY
p.space_before = Pt(8)

feats = [
    "- Area, Perimeter: Diện tích và chu vi hạt gạo (pixel).",
    "- Major_Axis_Length: Chiều dài trục lớn elip tương đương.",
    "- Minor_Axis_Length: Chiều rộng trục phụ elip tương đương.",
    "- Eccentricity: Độ lệch tâm hình học elip [0, 1].",
    "- Convex_Area: Diện tích đa giác bao lồi nhỏ nhất của hạt.",
    "- Extent: Tỷ lệ diện tích hạt / diện tích hộp bao ngoài."
]
for f in feats:
    p = tf_l4.add_paragraph()
    p.text = f
    p.font.size = Pt(12.5)
    p.space_before = Pt(2)

# Right: Correlation Matrix Image
img_corr = os.path.join(FIGURES_DIR, "03_correlation_matrix.png")
if os.path.exists(img_corr):
    add_card(s4, Inches(6.6), Inches(1.55), Inches(5.933), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
    s4.shapes.add_picture(img_corr, Inches(6.8), Inches(1.7), width=Inches(5.533))


# ============================================================
# SLIDE 5: PHẦN 2 - PIPELINE & CHỐNG RÒ RỈ DỮ LIỆU
# ============================================================
s5 = prs.slides.add_slide(blank_layout)
add_slide_header(s5, "Thiết Kế Pipeline Phân Chia Dữ Liệu & Baseline", "PHẦN 2: DỮ LIỆU & TIỀN XỬ LÝ")
add_slide_footer(s5, 5)

# 4 Step Pipeline Diagrams
steps = [
    ("Bước 1: Nạp Dữ Liệu Gốc", "Đọc ARFF, decode UTF-8, kiểm tra kiểu dữ liệu và giá trị biên."),
    ("Bước 2: Phân Tách Stratified", "Chia tầng 60/20/20 đảm bảo tỷ lệ 57.2% Osmancik / 42.8% Cammeo đồng nhất."),
    ("Bước 3: Huấn Luyện & Tuning", "Train trên 60%, tính ccp_alpha và 5-Fold CV, đối chứng trên Val 20%."),
    ("Bước 4: Nghiệm Thu Độc Lập", "Mở niêm phong Test Set 20% duy nhất 1 lần để khẳng định tính tổng quát hóa.")
]

for idx, (st_t, st_d) in enumerate(steps):
    x = Inches(0.8 + idx * 2.98)
    add_card(s5, x, Inches(1.55), Inches(2.78), Inches(1.35), bg_color=C_CARD_BG, border_color=C_SECONDARY)
    tb = s5.shapes.add_textbox(x + Inches(0.1), Inches(1.65), Inches(2.58), Inches(1.15))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = st_t
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_PRIMARY
    p = tf.add_paragraph()
    p.text = st_d
    p.font.size = Pt(11)
    p.font.color.rgb = C_DARK
    p.space_before = Pt(3)

# 3 Columns for Train / Val / Test
splits = [
    ("TẬP HUẤN LUYỆN (TRAIN SET)", "2.286 mẫu (60.0%)", C_PRIMARY, [
        "- Huấn luyện cây quyết định và rừng ngẫu nhiên.",
        "- Trích xuất chuỗi 70 giá trị ccp_alpha khả dĩ.",
        "- Thực hiện 5-Fold Stratified Cross-Validation.",
        "- Cố định random_state = 42 đảm bảo tái lập 100%."
    ]),
    ("TẬP PHÁT TRIỂN (VALIDATION SET)", "762 mẫu (20.0%)", C_SECONDARY, [
        "- Khảo sát hiện tượng Overfitting theo max_depth (TN1).",
        "- So sánh đối kháng Cây cắt tỉa vs Random Forest (TN3).",
        "- Phân tích ma trận nhầm lẫn và lỗi sai thực tế (TN4).",
        "- Quyết định lựa chọn mô hình triển khai Web."
    ]),
    ("TẬP KIỂM THỬ (TEST SET - NIÊM PHONG)", "762 mẫu (20.0%)", C_ACCENT, [
        "- NIÊM PHONG TUYỆT ĐỐI suốt quá trình nghiên cứu.",
        "- Không can thiệp vào quyết định chọn tham số.",
        "- Chống rò rỉ dữ liệu (Data Leakage) nghiêm ngặt.",
        "- Chỉ đánh giá duy nhất 1 lần khi nghiệm thu báo cáo."
    ])
]

for idx, (sp_name, sp_size, sp_color, sp_desc) in enumerate(splits):
    x = Inches(0.8 + idx * 3.98)
    add_card(s5, x, Inches(3.05), Inches(3.78), Inches(2.75), bg_color=C_WHITE, border_color=sp_color)
    tb = s5.shapes.add_textbox(x + Inches(0.15), Inches(3.18), Inches(3.48), Inches(2.5))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = sp_name
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = sp_color
    
    p = tf.add_paragraph()
    p.text = sp_size
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_DARK
    p.space_before = Pt(2)
    
    for item in sp_desc:
        p = tf.add_paragraph()
        p.text = item
        p.font.size = Pt(12)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(4)

# Baseline Bar
add_card(s5, Inches(0.8), Inches(5.95), Inches(11.733), Inches(0.8), bg_color=C_CARD_BG, border_color=C_CARD_BORDER)
tb_base = s5.shapes.add_textbox(Inches(1.0), Inches(6.05), Inches(11.333), Inches(0.6))
tf_base = tb_base.text_frame
p = tf_base.paragraphs[0]
p.text = "📌 BASELINE ĐỐI CHỨNG (DummyClassifier - Most Frequent): Luôn dự đoán nhãn đa số (Osmancik) → Accuracy = 57.22%, F1-score = 0.00%. Mọi mô hình máy học bắt buộc phải vượt mốc này."
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = C_NAVY_DARK


# ============================================================
# SLIDE 6: PHẦN 3 - LÝ THUYẾT: CÂY QUYẾT ĐỊNH & CCP_ALPHA
# ============================================================
s6 = prs.slides.add_slide(blank_layout)
add_slide_header(s6, "Cây Quyết Định (CART) & Cắt Tỉa Cost-Complexity", "PHẦN 3: PHƯƠNG PHÁP & THUẬT TOÁN")
add_slide_footer(s6, 6)

# Left Column: CART & Gini
add_card(s6, Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_l6 = s6.shapes.add_textbox(Inches(1.05), Inches(1.7), Inches(5.2), Inches(4.9))
tf_l6 = tb_l6.text_frame
tf_l6.word_wrap = True

p = tf_l6.paragraphs[0]
p.text = "1. Thuật Toán CART & Chỉ Số Tạp Chất Gini:"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s6_cart = [
    "- Cấu trúc: Cây nhị phân phân cấp rẽ nhánh đệ quy (Recursive Binary Splitting).",
    "- Tiêu chuẩn chọn điểm cắt: Giảm độ tạp chất Gini lớn nhất:",
    "       Gini(D) = 1 - Σ [p(i)]²   (với i ∈ {Cammeo, Osmancik})",
    "- Tại mỗi nút, thuật toán quét toàn bộ 7 thuộc tính và mọi ngưỡng giá trị khả dĩ để cực đại hóa ΔGini.",
    "- Điểm mạnh: Cực kỳ trực quan, phản ánh chính xác tư duy ra quyết định của con người, không cần chuẩn hóa thang đo.",
    "- Hạn chế cốt tử: Cây tự do có xu hướng phát triển quá sâu (187 nút lá), gây hiện tượng Overfitting nghiêm trọng."
]
for item in items_s6_cart:
    p = tf_l6.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(5)

# Right Column: Minimal Cost-Complexity Pruning
add_card(s6, Inches(6.833), Inches(1.55), Inches(5.7), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_r6 = s6.shapes.add_textbox(Inches(7.083), Inches(1.7), Inches(5.2), Inches(4.9))
tf_r6 = tb_r6.text_frame
tf_r6.word_wrap = True

p = tf_r6.paragraphs[0]
p.text = "2. Cơ Chế Minimal Cost-Complexity Pruning (ccp_alpha):"
p.font.size = Pt(15.5)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s6_prun = [
    "- Ý tưởng toán học: Đưa tham số phạt kích thước vào hàm mục tiêu:",
    "       R_α(T) = R(T) + α · |T|",
    "   * R(T): Tổng sai số phân lớp trên tập huấn luyện.",
    "   * |T|: Số lượng nút lá của cây con T (thước đo độ phức tạp).",
    "   * α ≥ 0: Hệ số phạt độ phức tạp (Complexity Parameter).",
    "- Cơ chế: Khi α tăng, các nhánh mang lại ít cải thiện sai số nhất sẽ bị loại bỏ trước tiên, biến nút nhánh thành nút lá.",
    "- Ưu thế vượt trội: Tối ưu toán học toàn cục, tránh việc dừng sớm (early stopping) nông nổi của max_depth.",
    "- Tìm kiếm α tối ưu: Thông qua 5-Fold Stratified Cross Validation."
]
for item in items_s6_prun:
    p = tf_r6.add_paragraph()
    p.text = item
    p.font.size = Pt(13)
    p.space_before = Pt(5)


# ============================================================
# SLIDE 7: PHẦN 3 - LÝ THUYẾT: RỪNG NGẪU NHIÊN (RANDOM FOREST)
# ============================================================
s7 = prs.slides.add_slide(blank_layout)
add_slide_header(s7, "Rừng Ngẫu Nhiên (Random Forest Ensemble)", "PHẦN 3: PHƯƠNG PHÁP & THUẬT TOÁN")
add_slide_footer(s7, 7)

# Left Column: Bagging & Double Randomness
add_card(s7, Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_l7 = s7.shapes.add_textbox(Inches(1.05), Inches(1.7), Inches(5.2), Inches(4.9))
tf_l7 = tb_l7.text_frame
tf_l7.word_wrap = True

p = tf_l7.paragraphs[0]
p.text = "1. Kỹ Thuật Ngẫu Nhiên Hóa Kép (Double Randomness):"
p.font.size = Pt(15.5)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s7_rf = [
    ("1. Bootstrap Aggregating (Bagging):", " Rút mẫu ngẫu nhiên có hoàn lại N = 2.286 mẫu từ Train Set để xây dựng 100 tập huấn luyện riêng cho 100 cây con."),
    ("2. Random Feature Subspace:", " Tại mỗi nút phân nhánh, chỉ chọn ngẫu nhiên m = √p ≈ √7 ≈ 2-3 thuộc tính để xét điểm cắt tối ưu nhất."),
    ("Mục tiêu toán học:", " Triệt tiêu tính tương quan giữa các cây con. Giúp các đặc trưng phụ vẫn có cơ hội đóng góp vào quyết định."),
    ("Cơ chế tổng hợp:", " Phân lớp theo đa số phiếu bầu (Majority Voting) và lấy trung bình xác suất của 100 cây con.")
]
for title, content in items_s7_rf:
    p = tf_l7.add_paragraph()
    p.text = f"- {title}{content}"
    p.font.size = Pt(13)
    p.space_before = Pt(5)

# Right Column: OOB & Feature Importance
add_card(s7, Inches(6.833), Inches(1.55), Inches(5.7), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_r7 = s7.shapes.add_textbox(Inches(7.083), Inches(1.7), Inches(5.2), Inches(4.9))
tf_r7 = tb_r7.text_frame
tf_r7.word_wrap = True

p = tf_r7.paragraphs[0]
p.text = "2. Cơ Chế Đánh Giá Out-of-Bag (OOB) & Độ Tin Cậy:"
p.font.size = Pt(15.5)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s7_oob = [
    ("Đánh giá Out-of-Bag (OOB):", " Mỗi mẫu có xác suất không được chọn vào bootstrap là (1 - 1/N)^N ≈ 36.8%. Bộ mẫu này làm tập kiểm định nội tại độc lập cho từng cây mà không cần trích thêm tập validation."),
    ("Ưu thế xác suất mượt mà:", " Thay vì dự đoán nhãn cứng 0 hoặc 1 như Cây đơn, Random Forest cung cấp xác suất dự đoán mượt mà (Calibrated Probability) phục vụ phát hiện mẫu bất định."),
    ("Độ quan trọng thuộc tính (MDI):", " Đo lường mức giảm tạp chất Gini trung bình của từng thuộc tính trên toàn bộ 100 cây, phản ánh chính xác tầm ảnh hưởng sinh học của hạt lúa.")
]
for title, content in items_s7_oob:
    p = tf_r7.add_paragraph()
    p.text = f"- {title}{content}"
    p.font.size = Pt(13)
    p.space_before = Pt(6)


# ============================================================
# SLIDE 8: PHẦN 4 - THÍ NGHIỆM 1: OVERFITTING THEO MAX_DEPTH
# ============================================================
s8 = prs.slides.add_slide(blank_layout)
add_slide_header(s8, "Thí Nghiệm 1: Hiện Tượng Overfitting Theo max_depth", "PHẦN 4: KẾT QUẢ THỰC NGHIỆM")
add_slide_footer(s8, 8)

# Left Column: Findings
add_card(s8, Inches(0.8), Inches(1.55), Inches(5.6), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_l8 = s8.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.2), Inches(4.9))
tf_l8 = tb_l8.text_frame
tf_l8.word_wrap = True

p = tf_l8.paragraphs[0]
p.text = "Mục Tiêu & Diễn Biến Thực Nghiệm:"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s8 = [
    "- Thiết kế thí nghiệm: Khảo sát max_depth chạy từ 1 đến 30 trên Cây quyết định không giới hạn.",
    "- Tập Huấn luyện (Train Accuracy):",
    "   * Tăng vọt từ 92.9% (depth=1) lên 100.0% khi depth ≥ 18.",
    "   * Cây đạt độ tinh khiết tuyệt đối trên dữ liệu huấn luyện.",
    "- Tập Phát triển (Validation Accuracy):",
    "   * Đạt đỉnh 93.70% ở độ sâu rất nhỏ (depth = 2).",
    "   * Sau đó tụt dốc và dao động ở mức 89.10% - 89.24% khi cây sâu hơn.",
    "- Phát hiện cốt lõi (Overfitting):",
    "   * Khoảng cách sai số Train - Val mở rộng tới hơn 10.8%.",
    "   * Cây unconstrained tạo ra tới 187 nút lá, ghi nhớ cả nhiễu.",
    "   → Bắt buộc phải áp dụng Cắt tỉa (Pruning) để tối ưu tổng quát hóa."
]
for item in items_s8:
    p = tf_l8.add_paragraph()
    p.text = item
    p.font.size = Pt(12.5)
    p.space_before = Pt(3)

# Right: Depth curve figure
img_depth = os.path.join(FIGURES_DIR, "05_depth_curve.png")
if os.path.exists(img_depth):
    add_card(s8, Inches(6.6), Inches(1.55), Inches(5.933), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
    s8.shapes.add_picture(img_depth, Inches(6.8), Inches(1.75), width=Inches(5.533))


# ============================================================
# SLIDE 9: PHẦN 4 - THÍ NGHIỆM 2: TỐI ƯU HÓA CẮT TỈA CCP_ALPHA
# ============================================================
s9 = prs.slides.add_slide(blank_layout)
add_slide_header(s9, "Thí Nghiệm 2: Tối Ưu Hóa Cắt Tỉa Cost-Complexity", "PHẦN 4: KẾT QUẢ THỰC NGHIỆM")
add_slide_footer(s9, 9)

# Left Column: Pruning Optimization
add_card(s9, Inches(0.8), Inches(1.55), Inches(5.6), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_l9 = s9.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.2), Inches(4.9))
tf_l9 = tb_l9.text_frame
tf_l9.word_wrap = True

p = tf_l9.paragraphs[0]
p.text = "Quy Trình Khảo Sát & Kết Quả Tối Ưu:"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s9 = [
    "- Trích xuất chuỗi tham số alpha: Dùng cost_complexity_pruning_path trên cây đầy đủ, thu được 70 giá trị alpha khả dĩ.",
    "- Đánh giá khách quan bằng 5-Fold Stratified CV trên tập Train:",
    "   * Tránh rò rỉ dữ liệu, đảm bảo alpha có khả năng tổng quát hóa cao.",
    "- Giá trị alpha tối ưu tìm được:",
    "   → ccp_alpha_opt = 0.008024",
    "   → Đạt điểm 5-Fold CV Accuracy: 92.48% ± 1.59%",
    "- Hiệu quả tối giản cấu trúc cây:",
    "   * Cây đầy đủ ban đầu: 187 nút lá.",
    "   * Cây sau khi cắt tỉa: Chỉ còn đúng 4 nút lá.",
    "   → Cắt giảm 97.9% độ phức tạp tính toán!",
    "   * Validation Accuracy tăng từ 89.24% lên 93.70% (vượt cả cây gốc)."
]
for item in items_s9:
    p = tf_l9.add_paragraph()
    p.text = item
    p.font.size = Pt(12.5)
    p.space_before = Pt(3)

# Right: Pruning Curve figure
img_prune = os.path.join(FIGURES_DIR, "06_pruning_curve.png")
if os.path.exists(img_prune):
    add_card(s9, Inches(6.6), Inches(1.55), Inches(5.933), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
    s9.shapes.add_picture(img_prune, Inches(6.8), Inches(1.75), width=Inches(5.533))


# ============================================================
# SLIDE 10: PHẦN 4 - CẤU TRÚC CÂY CẮT TỈA (INTERPRETABILITY)
# ============================================================
s10 = prs.slides.add_slide(blank_layout)
add_slide_header(s10, "Cấu Trúc Cây Cắt Tỉa & Tính Diễn Giải Minh Bạch", "PHẦN 4: KẾT QUẢ THỰC NGHIỆM")
add_slide_footer(s10, 10)

# Top / Left Figure: Decision Tree
img_tree = os.path.join(FIGURES_DIR, "07_decision_tree.png")
if os.path.exists(img_tree):
    add_card(s10, Inches(0.8), Inches(1.55), Inches(7.2), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
    s10.shapes.add_picture(img_tree, Inches(0.95), Inches(1.7), width=Inches(6.9))

# Right Column: Rules & Value
add_card(s10, Inches(8.2), Inches(1.55), Inches(4.333), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_r10 = s10.shapes.add_textbox(Inches(8.35), Inches(1.7), Inches(4.033), Inches(4.9))
tf_r10 = tb_r10.text_frame
tf_r10.word_wrap = True

p = tf_r10.paragraphs[0]
p.text = "Bộ Quy Tắc Phân Quyết (If-Else):"
p.font.size = Pt(15.5)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

rules = [
    ("Nút Gốc (Root):", " Chỉ dùng Major_Axis_Length (chiều dài) ngưỡng 218.47 pixel."),
    ("Quy tắc 1 (Osmancik):", " IF Major_Axis_Length ≤ 218.47 px → Dự đoán OSMANCIK (Gini = 0.165, độ thuần khiết cao)."),
    ("Quy tắc 2 (Nhánh 2):", " IF Major_Axis_Length > 218.47 px → Kiểm tra tiếp Perimeter ngưỡng 512.65 pixel."),
    ("Quy tắc 3 (Cammeo):", " IF Perimeter > 512.65 px → Dự đoán CAMMEO (Gini = 0.089, hạt dài và chu vi lớn).")
]
for title, r in rules:
    p = tf_r10.add_paragraph()
    p.text = f"- {title}\n  {r}"
    p.font.size = Pt(12)
    p.space_before = Pt(3)

p = tf_r10.add_paragraph()
p.text = "Giá Trị Chuyển Giao Thực Tế:"
p.font.size = Pt(15.5)
p.font.bold = True
p.font.color.rgb = C_PRIMARY
p.space_before = Pt(8)

p = tf_r10.add_paragraph()
p.text = "- Mô hình đạt độ minh bạch 100%. Có thể in thành cẩm nang cầm tay cho nông dân hoặc lập trình trên vi điều khiển nhúng giá rẻ mà không cần máy tính mạnh."
p.font.size = Pt(12)
p.font.color.rgb = C_DARK
p.space_before = Pt(3)


# ============================================================
# SLIDE 11: PHẦN 4 - THÍ NGHIỆM 3: BẢNG SO SÁNH ĐỐI KHÁNG CÁC MÔ HÌNH
# ============================================================
s11 = prs.slides.add_slide(blank_layout)
add_slide_header(s11, "Thí Nghiệm 3: Bảng So Sánh Đối Đầu Giữa Các Mô Hình", "PHẦN 4: KẾT QUẢ THỰC NGHIỆM")
add_slide_footer(s11, 11)

# Left Side: Formal PowerPoint Table
table_shape = s11.shapes.add_table(8, 5, Inches(0.8), Inches(1.55), Inches(7.5), Inches(4.55))
table = table_shape.table
table.columns[0].width = Inches(2.1)
table.columns[1].width = Inches(1.25)
table.columns[2].width = Inches(1.35)
table.columns[3].width = Inches(1.4)
table.columns[4].width = Inches(1.4)

headers = ["Tiêu Chí Đánh Giá", "Baseline", "Tree (Đầy đủ)", "Tree (Pruned)", "Random Forest"]
table_data = [
    ["Val Accuracy", "57.22%", "89.24%", "93.70%", "93.44%"],
    ["Val F1-score", "0.00%", "87.26%", "92.77%", "92.28%"],
    ["Val ROC-AUC", "0.5000", "0.8901", "0.9596", "0.9741"],
    ["Test Accuracy", "57.22%", "88.58%", "91.47%", "91.60%"],
    ["Test ROC-AUC", "0.5000", "0.8841", "0.9344", "0.9709"],
    ["Số nút lá / Quy mô", "1 luật", "187 lá (rối)", "4 lá (-97.9%)", "100 cây"],
    ["Tính diễn giải", "Không", "Thấp", "Rất cao (Bạch minh)", "Hộp đen (Phức tạp)"]
]

for col_idx, h in enumerate(headers):
    cell = table.cell(0, col_idx)
    cell.text = h
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    cell.fill.solid()
    cell.fill.fore_color.rgb = C_PRIMARY
    p = cell.text_frame.paragraphs[0]
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_WHITE
    p.alignment = PP_ALIGN.CENTER

for row_idx, row in enumerate(table_data):
    for col_idx, val in enumerate(row):
        cell = table.cell(row_idx + 1, col_idx)
        cell.text = val
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        cell.fill.solid()
        cell.fill.fore_color.rgb = C_WHITE if row_idx % 2 == 0 else C_CARD_BG
        p = cell.text_frame.paragraphs[0]
        p.font.size = Pt(12.5)
        if col_idx == 0:
            p.alignment = PP_ALIGN.LEFT
            p.font.bold = True
            p.font.color.rgb = C_NAVY_DARK
        elif col_idx in [3, 4]:
            p.alignment = PP_ALIGN.CENTER
            p.font.bold = True
            p.font.color.rgb = C_PRIMARY
        else:
            p.alignment = PP_ALIGN.CENTER
            p.font.color.rgb = C_DARK

# Right: ROC Curve figure
img_roc = os.path.join(FIGURES_DIR, "11_roc_curves.png")
if os.path.exists(img_roc):
    add_card(s11, Inches(8.5), Inches(1.55), Inches(4.033), Inches(4.55), bg_color=C_WHITE, border_color=C_CARD_BORDER)
    s11.shapes.add_picture(img_roc, Inches(8.65), Inches(1.75), width=Inches(3.733))

# Bottom insight card
add_card(s11, Inches(0.8), Inches(6.18), Inches(11.733), Inches(0.68), bg_color=C_CARD_BG, border_color=C_SECONDARY)
tb_ins = s11.shapes.add_textbox(Inches(1.0), Inches(6.25), Inches(11.333), Inches(0.55))
tf_ins = tb_ins.text_frame
p = tf_ins.paragraphs[0]
p.text = "💡 KẾT LUẬN ĐỐI SÁNH: Cây cắt tỉa tối ưu tuyệt đối về tính diễn giải (Val 93.70%, Test 91.47%). Random Forest tối ưu tuyệt đối về phân tách xác suất (Val ROC-AUC 0.9741, Test ROC-AUC 0.9709) → Chọn Random Forest làm mô hình lõi cho Web App."
p.font.size = Pt(12)
p.font.bold = True
p.font.color.rgb = C_NAVY_DARK


# ============================================================
# SLIDE 12: PHẦN 4 - THÍ NGHIỆM 4: PHÂN TÍCH LỖI SAI (ERROR ANALYSIS)
# ============================================================
s12 = prs.slides.add_slide(blank_layout)
add_slide_header(s12, "Thí Nghiệm 4: Phân Tích Lỗi Sai & Cơ Sở Cảnh Báo", "PHẦN 4: KẾT QUẢ THỰC NGHIỆM")
add_slide_footer(s12, 12)

# Left Column: Error Analysis Insights
add_card(s12, Inches(0.8), Inches(1.55), Inches(5.6), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_l12 = s12.shapes.add_textbox(Inches(1.0), Inches(1.7), Inches(5.2), Inches(4.9))
tf_l12 = tb_l12.text_frame
tf_l12.word_wrap = True

p = tf_l12.paragraphs[0]
p.text = "Phát Hiện Từ Ma Trận Nhầm Lẫn (Validation Set):"
p.font.size = Pt(15.5)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s12 = [
    "- Tổng số mẫu phân loại sai: 50 / 762 mẫu (Tỷ lệ sai 6.56%).",
    "- Tính đối xứng hai chiều:",
    "   * 27 hạt Cammeo nhầm sang Osmancik (False Negative).",
    "   * 23 hạt Osmancik nhầm sang Cammeo (False Positive).",
    "- Căn nguyên sinh học (Biological Boundary Overlap):",
    "   * Mẫu sai tập trung tại vùng kích thước trung gian: Area ~ 12.773 px² và Major_Axis ~ 191.7 px.",
    "   * Đây là vùng giao thoa tự nhiên giữa hạt Cammeo nhỏ và hạt Osmancik lớn.",
    "- Phân tích độ tự tin (Confidence Score):",
    "   * Mẫu đoán đúng: Độ tự tin trung bình đạt 94.4%.",
    "   * Mẫu đoán sai: Độ tự tin tụt xuống 78.3% và nhiều mẫu tiệm cận 50% (không chắc chắn).",
    "→ Tiền đề thiết lập module phát hiện mẫu bất định [0.40, 0.60]."
]
for item in items_s12:
    p = tf_l12.add_paragraph()
    p.text = item
    p.font.size = Pt(12.5)
    p.space_before = Pt(3)

# Right: Error analysis figure
img_err = os.path.join(FIGURES_DIR, "12_error_analysis.png")
if os.path.exists(img_err):
    add_card(s12, Inches(6.6), Inches(1.55), Inches(5.933), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
    s12.shapes.add_picture(img_err, Inches(6.75), Inches(1.7), width=Inches(5.633))


# ============================================================
# SLIDE 13: PHẦN 5 - HIỆN THỰC HÓA ỨNG DỤNG WEB FLASK
# ============================================================
s13 = prs.slides.add_slide(blank_layout)
add_slide_header(s13, "Hiện Thực Hóa Hệ Thống Web Flask & Kiểm Tra 3 Tầng", "PHẦN 5: TRIỂN KHAI ỨNG DỤNG")
add_slide_footer(s13, 13)

# 3 Architecture Cards (Horizontal)
arch_cards = [
    ("TẦNG TRÌNH DIỄN (FRONTEND)", "Bootstrap 5 Responsive", C_PRIMARY, [
        "- Giao diện tiếng Việt thân thiện, tương thích mobile và desktop.",
        "- 2 Tab chức năng: Phân loại đơn hạt & Phân loại file CSV.",
        "- Trực quan hóa thanh tiến trình xác suất tin cậy (Progress Bar)."
    ]),
    ("TẦNG DỊCH VỤ (BACKEND API)", "Python Flask Server", C_SECONDARY, [
        "- Nạp sẵn final_model.pkl trong RAM, không huấn luyện lại.",
        "- Kiến trúc RESTful API xử lý JSON payload.",
        "- Tốc độ phản hồi cực nhanh: < 50ms cho mỗi yêu cầu."
    ]),
    ("TẦNG AN TOÀN (VALIDATION 3 TẦNG)", "Input Guard System", C_ACCENT, [
        "- Tầng 1: Kiểm tra cấu trúc form (bắt lỗi thiếu trường HTTP 400).",
        "- Tầng 2: Kiểm tra kiểu dữ liệu số thực dương hợp lệ.",
        "- Tầng 3: Cảnh báo vàng khi thông số nằm ngoài [min, max] Train."
    ])
]

for idx, (c_name, c_sub, c_col, c_desc) in enumerate(arch_cards):
    x = Inches(0.8 + idx * 3.98)
    add_card(s13, x, Inches(1.55), Inches(3.78), Inches(2.75), bg_color=C_WHITE, border_color=c_col)
    tb = s13.shapes.add_textbox(x + Inches(0.15), Inches(1.7), Inches(3.48), Inches(2.45))
    tf = tb.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = c_name
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = c_col
    
    p = tf.add_paragraph()
    p.text = c_sub
    p.font.size = Pt(14.5)
    p.font.bold = True
    p.font.color.rgb = C_NAVY_DARK
    p.space_before = Pt(2)
    
    for item in c_desc:
        p = tf.add_paragraph()
        p.text = item
        p.font.size = Pt(12)
        p.font.color.rgb = C_DARK
        p.space_before = Pt(4)

# Bottom 2 Feature Cards
add_card(s13, Inches(0.8), Inches(4.45), Inches(5.75), Inches(2.3), bg_color=C_CARD_BG, border_color=C_CARD_BORDER)
tb_b1 = s13.shapes.add_textbox(Inches(1.0), Inches(4.6), Inches(5.35), Inches(2.0))
tf_b1 = tb_b1.text_frame
tf_b1.word_wrap = True

p = tf_b1.paragraphs[0]
p.text = "🎯 Tính Năng Phân Loại Hạt Đơn Lẻ (Single Grain):"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s13_single = [
    "- Cho phép người dùng nhập 7 thông số hình học hạt gạo.",
    "- Tích hợp nút điền mẫu thử nhanh 1-click (Cammeo mẫu / Osmancik mẫu).",
    "- Trả về kết quả phân loại giống lúa kèm tỷ lệ xác suất tin cậy."
]
for item in items_s13_single:
    p = tf_b1.add_paragraph()
    p.text = item
    p.font.size = Pt(12.5)
    p.space_before = Pt(4)

add_card(s13, Inches(6.783), Inches(4.45), Inches(5.75), Inches(2.3), bg_color=C_CARD_BG, border_color=C_CARD_BORDER)
tb_b2 = s13.shapes.add_textbox(Inches(6.983), Inches(4.6), Inches(5.35), Inches(2.0))
tf_b2 = tb_b2.text_frame
tf_b2.word_wrap = True

p = tf_b2.paragraphs[0]
p.text = "📁 Tính Năng Phân Loại Theo Lô (Batch CSV Upload):"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s13_batch = [
    "- Cho phép tải lên tệp CSV chứa hàng trăm đến hàng nghìn mẫu hạt lúa.",
    "- Xử lý tính toán vector song song, hiển thị bảng thống kê tổng hợp.",
    "- Cho phép tải về tệp CSV kết quả có đính kèm cờ cảnh báo bất định."
]
for item in items_s13_batch:
    p = tf_b2.add_paragraph()
    p.text = item
    p.font.size = Pt(12.5)
    p.space_before = Pt(4)


# ============================================================
# SLIDE 14: PHẦN 5 - TÍNH NĂNG MỞ RỘNG: UNCERTAINTY & MODEL CARD
# ============================================================
s14 = prs.slides.add_slide(blank_layout)
add_slide_header(s14, "Cảnh Báo Bất Định & Cam Kết AI Có Trách Nhiệm", "PHẦN 5: TRIỂN KHAI ỨNG DỤNG")
add_slide_footer(s14, 14)

# Left Column: Uncertainty Detection
add_card(s14, Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.2), bg_color=C_WHITE, border_color=C_ACCENT)
tb_l14 = s14.shapes.add_textbox(Inches(1.05), Inches(1.7), Inches(5.2), Inches(4.9))
tf_l14 = tb_l14.text_frame
tf_l14.word_wrap = True

p = tf_l14.paragraphs[0]
p.text = "1. Cơ Chế Phát Hiện Mẫu Bất Định (Uncertainty):"
p.font.size = Pt(15.5)
p.font.bold = True
p.font.color.rgb = C_ACCENT

items_s14_unc = [
    ("Thiết lập ngưỡng toán học:", " Khi xác suất dự đoán của mô hình rơi vào khoảng bất định P(Cammeo) ∈ [0.40, 0.60]."),
    ("Cảnh báo giao diện:", " Hệ thống tự động chuyển sang khung cảnh báo màu đỏ với thông điệp: 'CẢNH BÁO: MẪU BẤT ĐỊNH — ĐỀ NGHỊ CHUYỂN CHUYÊN GIA SOI KÍNH HIỂN VI THỦ CÔNG'."),
    ("Ý nghĩa an toàn trong nông nghiệp:", " Trong dây chuyền sản xuất thực tế, việc đưa mẫu nghi ngờ vào đóng gói tự động sẽ gây rủi ro thương hiệu xuất khẩu."),
    ("Cơ chế Human-in-the-loop:", " Hệ thống AI sàng lọc 93.4% các mẫu rõ ràng, và tự giác 'bỏ phiếu trắng' để chuyển con người can thiệp đối với 6.6% các mẫu ranh giới khó.")
]
for title, content in items_s14_unc:
    p = tf_l14.add_paragraph()
    p.text = f"- {title}{content}"
    p.font.size = Pt(13)
    p.space_before = Pt(5)

# Right Column: Model Card & Ethics
add_card(s14, Inches(6.833), Inches(1.55), Inches(5.7), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_r14 = s14.shapes.add_textbox(Inches(7.083), Inches(1.7), Inches(5.2), Inches(4.9))
tf_r14 = tb_r14.text_frame
tf_r14.word_wrap = True

p = tf_r14.paragraphs[0]
p.text = "2. Trang Model Card & Cam Kết AI Có Trách Nhiệm:"
p.font.size = Pt(15.5)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s14_card = [
    ("Minh bạch kỹ thuật:", " Ứng dụng tích hợp trang Model Card (/model_card) công khai kiến trúc mô hình, tác giả, ngày huấn luyện và độ đo thực nghiệm."),
    ("Phạm vi hoạt động (Intended Use):", " Chỉ áp dụng phân loại giữa 2 giống gạo trắng bóc vỏ Cammeo và Osmancik; không dùng cho thóc nguyên vỏ, gạo lứt hay giống khác."),
    ("Phụ thuộc quy chuẩn trích xuất:", " Đầu vào bắt buộc phải tuân theo quy chuẩn trích xuất hình học đồng bộ với nghiên cứu gốc (độ phân giải camera, góc chụp)."),
    ("Đạo đức dữ liệu & Quyền riêng tư:", " Dữ liệu mở CC BY 4.0, không chứa thông tin cá nhân nhạy cảm, mã nguồn sạch.")
]
for title, content in items_s14_card:
    p = tf_r14.add_paragraph()
    p.text = f"- {title}{content}"
    p.font.size = Pt(13)
    p.space_before = Pt(5)


# ============================================================
# SLIDE 15: PHẦN 6 - ĐÓNG GÓP & PHÂN CHIA NHIỆM VỤ NHÓM
# ============================================================
s15 = prs.slides.add_slide(blank_layout)
add_slide_header(s15, "Đóng Góp Của Đề Tài & Phân Chia Trách Nhiệm Nhóm", "PHẦN 6: KẾT LUẬN & ĐÓNG GÓP")
add_slide_footer(s15, 15)

# Left Column: Key Contributions
add_card(s15, Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_l15 = s15.shapes.add_textbox(Inches(1.05), Inches(1.7), Inches(5.2), Inches(4.9))
tf_l15 = tb_l15.text_frame
tf_l15.word_wrap = True

p = tf_l15.paragraphs[0]
p.text = "1. Các Đóng Góp Chính Của Đề Tài:"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

contribs = [
    ("1. Minh chứng định lượng ccp_alpha:", " Cắt tỉa cây từ 187 nút lá xuống đúng 4 nút lá (giảm 97.9% độ phức tạp) mà độ chính xác Val lại tăng từ 89.24% lên 93.70%."),
    ("2. Trả lời câu hỏi nghiên cứu:", " Cây quyết định tối giản hoàn toàn có thể thay thế Rừng ngẫu nhiên trong các tình huống cần giải trình quy tắc logic rõ ràng cho nông dân."),
    ("3. Quy trình chuẩn, không rò rỉ dữ liệu:", " Thiết kế Stratified 60/20/20 nghiêm ngặt, niêm phong Test Set 100%, hạt giống seed=42 đảm bảo tái lập hoàn hảo."),
    ("4. Sản phẩm thực tế có trách nhiệm AI:", " Triển khai Web App hoàn chỉnh có cơ chế phát hiện mẫu bất định và Model Card minh bạch kỹ thuật.")
]
for title, c in contribs:
    p = tf_l15.add_paragraph()
    p.text = f"- {title}\n  {c}"
    p.font.size = Pt(12.5)
    p.space_before = Pt(5)

# Right Column: 2 Students 50% - 50%
# Student 1 Card
add_card(s15, Inches(6.833), Inches(1.55), Inches(5.7), Inches(2.5), bg_color=C_CARD_BG, border_color=C_PRIMARY)
tb_s1 = s15.shapes.add_textbox(Inches(7.033), Inches(1.68), Inches(5.3), Inches(2.25))
tf_s1 = tb_s1.text_frame
tf_s1.word_wrap = True

p = tf_s1.paragraphs[0]
p.text = "THÀNH VIÊN 1: CHỬ QUỐC HÀO — MSV: 10123112 (50%)"
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

items_s1 = [
    "- Quản lý dữ liệu, kiểm soát rò rỉ (Stratified Split 60/20/20).",
    "- Huấn luyện Baseline, thực hiện Thí nghiệm 1 (Overfitting max_depth).",
    "- Tối ưu hóa Thí nghiệm 2: Cắt tỉa ccp_alpha qua 5-Fold Cross Validation.",
    "- Lập trình máy chủ Backend REST API Flask & Kiểm tra dữ liệu 3 tầng.",
    "- Mã nguồn minh chứng: src/data_processing.py, src/train.py, app/app.py."
]
for item in items_s1:
    p = tf_s1.add_paragraph()
    p.text = item
    p.font.size = Pt(11.5)
    p.space_before = Pt(2)

# Student 2 Card
add_card(s15, Inches(6.833), Inches(4.25), Inches(5.7), Inches(2.5), bg_color=C_CARD_BG, border_color=C_SECONDARY)
tb_s2 = s15.shapes.add_textbox(Inches(7.033), Inches(4.38), Inches(5.3), Inches(2.25))
tf_s2 = tb_s2.text_frame
tf_s2.word_wrap = True

p = tf_s2.paragraphs[0]
p.text = "THÀNH VIÊN 2: [THÀNH VIÊN 2] — MSV: [MSV2] (50%)"
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = C_SECONDARY

items_s2 = [
    "- Huấn luyện Rừng ngẫu nhiên 100 cây, thực hiện Thí nghiệm 3 (Đối sánh).",
    "- Thực hiện Thí nghiệm 4: Phân tích lỗi sai ma trận nhầm lẫn & bất định.",
    "- Đánh giá chốt hạ trên tập Test Set độc lập.",
    "- Thiết kế giao diện Frontend Bootstrap 5, tính năng nộp file Batch CSV.",
    "- Hoàn thiện Báo cáo Word học thuật chuẩn trường & Slide thuyết trình."
]
for item in items_s2:
    p = tf_s2.add_paragraph()
    p.text = item
    p.font.size = Pt(11.5)
    p.space_before = Pt(2)


# ============================================================
# SLIDE 16: PHẦN 6 - KẾT LUẬN & HƯỚNG PHÁT TRIỂN
# ============================================================
s16 = prs.slides.add_slide(blank_layout)
add_slide_header(s16, "Kết Luận Chung & Hướng Mở Rộng Trong Tương Lai", "PHẦN 6: KẾT LUẬN & HƯỚNG PHÁT TRIỂN")
add_slide_footer(s16, 16)

# Left Column: Conclusions
add_card(s16, Inches(0.8), Inches(1.55), Inches(5.7), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_l16 = s16.shapes.add_textbox(Inches(1.05), Inches(1.7), Inches(5.2), Inches(4.9))
tf_l16 = tb_l16.text_frame
tf_l16.word_wrap = True

p = tf_l16.paragraphs[0]
p.text = "1. Kết Luận Chung Đề Tài:"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

concls = [
    "✓ Đề tài hoàn thành xuất sắc 100% mục tiêu và trả lời thấu đáo 3 câu hỏi nghiên cứu.",
    "✓ Mô hình Rừng ngẫu nhiên (100 cây) đạt hiệu năng tổng thể tối ưu nhất:",
    "   - Test Accuracy: 91.60% (vượt xa Baseline 57.22%).",
    "   - Test ROC-AUC: 0.9709 (phân tách xác suất tối ưu).",
    "✓ Mô hình Cây quyết định cắt tỉa (Pruned Tree):",
    "   - Test Accuracy: 91.47% (tương đương Rừng ngẫu nhiên).",
    "   - Cực kỳ tối giản (chỉ 4 nút lá, dùng 2 đặc trưng Major_Axis và Perimeter), minh bạch 100%.",
    "✓ Ứng dụng Web Flask đã tích hợp hoàn hảo cơ chế an toàn AI (Validation 3 tầng & Cảnh báo bất định)."
]
for item in concls:
    p = tf_l16.add_paragraph()
    p.text = item
    p.font.size = Pt(12.5)
    p.space_before = Pt(4)

# Right Column: Future Works
add_card(s16, Inches(6.833), Inches(1.55), Inches(5.7), Inches(5.2), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_r16 = s16.shapes.add_textbox(Inches(7.083), Inches(1.7), Inches(5.2), Inches(4.9))
tf_r16 = tb_r16.text_frame
tf_r16.word_wrap = True

p = tf_r16.paragraphs[0]
p.text = "2. Ba Hướng Mở Rộng Nghiên Cứu:"
p.font.size = Pt(16)
p.font.bold = True
p.font.color.rgb = C_PRIMARY

futures = [
    ("1. Tích hợp Thị giác máy tính (Computer Vision):", " Phát triển module tiền xử lý ảnh sử dụng OpenCV / YOLOv8 để tự động tách hạt gạo từ ảnh chụp thô trên băng chuyền và đo đạc trực tiếp 7 thông số 2D."),
    ("2. Mở rộng bộ dữ liệu lúa gạo đặc sản Việt Nam:", " Mở rộng nghiên cứu sang bài toán đa lớp (Multi-class Classification) cho các giống gạo xuất khẩu chủ lực của Việt Nam (ST24, ST25, Đài Thơm 8, Nàng Hoa 9, OM5451)."),
    ("3. Triển khai thiết bị nhúng biên (Edge AI):", " Nạp bộ luật cây quyết định 4 lá vào vi điều khiển giá rẻ (ESP32-S3, Raspberry Pi Pico) kết hợp màn hình LCD để tạo thiết bị cầm tay kiểm định lưu động cho cán bộ khuyến nông.")
]
for title, f in futures:
    p = tf_r16.add_paragraph()
    p.text = f"- {title}\n  {f}"
    p.font.size = Pt(12.5)
    p.space_before = Pt(6)


# ============================================================
# SLIDE 17: TRANG KẾT THÚC & Q&A
# ============================================================
s17 = prs.slides.add_slide(blank_layout)

# Top Bar
tb_top17 = s17.shapes.add_textbox(Inches(0.8), Inches(0.55), Inches(11.733), Inches(0.8))
tf_top17 = tb_top17.text_frame
p = tf_top17.paragraphs[0]
p.text = "TRƯỜNG ĐẠI HỌC SƯ PHẠM KỸ THUẬT HƯNG YÊN — KHOA CÔNG NGHỆ THÔNG TIN"
p.font.size = Pt(15)
p.font.bold = True
p.font.color.rgb = C_PRIMARY
p.alignment = PP_ALIGN.CENTER

# Big Center Card
c_center = add_card(s17, Inches(1.5), Inches(1.55), Inches(10.333), Inches(3.7), bg_color=C_CARD_BG, border_color=C_SECONDARY)

tb_c17 = s17.shapes.add_textbox(Inches(1.8), Inches(1.8), Inches(9.733), Inches(3.2))
tf_c17 = tb_c17.text_frame
tf_c17.word_wrap = True

p = tf_c17.paragraphs[0]
p.text = "XIN TRÂN TRỌNG CẢM ƠN QUÝ THẦY/CÔ HỘI ĐỒNG\nVÀ CÁC BẠN SINH VIÊN ĐÃ CHÚ Ý LẮNG NGHE!"
p.font.size = Pt(26)
p.font.bold = True
p.font.color.rgb = C_NAVY_DARK
p.alignment = PP_ALIGN.CENTER

p = tf_c17.add_paragraph()
p.text = "PHIÊN HỎI ĐÁP & BẢO VỆ ĐỒ ÁN (Q & A)"
p.font.size = Pt(20)
p.font.bold = True
p.font.color.rgb = C_ACCENT
p.alignment = PP_ALIGN.CENTER
p.space_before = Pt(14)

p = tf_c17.add_paragraph()
p.text = "Kính mong nhận được những nhận xét, góp ý quý báu của Quý Thầy/Cô để đề tài được hoàn thiện hơn nữa!"
p.font.size = Pt(14.5)
p.font.italic = True
p.font.color.rgb = C_DARK
p.alignment = PP_ALIGN.CENTER
p.space_before = Pt(10)

# Bottom Reference Card
add_card(s17, Inches(1.5), Inches(5.45), Inches(10.333), Inches(1.45), bg_color=C_WHITE, border_color=C_CARD_BORDER)
tb_ref = s17.shapes.add_textbox(Inches(1.8), Inches(5.6), Inches(9.733), Inches(1.15))
tf_ref = tb_ref.text_frame
tf_ref.word_wrap = True

p = tf_ref.paragraphs[0]
p.text = "THÔNG TIN LIÊN HỆ & TÀI NGUYÊN DỰ ÁN:"
p.font.size = Pt(13)
p.font.bold = True
p.font.color.rgb = C_SECONDARY

p = tf_ref.add_paragraph()
p.text = "- Mã nguồn GitHub: https://github.com/hao0979/Rice_Classification"
p.font.size = Pt(12.5)
p.font.color.rgb = C_DARK
p.space_before = Pt(3)

p = tf_ref.add_paragraph()
p.text = "- Báo cáo & Cẩm nang: BaoCao_BTL_PhanLoaiGiongLua_final.docx | Huong_dan_chi_tiet_va_Ly_thuyet_Project_17.docx"
p.font.size = Pt(12.5)
p.font.color.rgb = C_DARK
p.space_before = Pt(3)

# Save to both paths to ensure always in sync
try:
    prs.save(SLIDES_PATH)
    print(f"Presentation saved successfully to: {SLIDES_PATH}")
except PermissionError:
    print(f"presentation.pptx is locked by PowerPoint.")

alt_path = os.path.join(BASE_DIR, "slides", "presentation_final.pptx")
try:
    prs.save(alt_path)
    print(f"Presentation also saved to: {alt_path}")
except Exception as e:
    print(f"Error saving alt: {e}")


