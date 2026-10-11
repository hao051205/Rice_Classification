"""
generate_study_guide.py - Tạo file Word (.docx) Cẩm nang ôn tập & Hướng dẫn chi tiết từ A-Z
Bao gồm toàn bộ các bước làm, ý nghĩa thực tiễn, lý thuyết máy học, biểu đồ và câu hỏi vấn đáp.
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCX_PATH = os.path.join(BASE_DIR, "report", "Huong_dan_chi_tiet_va_Ly_thuyet_Project_17.docx")
ROOT_DOCX_PATH = os.path.join(os.path.dirname(BASE_DIR), "Huong_dan_chi_tiet_va_Ly_thuyet_Project_17.docx")
FIGURES_DIR = os.path.join(BASE_DIR, "report", "figures")


def set_cell_background(cell, fill_hex):
    """Màu nền cho ô bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Đặt lề trong cho ô bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)


def add_callout(doc, text, title="GHI NHỚ QUAN TRỌNG:"):
    """Hộp ghi chú nổi bật."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "F0F4F8")
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)

    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    r_title = p.add_run(f"📌 {title} ")
    r_title.bold = True
    r_title.font.color.rgb = RGBColor(24, 76, 120)
    r_text = p.add_run(text)
    r_text.font.size = Pt(10.5)
    doc.add_paragraph()


def add_image_with_caption(doc, filename, caption_text, width_inch=5.8):
    """Chèn ảnh và chú thích căn giữa."""
    img_path = os.path.join(FIGURES_DIR, filename)
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(4)
        run = p_img.add_run()
        run.add_picture(img_path, width=Inches(width_inch))

        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_after = Pt(10)
        run_cap = p_cap.add_run(caption_text)
        run_cap.font.size = Pt(9.5)
        run_cap.font.italic = True
        run_cap.font.color.rgb = RGBColor(100, 100, 100)


def format_table(tbl, col_widths, headers, data_rows):
    """Định dạng bảng đẹp."""
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    hdr_cells = tbl.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        hdr_cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        hdr_cells[i].paragraphs[0].runs[0].font.size = Pt(10)
        set_cell_background(hdr_cells[i], "184C78")
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=150, right=150)

    for row_idx, data in enumerate(data_rows):
        row = tbl.add_row()
        fill_color = "F9FBFC" if row_idx % 2 == 1 else "FFFFFF"
        for col_idx, val in enumerate(data):
            cell = row.cells[col_idx]
            cell.text = str(val)
            p = cell.paragraphs[0]
            p.runs[0].font.size = Pt(9.5)
            set_cell_background(cell, fill_color)
            set_cell_margins(cell, top=100, bottom=100, left=150, right=150)

    for row in tbl.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = Inches(width)


def generate_study_guide():
    doc = Document()

    # Lề trang
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Phong cách mặc định
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(40, 40, 40)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # ============================================================
    # TRANG BÌA
    # ============================================================
    p_header = doc.add_paragraph()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_hdr = p_header.add_run("CHUYÊN ĐỀ HỌC MÁY & THỊ GIÁC MÁY TÍNH\n-----------------------------------------")
    r_hdr.font.size = Pt(11)
    r_hdr.font.bold = True
    r_hdr.font.color.rgb = RGBColor(24, 76, 120)

    for _ in range(2):
        doc.add_paragraph()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t1 = p_title.add_run("CẨM NANG ÔN TẬP & HƯỚNG DẪN CHI TIẾT TỪ A - Z\n")
    r_t1.font.size = Pt(16)
    r_t1.font.bold = True
    r_t1.font.color.rgb = RGBColor(100, 100, 100)

    r_t2 = p_title.add_run("CÁC BƯỚC LÀM, Ý NGHĨA THỰC TIỄN & CƠ SỞ LÝ THUYẾT\nPROJECT 17: PHÂN LOẠI HAI GIỐNG GẠO TỪ HÌNH DẠNG HẠT")
    r_t2.font.size = Pt(18)
    r_t2.font.bold = True
    r_t2.font.color.rgb = RGBColor(24, 76, 120)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("\nTài liệu phục vụ học tập, nắm vững bản chất thuật toán và bảo vệ vấn đáp kết thúc học phần")
    r_sub.font.size = Pt(11.5)
    r_sub.font.italic = True

    for _ in range(4):
        doc.add_paragraph()

    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_meta = p_meta.add_run(
        "Chuyên đề trọng tâm: Bài 6 — Cây quyết định, Cắt tỉa (Pruning) và Rừng ngẫu nhiên\n"
        "Đối tượng: Sinh viên thực hiện Project 17\n"
        "Năm học: 2026"
    )
    r_meta.font.size = Pt(11)
    r_meta.font.bold = True

    doc.add_page_break()

    # ============================================================
    # LỜI MỞ ĐẦU
    # ============================================================
    h0 = doc.add_heading("LỜI MỞ ĐẦU: LỘ TRÌNH 7 BƯỚC THỰC HIỆN PROJECT", level=1)
    h0.paragraph_format.space_before = Pt(12)

    doc.add_paragraph(
        "Tài liệu này được biên soạn nhằm giúp sinh viên hiểu sâu sắc, toàn diện từng công đoạn trong toàn bộ chu trình "
        "phát triển một dự án Học máy (Machine Learning Pipeline) hoàn chỉnh. Dự án không chỉ dừng lại ở việc chạy mã nguồn, "
        "mà đòi hỏi người thực hiện phải nắm vững bản chất toán học, trả lời được lý do vì sao chọn phương pháp đó và phân tích "
        "được các trường hợp thành công lẫn thất bại của mô hình."
    )
    doc.add_paragraph(
        "Toàn bộ dự án được triển khai tuần tự theo 7 bước cốt lõi:\n"
        "• Bước 1: Khám phá dữ liệu (EDA), kiểm tra chất lượng và phân chia dữ liệu chống rò rỉ (Leakage-free Split).\n"
        "• Bước 2: Xây dựng mô hình cơ sở thô sơ (Dummy Baseline) làm cột mốc đối chứng.\n"
        "• Bước 3: Huấn luyện Cây quyết định (Decision Tree) và làm Thí nghiệm 1 khảo sát hiện tượng Quá khớp (Overfitting).\n"
        "• Bước 4: Ứng dụng kỹ thuật Cắt tỉa chi phí độ phức tạp (Cost-Complexity Pruning - ccp_alpha) và làm Thí nghiệm 2 trên 5-fold CV.\n"
        "• Bước 5: Xây dựng Rừng ngẫu nhiên (Random Forest) và làm Thí nghiệm 3 so sánh toàn diện giữa Cây đơn và Rừng.\n"
        "• Bước 6: Đánh giá một lần duy nhất trên tập Test độc lập và làm Thí nghiệm 4 phân tích bản chất vùng đặc trưng sai số.\n"
        "• Bước 7: Đóng gói thành Web App Flask phục vụ mô hình offline, kiểm tra dữ liệu 3 tầng và tích hợp tính năng mở rộng."
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 1
    # ============================================================
    doc.add_heading("1. BẢN CHẤT BÀI TOÁN & 7 ĐẶC TRƯNG HÌNH HỌC", level=1)
    
    doc.add_heading("1.1. Bối cảnh và bài toán thực tế", level=2)
    doc.add_paragraph(
        "Trong ngành công nghiệp chế biến và đóng gói nông sản, hai giống lúa chất lượng cao là Cammeo và Osmancik có giá thành và "
        "yêu cầu bảo quản khác nhau. Việc lẫn tạp giữa hai giống làm giảm uy tín thương hiệu. Nếu để công nhân phân loại thủ công bằng mắt, "
        "sau vài giờ làm việc mắt sẽ mỏi, dẫn đến tốc độ chậm và phán đoán sai lệch. Giải pháp tự động hóa bằng máy học sử dụng camera "
        "chụp ảnh từng hạt gạo trên băng chuyền, tách nền và đo đạc 7 đặc trưng hình học để máy tính tự động ra quyết định trong vài mili-giây."
    )

    doc.add_heading("1.2. Bản chất toán học của bài toán", level=2)
    doc.add_paragraph(
        "• Loại bài toán: Phân loại nhị phân có giám sát (Supervised Binary Classification).\n"
        "• Đơn vị quan sát (Unit of Observation): Đúng 1 hạt gạo duy nhất.\n"
        "• Không gian đặc trưng đầu vào: Véc-tơ 7 chiều X = [x1, x2, ..., x7] thuộc tập số thực R^7.\n"
        "• Nhãn đầu ra: y thuộc {Cammeo, Osmancik}."
    )

    doc.add_heading("1.3. Ý nghĩa vật lý và hình thái học của 7 đặc trưng", level=2)
    tbl_feat = doc.add_table(rows=1, cols=4)
    format_table(
        tbl_feat,
        [1.8, 1.0, 1.2, 2.5],
        ["Tên đặc trưng", "Đơn vị", "Khoảng giá trị", "Ý nghĩa hình thái học"],
        [
            ["Area", "pixel²", "7,551 – 18,913", "Tổng số pixel thuộc về hạt gạo. Thể hiện kích thước tổng thể."],
            ["Perimeter", "pixel", "359.1 – 548.5", "Chu vi đường biên bao quanh chu vi hạt gạo."],
            ["Major_Axis_Length", "pixel", "145.3 – 239.0", "Chiều dài trục chính của elip tương đương (khoảng cách dài nhất). Đây là đặc trưng phân biệt mạnh nhất."],
            ["Minor_Axis_Length", "pixel", "59.5 – 107.5", "Chiều dài trục phụ của elip tương đương (bề ngang rộng nhất)."],
            ["Eccentricity", "Vô thứ nguyên", "0.78 – 0.95", "Độ lệch tâm elip. 0 là tròn tuyệt đối; càng gần 1 thì hạt càng dẹp và dài."],
            ["Convex_Area", "pixel²", "7,723 – 19,099", "Diện tích đa giác lồi nhỏ nhất bao kín hạt gạo."],
            ["Extent", "Tỷ lệ [0, 1]", "0.50 – 0.86", "Tỷ lệ diện tích: Extent = Area / (Diện tích khung chữ nhật bao ngoài)."]
        ]
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 2
    # ============================================================
    doc.add_heading("2. BƯỚC 1: TIỀN XỬ LÝ, EDA & QUY TẮC CHỐNG RÒ RỈ DỮ LIỆU", level=1)

    doc.add_heading("2.1. Khám phá chất lượng dữ liệu (EDA)", level=2)
    doc.add_paragraph(
        "Trước khi đưa dữ liệu vào bất kỳ thuật toán nào, ta phải thực hiện EDA (Exploratory Data Analysis):\n"
        "• Missing Values: 0 giá trị khuyết thiếu trên toàn bộ 3,810 dòng.\n"
        "• Trùng lặp (Duplicates): 0 bản ghi trùng lặp hoàn toàn.\n"
        "• Phân bố lớp: Osmancik chiếm 2,180 mẫu (57.2%), Cammeo chiếm 1,630 mẫu (42.8%). Dữ liệu mất cân bằng nhẹ (~1.34 : 1), "
        "chưa đến mức phải can thiệp bằng các kỹ thuật oversampling phức tạp (như SMOTE) nhưng cần lưu ý khi đánh giá thước đo F1-score."
    )

    add_image_with_caption(doc, "01_distributions.png", "Hình 1: Phân phối 7 đặc trưng hình học giữa hai giống gạo Cammeo và Osmancik.", width_inch=6.0)
    add_image_with_caption(doc, "03_correlation_matrix.png", "Hình 2: Ma trận tương quan Pearson giữa 7 đặc trưng hình học.", width_inch=5.2)

    doc.add_paragraph(
        "Nhận xét ma trận tương quan (Hình 2):\n"
        "Nhóm 4 biến kích cỡ (Area, Perimeter, Major_Axis_Length, Convex_Area) có hệ số tương quan rất cao (r > 0.95). "
        "Hiện tượng này gọi là Đa cộng tuyến (Multicollinearity). Nếu sử dụng Hồi quy Logistic, đa cộng tuyến sẽ làm ma trận nghịch đảo "
        "bị bất ổn định và làm sai lệch trọng số hồi quy. Tuy nhiên, đối với Cây quyết định và Rừng ngẫu nhiên, thuật toán chỉ xét từng đặc trưng "
        "đơn lẻ tại mỗi nút phân chia, nên hoàn toàn không bị ảnh hưởng tiêu cực bởi đa cộng tuyến."
    )

    doc.add_heading("2.2. Chiến lược phân chia dữ liệu: Stratified 60/20/20", level=2)
    doc.add_paragraph(
        "Bộ dữ liệu được chia tách thành 3 tập riêng biệt:\n"
        "• Train Set (60% = 2,286 mẫu): Dùng để học cây, tính toán bảng ccp_alpha và huấn luyện các mô hình.\n"
        "• Validation Set (20% = 762 mẫu): Dùng để thử nghiệm các độ sâu max_depth, chọn siêu tham số alpha và so sánh các mô hình.\n"
        "• Test Set (20% = 762 mẫu): Dùng để nghiệm thu cuối cùng."
    )

    doc.add_paragraph(
        "Lý thuyết Phân chia tầng (Stratified Splitting):\n"
        "Nếu chia ngẫu nhiên thông thường (random split), có thể do ngẫu nhiên mà tập Train nhận quá nhiều hạt Cammeo, trong khi tập Test "
        "lại toàn hạt Osmancik. Phân chia tầng (Stratified) can thiệp để giữ nguyên tỷ lệ 57.2% Osmancik và 42.8% Cammeo đồng nhất trên cả 3 tập."
    )

    add_callout(
        doc,
        "Rò rỉ dữ liệu (Data Leakage) xảy ra khi thông tin từ tập Test đi vào quá trình huấn luyện hoặc chọn mô hình. "
        "Để ngăn chặn rò rỉ tuyệt đối: (1) Chia Train/Val/Test ngay từ đầu; (2) Mọi thống kê [min, max] chỉ tính trên Train; "
        "(3) Khóa chặt tập Test, chỉ nạp ra đánh giá đúng một lần duy nhất lúc nghiệm thu.",
        "QUY TẮC PHÒNG CHỐNG DATA LEAKAGE"
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 3
    # ============================================================
    doc.add_heading("3. BƯỚC 2: MÔ HÌNH CƠ SỞ (BASELINE) & Ý NGHĨA ĐỐI CHỨNG", level=1)

    doc.add_paragraph(
        "Trong nghiên cứu khoa học, không bao giờ được đánh giá một mô hình trong sự cô lập. Nếu một mô hình đạt 60% độ chính xác, "
        "con số đó tốt hay xấu? Nếu một lớp chiếm sẵn 57.2% dữ liệu, thì một mô hình 'ngu ngốc' chỉ cần luôn luôn đoán lớp đó cũng đã "
        "đạt sẵn 57.2% độ chính xác mà không cần học bất kỳ đặc trưng nào!\n\n"
        "Do đó, ta thiết lập DummyClassifier với chiến lược 'most_frequent': Luôn đoán nhãn đa số là Osmancik.\n"
        "• Baseline Accuracy = 57.22%\n"
        "• Baseline F1-score (cho Cammeo) = 0.0000\n"
        "Mục tiêu bắt buộc của dự án: Mọi mô hình máy học đề xuất phải vượt xa mốc 57.22% này."
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 4
    # ============================================================
    doc.add_heading("4. BƯỚC 3: CÂY QUYẾT ĐỊNH & THÍ NGHIỆM 1 (OVERFITTING)", level=1)

    doc.add_heading("4.1. Cơ sở lý thuyết của Cây quyết định (Decision Tree)", level=2)
    doc.add_paragraph(
        "Cây quyết định phân chia không gian đặc trưng thành các hình siêu chữ nhật trực giao thông qua cấu trúc cây:\n"
        "• Nút gốc (Root node): Chứa toàn bộ mẫu ban đầu.\n"
        "• Nút quyết định (Decision node): Chứa điều kiện nhị phân (Ví dụ: Major_Axis_Length <= 191.05).\n"
        "• Nút lá (Leaf node): Chứa kết luận dự đoán nhãn lớp."
    )

    doc.add_paragraph(
        "Chỉ số tạp chất Gini (Gini Impurity):\n"
        "Cây sử dụng thuật toán CART với chỉ số Gini để đo độ hỗn loạn tại mỗi nút m:\n"
        "   Gini(m) = 1 - sum(p_i^2)\n"
        "Trong đó p_i là tỷ lệ mẫu thuộc lớp i tại nút m.\n"
        "• Nếu nút chứa 100% hạt Cammeo: Gini = 1 - 1^2 = 0 (Thuần khiết tuyệt đối).\n"
        "• Nếu nút chứa 50% Cammeo và 50% Osmancik: Gini = 1 - (0.5^2 + 0.5^2) = 0.5 (Tạp chất cực đại).\n"
        "Tại mỗi bước, cây quét qua tất cả đặc trưng và các ngưỡng cắt khả dĩ để chọn điểm cắt làm giảm tạp chất (Gini Gain) lớn nhất."
    )

    doc.add_heading("4.2. Thí nghiệm 1: Khảo sát max_depth và Hiện tượng Overfitting", level=2)
    doc.add_paragraph(
        "Khi cho max_depth chạy từ 1 đến 20 (Hình 3):\n"
        "• Trên tập Train: Độ chính xác tăng liên tục từ 92.9% lên 100.0% tuyệt đối khi depth >= 16.\n"
        "• Trên tập Validation: Độ chính xác đạt đỉnh 93.70% ở độ sâu rất nhỏ (depth = 1, 2), nhưng sau đó giảm dần xuống còn 89.10%.\n\n"
        "Giải thích hiện tượng Quá khớp (Overfitting):\n"
        "Cây không bị giới hạn đã học thuộc lòng cả các mẫu dị dạng và nhiễu ngẫu nhiên của tập Train, dẫn đến phương sai (Variance) "
        "tăng vọt, làm mất khả năng tổng quát hóa khi gặp dữ liệu mới trên tập Validation. Do đó, bắt buộc phải Cắt tỉa (Pruning)."
    )

    add_image_with_caption(doc, "05_depth_curve.png", "Hình 3: Đường hiệu năng Accuracy theo max_depth trên tập Train và Validation.", width_inch=5.8)

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 5
    # ============================================================
    doc.add_heading("5. BƯỚC 4: CẮT TỈA CÂY (COST-COMPLEXITY PRUNING) & THÍ NGHIỆM 2", level=1)

    doc.add_heading("5.1. Lý thuyết tham số ccp_alpha", level=2)
    doc.add_paragraph(
        "Thay vì chặn cây sớm (Pre-pruning), kỹ thuật Hậu cắt tỉa Cost-Complexity Pruning (ccp_alpha) cho cây lớn đầy đủ rồi tỉa ngược "
        "từ dưới lá lên dựa trên hàm chi phí:\n"
        "   R_alpha(T) = R(T) + alpha * |T|\n"
        "Trong đó:\n"
        "• R(T): Tổng sai số phân loại của cây T trên tập huấn luyện.\n"
        "• |T|: Số lượng nút lá (thể hiện kích thước/độ cồng kềnh của cây).\n"
        "• alpha (ccp_alpha): Hệ số phạt độ phức tạp.\n\n"
        "Ý nghĩa: Khi alpha tăng, các nhánh phụ có độ suy giảm tạp chất đem lại nhỏ hơn chi phí phạt alpha sẽ bị chặt bỏ. "
        "Điều này giúp cân bằng hoàn hảo giữa độ chính xác và độ gọn gàng của mô hình."
    )

    doc.add_heading("5.2. Thí nghiệm 2: 5-Fold Stratified Cross-Validation", level=2)
    doc.add_paragraph(
        "Hàm cost_complexity_pruning_path tìm được 70 giá trị alpha khả dĩ. Chạy 5-fold CV trên tập Train (Hình 4):\n"
        "• Giá trị alpha tối ưu tìm được: ccp_alpha = 0.008024.\n"
        "• Điểm số CV Accuracy đạt: 92.48% +- 1.59%.\n"
        "Cây sau khi cắt tỉa (Hình 5) chỉ còn lại cấu trúc cực kỳ tối giản: luật phân loại nút gốc duy nhất dựa vào Major_Axis_Length <= 191.05 pixel."
    )

    add_image_with_caption(doc, "06_pruning_curve.png", "Hình 4: Điểm số Cross-Validation theo hệ số phạt ccp_alpha.", width_inch=5.8)
    add_image_with_caption(doc, "07_decision_tree.png", "Hình 5: Cấu trúc Cây quyết định tối giản sau khi cắt tỉa ccp_alpha.", width_inch=6.2)

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 6
    # ============================================================
    doc.add_heading("6. BƯỚC 5: RỪNG NGẪU NHIÊN & THÍ NGHIỆM 3 (SO SÁNH ĐỐI CHỨNG)", level=1)

    doc.add_heading("6.1. Cơ sở lý thuyết của Rừng ngẫu nhiên (Random Forest)", level=2)
    doc.add_paragraph(
        "Rừng ngẫu nhiên là phương pháp học kết hợp (Ensemble Learning) theo cơ chế Bagging (Bootstrap Aggregating) gồm 100 cây con:\n"
        "1. Bootstrap Sampling: Mỗi cây con chỉ học trên một tập mẫu con được lấy ngẫu nhiên có hoàn lại từ tập Train.\n"
        "2. Random Feature Subsampling: Tại mỗi nút chia, thuật toán chỉ chọn ngẫu nhiên một tập con đặc trưng (sqrt(7) xấp xỉ 2 đặc trưng) "
        "để tìm điểm chia, giúp các cây con không bị tương quan với nhau.\n"
        "3. Majority Voting: Kết quả dự đoán là số phiếu bầu đa số của 100 cây con."
    )

    doc.add_heading("6.2. Thí nghiệm 3: Bảng so sánh Cây đơn vs Rừng ngẫu nhiên", level=2)
    tbl_cmp = doc.add_table(rows=1, cols=4)
    format_table(
        tbl_cmp,
        [2.2, 1.4, 1.4, 1.4],
        ["Chỉ số đánh giá", "Baseline (Dummy)", "Cây quyết định (Pruned)", "Rừng ngẫu nhiên (100 cây)"],
        [
            ["Validation Accuracy", "57.22%", "93.70%", "93.44%"],
            ["Validation F1-score", "0.00%", "92.77%", "92.28%"],
            ["Validation ROC-AUC", "0.5000", "0.9596", "0.9741"],
            ["Test Accuracy (Chạy 1 lần)", "57.22%", "91.47%", "91.60%"],
            ["Test ROC-AUC", "0.5000", "0.9344", "0.9709"],
            ["Thời gian huấn luyện", "0.001s", "0.036s", "0.320s"],
            ["Tính minh bạch / Giải thích", "Không", "Rất cao (Đọc được luật)", "Hộp đen (Khó giải thích)"]
        ]
    )

    doc.add_paragraph()
    add_image_with_caption(doc, "10_confusion_matrices.png", "Hình 6: Ma trận nhầm lẫn trên tập Validation.", width_inch=6.0)
    add_image_with_caption(doc, "11_roc_curves.png", "Hình 7: Đường cong ROC so sánh năng lực phân tách giữa Cây và Rừng.", width_inch=5.5)

    doc.add_paragraph(
        "Trả lời câu hỏi nghiên cứu của đề tài:\n"
        "• Về độ chính xác: Cây quyết định cắt tỉa (91.47%) hoàn toàn tương đương Rừng ngẫu nhiên (91.60%).\n"
        "• Về tính minh bạch: Cây quyết định vượt trội vì có thể diễn giải thành luật if-else rõ ràng cho kỹ thuật viên tại cơ sở đóng gói.\n"
        "• Khi nào cần Rừng ngẫu nhiên? Khi hệ thống đòi hỏi đường cong xác suất mượt mà (ROC-AUC = 0.9709) để điều khiển van khí nén tự động."
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 7
    # ============================================================
    doc.add_heading("7. BƯỚC 6: ĐÁNH GIÁ TẬP TEST & PHÂN TÍCH LỖI CHI TIẾT (TN4)", level=1)

    doc.add_heading("7.1. Đánh giá tập Test độc lập (Chạy 1 lần duy nhất)", level=2)
    doc.add_paragraph(
        "Trên tập Test (762 mẫu độc lập), mô hình Rừng ngẫu nhiên đạt Test Accuracy 91.60%, trong đó:\n"
        "• Cammeo: Đoán đúng 285 hạt, đoán sai 41 hạt.\n"
        "• Osmancik: Đoán đúng 413 hạt, đoán sai 23 hạt."
    )
    add_image_with_caption(doc, "13_test_confusion_matrices.png", "Hình 8: Ma trận nhầm lẫn trên tập Test độc lập.", width_inch=6.0)

    doc.add_heading("7.2. Thí nghiệm 4: Phân tích vùng đặc trưng sai số (Error Analysis)", level=2)
    doc.add_paragraph(
        "Bóc tách 50 mẫu bị phân loại sai trên tập Validation:\n"
        "1. Nguyên nhân hình thái học: Chiều dài trục chính (Major_Axis_Length) của mẫu đúng có trung bình là 188.83 pixel, "
        "trong khi mẫu đoán sai có trung bình là 191.75 pixel. Ngưỡng cắt tự nhiên của 2 giống lúa nằm ở đúng 191.05 pixel! "
        "Nghĩa là các mẫu sai rơi đúng vào 'vùng xám' giao thoa kích thước tự nhiên giữa hai giống lúa.\n"
        "2. Năng lực tự nhận thức xác suất của mô hình: Các mẫu đoán đúng có độ tự tin trung bình rất cao (94.37%), "
        "trong khi các mẫu đoán sai có độ tự tin sụt giảm mạnh chỉ còn 78.28% (xác suất gần ngưỡng phân vân 50/50)."
    )
    add_image_with_caption(doc, "12_error_analysis.png", "Hình 9: So sánh phân phối đặc trưng giữa mẫu dự đoán Đúng và Sai.", width_inch=6.0)

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 8
    # ============================================================
    doc.add_heading("8. BƯỚC 7: ỨNG DỤNG WEB FLASK & SỬ DỤNG CÓ TRÁCH NHIỆM", level=1)

    doc.add_paragraph(
        "8.1. Tách biệt Train Offline khỏi Serving Online:\n"
        "Mô hình được lưu tĩnh thành tệp final_model.pkl. Server Flask nạp mô hình một lần duy nhất vào bộ nhớ RAM khi khởi động. "
        "Mỗi request từ người dùng chỉ thực hiện hàm predict() trong thời gian chưa đầy 1 mili-giây, không huấn luyện lại mô hình.\n\n"
        "8.2. Hệ thống kiểm soát dữ liệu đầu vào 3 tầng:\n"
        "• Tầng 1: Kiểm tra thiếu trường dữ liệu (Báo lỗi HTTP 400).\n"
        "• Tầng 2: Kiểm tra sai định dạng số (Báo lỗi 400).\n"
        "• Tầng 3: Kiểm tra ngoài miền [min, max] của tập Train (Cảnh báo vàng: dữ liệu ngoài miền huấn luyện).\n\n"
        "8.3. Hai tính năng mở rộng đạt điểm thưởng:\n"
        "⭐ Tính năng 1: Cảnh báo bất định (Uncertainty Detection) khi xác suất phân lớp nằm trong dải [0.40, 0.60] (gần ngưỡng 50/50), "
        "yêu cầu chuyển sang chuyên viên kiểm tra trực tiếp bằng mắt.\n"
        "⭐ Tính năng 2: Phân loại theo lô (Batch CSV) cho phép upload file CSV hàng trăm hạt, tính toán tức thì và xuất báo cáo kết quả.\n\n"
        "8.4. Sử dụng có trách nhiệm:\n"
        "• Không áp dụng cho các giống gạo khác (như ST25, Jasmine) vì mô hình sẽ ép gán nhãn sai lệch.\n"
        "• Không thay thế chứng thư kiểm định chất lượng chính thức của cơ quan nhà nước."
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 9: CÂU HỎI VẤN ĐÁP
    # ============================================================
    doc.add_heading("9. BỘ CÂU HỎI VẤN ĐÁP BẢO VỆ ĐỒ ÁN (HỎI & ĐÁP)", level=1)

    doc.add_paragraph(
        "Câu 1: ccp_alpha làm cây thay đổi như thế nào?\n"
        "Trả lời: ccp_alpha là hệ số phạt độ phức tạp trong hàm chi phí R_alpha(T) = R(T) + alpha * |T|. Khi alpha = 0, cây giữ nguyên kích thước đầy đủ (dễ quá khớp). "
        "Khi alpha tăng, các nhánh phụ có độ giảm tạp chất nhỏ hơn chi phí phạt alpha sẽ bị cắt bỏ, giúp cây nông hơn và tăng khả năng tổng quát hóa.\n\n"
        "Câu 2: Feature Importance của Rừng ngẫu nhiên có thiên lệch gì?\n"
        "Trả lời: MDI Feature Importance có xu hướng thiên lệch cho các biến liên tục có nhiều giá trị hoặc các biến có tương quan mạnh với nhau. "
        "Trong bài này, nhóm 4 biến kích thước (Area, Convex_Area, Perimeter, Major_Axis_Length) có tương quan r > 0.95, do đó chúng phân chia độ quan trọng cho nhau giữa các cây con.\n\n"
        "Câu 3: Khi nào chọn Cây quyết định thay vì Rừng ngẫu nhiên?\n"
        "Trả lời: Chọn Cây quyết định khi bài toán ưu tiên tính minh bạch (giải thích được luật cho công nhân), khi triển khai trên phần cứng nhúng bị giới hạn tài nguyên tính toán, "
        "và khi cây đơn đã đạt độ chính xác tiệm cận rừng (91.47% so với 91.60%).\n\n"
        "Câu 4: Tại sao không được tune tham số trên tập Test?\n"
        "Trả lời: Vì nếu dùng tập Test để chọn tham số hay chọn mô hình, tập Test sẽ bị 'ô nhiễm' (data leakage), mô hình sẽ gián tiếp học thuộc tập Test và kết quả đánh giá không còn khách quan."
    )

    doc.save(DOCX_PATH)
    doc.save(ROOT_DOCX_PATH)
    print(f"Study guide DOCX saved to:\n1. {DOCX_PATH}\n2. {ROOT_DOCX_PATH}")


if __name__ == "__main__":
    generate_study_guide()
