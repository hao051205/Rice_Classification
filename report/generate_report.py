"""
generate_report.py - Tạo file Báo cáo học thuật chi tiết DOCX chuẩn mẫu
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DOCX_PATH = os.path.join(BASE_DIR, "report", "Bao_cao_Project_17_Phan_loai_hat_gao.docx")
FIGURES_DIR = os.path.join(BASE_DIR, "report", "figures")


def set_cell_background(cell, fill_hex):
    """Đặt màu nền cho ô trong bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Đặt padding cho cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)


def add_callout(doc, text, title="LƯU Ý QUAN TRỌNG:"):
    """Tạo hộp thông báo nổi bật."""
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, "F0F4F8")
    set_cell_margins(cell, top=140, bottom=140, left=200, right=200)

    p = cell.paragraphs[0]
    r_title = p.add_run(f"📌 {title} ")
    r_title.bold = True
    r_title.font.color.rgb = RGBColor(24, 76, 120)
    r_text = p.add_run(text)
    r_text.font.size = Pt(10.5)
    doc.add_paragraph()  # khoảng trống


def add_image_with_caption(doc, filename, caption_text, width_inch=6.0):
    """Thêm ảnh kèm chú thích căn giữa."""
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
    """Định dạng bảng chuẩn học thuật."""
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Header
    hdr_cells = tbl.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        hdr_cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        hdr_cells[i].paragraphs[0].runs[0].font.bold = True
        hdr_cells[i].paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 255, 255)
        hdr_cells[i].paragraphs[0].runs[0].font.size = Pt(10)
        set_cell_background(hdr_cells[i], "184C78")
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=150, right=150)

    # Data
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

    # Set width
    for row in tbl.rows:
        for idx, width in enumerate(col_widths):
            row.cells[idx].width = Inches(width)


def generate_document():
    doc = Document()

    # Page Margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base styles
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(40, 40, 40)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # ============================================================
    # TRANG BÌA
    # ============================================================
    p_uni = doc.add_paragraph()
    p_uni.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_uni = p_uni.add_run("TRƯỜNG ĐẠI HỌC BÁCH KHOA\nKHOA CÔNG NGHỆ THÔNG TIN\n-------------------------")
    r_uni.font.size = Pt(12)
    r_uni.font.bold = True
    r_uni.font.color.rgb = RGBColor(24, 76, 120)

    for _ in range(3):
        doc.add_paragraph()

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_t1 = p_title.add_run("BÁO CÁO DỰ ÁN KẾT THÚC HỌC PHẦN\n")
    r_t1.font.size = Pt(16)
    r_t1.font.bold = True
    r_t1.font.color.rgb = RGBColor(100, 100, 100)

    r_t2 = p_title.add_run("PROJECT 17: PHÂN LOẠI HAI GIỐNG GẠO TỪ HÌNH DẠNG HẠT\n(CAMMEO VÀ OSMANCIK)")
    r_t2.font.size = Pt(20)
    r_t2.font.bold = True
    r_t2.font.color.rgb = RGBColor(24, 76, 120)

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("\nChuyên đề: Cây quyết định, Cắt tỉa (Pruning) và Rừng ngẫu nhiên\nHọc phần: Học máy & Thị giác máy tính")
    r_sub.font.size = Pt(12)
    r_sub.font.italic = True

    for _ in range(5):
        doc.add_paragraph()

    p_info = doc.add_paragraph()
    p_info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_info = p_info.add_run(
        "NHÓM SINH VIÊN THỰC HIỆN: NHÓM 17\n"
        "Thành viên 1: [Họ và tên SV 1] — MSSV: [Mã số SV 1]\n"
        "Thành viên 2: [Họ và tên SV 2] — MSSV: [Mã số SV 2]\n\n"
        "HÀ NỘI — 2026"
    )
    r_info.font.size = Pt(12)
    r_info.font.bold = True

    doc.add_page_break()

    # ============================================================
    # MỤC LỤC & TÓM TẮT
    # ============================================================
    h_abs = doc.add_heading("TÓM TẮT DỰ ÁN (ABSTRACT)", level=1)
    h_abs.paragraph_format.space_before = Pt(12)

    doc.add_paragraph(
        "Báo cáo này trình bày toàn bộ quy trình thiết kế, thực nghiệm và triển khai giải pháp học máy cho bài toán "
        "phân loại hai giống gạo chất lượng cao (Cammeo và Osmancik) dựa trên 7 đặc trưng hình học trích xuất từ hình ảnh hạt. "
        "Xuất phát từ nhu cầu thực tiễn của các cơ sở phân loại và đóng gói nông sản, dự án tập trung giải quyết câu hỏi nghiên cứu cốt lõi: "
        "'Liệu một mô hình Cây quyết định tối giản đã qua cắt tỉa có đủ chính xác và dễ giải thích hay bắt buộc phải đánh đổi bằng Rừng ngẫu nhiên phức tạp?'"
    )
    doc.add_paragraph(
        "Quy trình nghiên cứu tuân thủ nghiêm ngặt nguyên tắc ngăn chặn rò rỉ dữ liệu (Data Leakage) bằng phân chia tầng (Stratified Split) "
        "với tỷ lệ 60% Train (2,286 mẫu), 20% Validation (762 mẫu) và 20% Test (762 mẫu) được bảo vệ tuyệt đối. "
        "Mô hình đối chứng cơ sở (Dummy Baseline) đạt độ chính xác 57.22%. Qua 4 thí nghiệm thực chứng, kết quả cho thấy: "
        "(1) Cây quyết định không khống chế độ sâu bị quá khớp trầm trọng (Train 100%, Val 89.1%); "
        "(2) Kỹ thuật Cost-Complexity Pruning với ccp_alpha = 0.008024 tối ưu hóa cây trên 5-fold CV đạt độ chính xác 92.48% ± 1.59%; "
        "(3) Trên tập Test độc lập, Cây quyết định cắt tỉa đạt Accuracy 91.47%, F1-score 90.08%, ROC-AUC 0.9344; trong khi Rừng ngẫu nhiên "
        "(100 cây) đạt Accuracy 91.60%, F1-score 89.91%, ROC-AUC 0.9709. "
        "Kết quả khẳng định Cây quyết định tối giản hoàn toàn có thể thay thế Rừng ngẫu nhiên nếu ưu tiên khả năng giải thích và kiểm toán quy tắc. "
        "Hệ thống được đóng gói hoàn chỉnh thành ứng dụng web Flask tích hợp API, kiểm tra miền giá trị, nhận diện trạng thái bất định (Uncertainty) "
        "và xử lý phân loại hàng loạt theo lô (Batch CSV) với khả năng tái lập 100%."
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 1: BỐI CẢNH VÀ CÂU HỎI NGHIÊN CỨU
    # ============================================================
    doc.add_heading("1. Bối Cảnh Thực Tiễn và Câu Hỏi Nghiên Cứu", level=1)
    
    doc.add_heading("1.1. Bối cảnh thực tiễn", level=2)
    doc.add_paragraph(
        "Trong ngành chế biến và xuất khẩu lúa gạo, việc lẫn tạp giữa các giống lúa có giá trị kinh tế khác nhau là thách thức lớn đối với "
        "các nhà máy xay xát và cơ sở đóng gói. Giống gạo Cammeo và Osmancik là hai giống lúa phổ biến có giá thành chênh lệch và yêu cầu bảo quản riêng biệt. "
        "Phương pháp truyền thống phân loại thủ công bằng mắt thường của công nhân giám định bộc lộ nhiều nhược điểm: tốn kém nhân lực, "
        "tốc độ xử lý chậm và phụ thuộc vào cảm tính, dễ xảy ra sai sót khi mắt bị mỏi mệt sau nhiều giờ làm việc liên tục."
    )
    doc.add_paragraph(
        "Để giải quyết vấn đề trên, hệ thống thị giác máy tính kết hợp máy học được đề xuất nhằm tự động hóa công đoạn sàng lọc sơ bộ. "
        "Hình ảnh từng hạt gạo trên băng chuyền được chụp bằng camera công nghiệp độ phân giải cao, sau đó các thuật toán xử lý ảnh tách nền "
        "và trích xuất 7 đặc trưng hình học. Nhiệm vụ của mô hình học máy là nhận diện chính xác từng hạt là Cammeo hay Osmancik chỉ trong vài mili-giây."
    )

    doc.add_heading("1.2. Câu hỏi nghiên cứu trọng tâm", level=2)
    doc.add_paragraph(
        "Đề tài tập trung giải quyết câu hỏi mang tính nguyên lý của Bài 6 (Cây quyết định, Cắt tỉa và Rừng ngẫu nhiên):\n"
        "👉 'Một mô hình Cây quyết định tối giản (đã cắt tỉa ccp_alpha) có đủ độ chính xác và khả năng diễn giải cho người vận hành, "
        "hay bắt buộc phải triển khai Rừng ngẫu nhiên (Random Forest) với chi phí tính toán cao hơn và mất tính minh bạch?'"
    )

    doc.add_heading("1.3. Đặc tả bài toán và Tiêu chí thành công", level=2)
    tbl_prob = doc.add_table(rows=1, cols=2)
    format_table(
        tbl_prob,
        [2.2, 4.3],
        ["Thành phần bài toán", "Mô tả chi tiết"],
        [
            ["Loại bài toán", "Phân loại nhị phân có giám sát (Binary Classification)."],
            ["Đơn vị quan sát", "Một hạt gạo duy nhất."],
            ["Không gian đầu vào", "7 biến hình học thực (Area, Perimeter, Major_Axis, Minor_Axis, Eccentricity, Convex_Area, Extent)."],
            ["Đầu ra mục tiêu", "Nhãn giống lúa: Cammeo hoặc Osmancik, kèm phân phối xác suất và quy tắc cây."],
            ["Tiêu chí thành công tối thiểu", "Vượt Baseline đa số (57.22%), đạt Accuracy & F1 > 90% trên tập Test độc lập, pipeline tái lập sạch, Web app kiểm soát lỗi nhập liệu."]
        ]
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 2: DỮ LIỆU THỰC TẾ VÀ PHÂN TÍCH KHÁM PHÁ (EDA)
    # ============================================================
    doc.add_heading("2. Dữ Liệu Thực Tế và Phân Tích Khám Phá (EDA)", level=1)

    doc.add_heading("2.1. Nguồn gốc, giấy phép và tính hợp pháp", level=2)
    doc.add_paragraph(
        "Bộ dữ liệu sử dụng là Rice (Cammeo and Osmancik) được công bố trên UCI Machine Learning Repository "
        "(DOI: 10.24432/C5MW4Z). Dữ liệu được trích xuất từ công trình khoa học của Cinar & Koklu (2019). "
        "Bộ dữ liệu được phát hành theo giấy phép Creative Commons Attribution 4.0 International (CC BY 4.0), "
        "hoàn toàn hợp pháp cho mục đích nghiên cứu, học thuật và phát triển công nghệ phi thương mại."
    )

    doc.add_heading("2.2. Từ điển dữ liệu (Data Dictionary)", level=2)
    tbl_dict = doc.add_table(rows=1, cols=5)
    format_table(
        tbl_dict,
        [1.3, 0.9, 0.9, 0.9, 2.5],
        ["Tên biến", "Kiểu dữ liệu", "Đơn vị", "Vai trò", "Ý nghĩa hình thái học"],
        [
            ["Area", "Integer", "pixel²", "Feature", "Số lượng pixel nằm trong vùng hạt gạo."],
            ["Perimeter", "Real", "pixel", "Chiều dài đường bao xung quanh hạt."],
            ["Major_Axis_Length", "Real", "pixel", "Chiều dài trục chính (khoảng cách dài nhất của elip tương đương)."],
            ["Minor_Axis_Length", "Real", "pixel", "Chiều dài trục phụ (chiều rộng ngắn nhất của elip tương đương)."],
            ["Eccentricity", "Real", "Vô thứ nguyên", "Độ lệch tâm hình elip (0: tròn tuyệt đối, tiến tới 1: dẹp dài)."],
            ["Convex_Area", "Integer", "pixel²", "Diện tích của đa giác lồi nhỏ nhất bao trọn hạt gạo."],
            ["Extent", "Real", "Tỷ lệ [0, 1]", "Tỷ lệ giữa diện tích hạt và diện tích khung hình chữ nhật bao quanh."],
            ["Class", "Nominal", "—", "Target", "Nhãn phân loại thực tế: Cammeo hoặc Osmancik."]
        ]
    )

    doc.add_paragraph()
    doc.add_heading("2.3. Báo cáo chất lượng dữ liệu ban đầu", level=2)
    doc.add_paragraph(
        "Qua kiểm tra tự động trên toàn bộ 3,810 mẫu hạt gạo:\n"
        "• Dữ liệu khuyết thiếu (Missing values): 0 mẫu (100% bản ghi đầy đủ).\n"
        "• Dữ liệu trùng lặp hoàn toàn (Exact duplicates): 0 bản ghi.\n"
        "• Phân bố lớp mục tiêu: Osmancik chiếm 2,180 mẫu (57.2%), Cammeo chiếm 1,630 mẫu (42.8%). "
        "Tỷ lệ lớp mất cân bằng nhẹ (~1.34 : 1), không đòi hỏi các kỹ thuật oversampling phức tạp nhưng cần lưu ý khi đánh giá thước đo F1-score và ROC-AUC."
    )

    doc.add_heading("2.4. Phân tích phân phối và tương quan đặc trưng", level=2)
    doc.add_paragraph(
        "Hình 1 biểu diễn phân phối của 7 đặc trưng giữa hai giống lúa. Có thể nhận thấy rõ: Giống Cammeo có xu hướng dài hơn và to hơn, "
        "thể hiện qua các chỉ số Area, Perimeter, Major_Axis_Length và Convex_Area dịch chuyển hẳn về phía giá trị cao hơn so với Osmancik. "
        "Ngược lại, giống Osmancik có hình dáng hạt tròn trịa, đều đặn hơn."
    )

    add_image_with_caption(doc, "01_distributions.png", "Hình 1: Phân phối xác suất 7 đặc trưng hình học giữa hai giống gạo Cammeo và Osmancik.", width_inch=6.2)
    add_image_with_caption(doc, "02_boxplots.png", "Hình 2: Biểu đồ hộp (Boxplot) thể hiện độ phân tán và điểm ngoại lai của từng đặc trưng.", width_inch=6.2)
    add_image_with_caption(doc, "03_correlation_matrix.png", "Hình 3: Ma trận tương quan Pearson giữa 7 đặc trưng hình học.", width_inch=5.2)

    doc.add_paragraph(
        "Ma trận tương quan tại Hình 3 chỉ ra rằng nhóm biến kích thước (Area, Perimeter, Major_Axis_Length, Convex_Area) có tương quan tuyến tính rất cao (r > 0.95). "
        "Đây là lý do thuật toán Cây quyết định rất phù hợp, bởi cây không bị ảnh hưởng tiêu cực bởi hiện tượng đa cộng tuyến (multicollinearity) "
        "như hồi quy tuyến tính hay logistic."
    )

    add_image_with_caption(doc, "04_scatter_major_minor.png", "Hình 4: Đồ thị phân tán giữa Trục chính (Major_Axis) và Trục phụ (Minor_Axis).", width_inch=5.2)

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 3: THIẾT KẾ PHÂN CHIA VÀ CHỐNG RÒ RỈ DỮ LIỆU
    # ============================================================
    doc.add_heading("3. Thiết Kế Phân Chia Dữ Liệu và Ngăn Chặn Rò Rỉ (Leakage)", level=1)
    
    doc.add_heading("3.1. Phân chia tầng (Stratified Splitting)", level=2)
    doc.add_paragraph(
        "Nhằm đảm bảo đánh giá khách quan và kiểm soát phương sai, bộ dữ liệu được chia theo cơ chế phân tầng (Stratified 60/20/20):\n"
        "1. Tập Huấn Luyện (Train Set - 60% = 2,286 mẫu): Dùng để học cây, tính toán bảng cost-complexity path và huấn luyện rừng.\n"
        "2. Tập Kiểm Tra Xác Thực (Validation Set - 20% = 762 mẫu): Dùng chọn siêu tham số max_depth, chọn mô hình tối ưu và phân tích sai số.\n"
        "3. Tập Kiểm Thử Cuối Cùng (Test Set - 20% = 762 mẫu): BẢO VỆ TUYỆT ĐỐI. Tuyệt đối không mở tập test trong suốt quá trình thử nghiệm."
    )

    add_callout(
        doc,
        "Tập Test chỉ được nạp và đánh giá đúng DUY NHẤT một lần ở khâu cuối cùng của báo cáo. "
        "Không có bất kỳ quyết định chọn tham số hay tiền xử lý nào được nhìn trước dữ liệu của tập Test. "
        "Tất cả giá trị thống kê miền [min, max] dùng kiểm tra tính hợp lệ trên Web App đều được trích xuất từ tập Train.",
        "QUY TẮC BẢO VỆ TẬP TEST NGHIÊM NGẶT"
    )

    doc.add_heading("3.2. Mô hình đối chứng cơ sở (Dummy Baseline)", level=2)
    doc.add_paragraph(
        "Mô hình Baseline được thiết lập bằng DummyClassifier với chiến lược 'most_frequent'. Baseline luôn đưa ra dự đoán cố định "
        "là giống lúa chiếm đa số (Osmancik). Kết quả Baseline trên tập Validation:\n"
        "• Baseline Accuracy: 57.22%\n"
        "• Baseline F1-score (cho lớp Cammeo): 0.0000\n"
        "Mục tiêu bắt buộc của các mô hình học máy đề xuất là phải vượt qua ngưỡng 57.22% này một cách vượt trội."
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 4: PHƯƠNG PHÁP HỌC MÁY VÀ THIẾT KẾ THÍ NGHIỆM
    # ============================================================
    doc.add_heading("4. Phương Pháp Học Máy và Thiết Kế Thí Nghiệm", level=1)

    doc.add_heading("4.1. Cơ sở lý thuyết thuật toán", level=2)
    doc.add_paragraph(
        "1. Cây quyết định (Decision Tree): Sử dụng chỉ số Gini Impurity để tìm điểm phân chia tối ưu tại từng nút. Cây có ưu điểm trực quan, "
        "dễ giải thích thành các quy tắc if-else, nhưng rất dễ bị quá khớp nếu phát triển tự do.\n"
        "2. Cắt tỉa Cost-Complexity Pruning (ccp_alpha): Cơ chế cắt tỉa hậu kỳ dựa trên hàm chi phí phạt độ phức tạp:\n"
        "     R_alpha(T) = R(T) + alpha * |T|\n"
        "Trong đó R(T) là tỷ lệ phân loại sai của cây T, |T| là số lượng nút lá, và alpha là tham số phạt độ cồng kềnh của cây. "
        "Khi alpha tăng, cây sẽ bị tỉa bớt các nhánh phụ, giúp giảm phương sai (variance) và tăng khả năng tổng quát hóa.\n"
        "3. Rừng ngẫu nhiên (Random Forest): Thuật toán học kết hợp (Bagging Ensemble) tập hợp 100 cây quyết định được huấn luyện "
        "trên các mẫu Bootstrap khác nhau kết hợp chọn ngẫu nhiên tập con đặc trưng (Feature Subsampling). Rừng giúp triệt tiêu hiện tượng "
        "overfitting của cây đơn lẻ và cung cấp xác suất dự đoán mượt mà."
    )

    doc.add_heading("4.2. Bốn thí nghiệm bắt buộc theo hồ sơ nhiệm vụ", level=2)
    tbl_exp = doc.add_table(rows=1, cols=3)
    format_table(
        tbl_exp,
        [0.8, 2.0, 3.7],
        ["TN", "Tên thí nghiệm", "Mục tiêu và bằng chứng cần đạt"],
        [
            ["TN 1", "Đường hiệu năng theo max_depth", "Khảo sát hiện tượng overfitting khi cây tăng độ sâu từ 1 đến 20 trên Train và Val."],
            ["TN 2", "Pruning bằng ccp_alpha trên CV", "Tìm effective alphas, chạy 5-fold Stratified Cross-Validation để chọn alpha tối ưu."],
            ["TN 3", "So sánh Cây đơn vs Rừng ngẫu nhiên", "Đối chứng toàn diện về Accuracy, F1, ROC-AUC, thời gian huấn luyện và tính giải thích."],
            ["TN 4", "Phân tích vùng đặc trưng sai số", "Phân tích cụ thể các mẫu bị dự đoán sai, so sánh phân phối đặc trưng đúng vs sai, khảo sát độ tin cậy xác suất."]
        ]
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 5: KẾT QUẢ THỰC NGHIỆM VÀ PHÂN TÍCH
    # ============================================================
    doc.add_heading("5. Kết Quả Thực Nghiệm và Đánh Giá Chi Tiết", level=1)

    doc.add_heading("5.1. Kết quả Thí nghiệm 1: Hiệu năng theo độ sâu max_depth", level=2)
    doc.add_paragraph(
        "Kết quả tại Hình 5 minh chứng rõ nét hiện tượng Overfitting:\n"
        "• Train Accuracy tăng đều đặn từ 92.9% (depth=1) lên 100.0% tuyệt đối khi depth >= 16.\n"
        "• Tuy nhiên, Validation Accuracy chỉ đạt cao nhất ở độ sâu rất nông (depth=1, 2) đạt 93.70%, sau đó suy giảm dần xuống còn ~89.1% ở độ sâu 17.\n"
        "Điều này khẳng định việc cho cây phát triển sâu tối đa sẽ khiến mô hình ghi nhớ cả những nhiễu ngẫu nhiên trong tập train."
    )
    add_image_with_caption(doc, "05_depth_curve.png", "Hình 5: Đường hiệu năng Accuracy theo max_depth trên tập Train và Validation.", width_inch=5.8)

    doc.add_heading("5.2. Kết quả Thí nghiệm 2: Cắt tỉa Cost-Complexity (ccp_alpha)", level=2)
    doc.add_paragraph(
        "Từ cây đầy đủ, hàm cost_complexity_pruning_path xác định được 70 ngưỡng alpha khả dĩ. Tiến hành 5-fold CV trên tập Train (Hình 6):\n"
        "• Ngưỡng tối ưu tìm được: ccp_alpha = 0.008024.\n"
        "• Độ chính xác trung bình trên 5 folds: 92.48% ± 1.59%.\n"
        "• Cây sau khi cắt tỉa (Hình 7) chỉ còn lại cấu trúc cực kỳ tối giản với nút gốc chia theo Major_Axis_Length <= 191.05 pixel. "
        "Đây là kết quả mang tính ứng dụng thực tiễn rất cao: Chỉ cần đo một đặc trưng chiều dài, hệ thống đã phân loại đúng hơn 93% số mẫu."
    )
    add_image_with_caption(doc, "06_pruning_curve.png", "Hình 6: Điểm số Cross-Validation theo tham số cắt tỉa ccp_alpha.", width_inch=5.8)
    add_image_with_caption(doc, "07_decision_tree.png", "Hình 7: Cây quyết định tối giản sau cắt tỉa ccp_alpha (rõ ràng, minh bạch, dễ giải thích).", width_inch=6.2)

    doc.add_heading("5.3. Kết quả Thí nghiệm 3: So sánh Cây quyết định vs Rừng ngẫu nhiên", level=2)
    doc.add_paragraph(
        "Bảng 1 tổng hợp kết quả đối chứng giữa các phương pháp trên tập Validation và tập Test độc lập (chạy 1 lần cuối cùng):"
    )

    tbl_res = doc.add_table(rows=1, cols=6)
    format_table(
        tbl_res,
        [2.0, 0.9, 0.9, 0.9, 0.9, 0.9],
        ["Mô hình", "Val Acc", "Val F1", "Val AUC", "Test Acc", "Test AUC"],
        [
            ["Baseline (Most Frequent)", "57.22%", "0.00%", "N/A", "57.22%", "0.5000"],
            ["Decision Tree (Pruned)", "93.70%", "92.77%", "0.9596", "91.47%", "0.9344"],
            ["Random Forest (100 Trees)", "93.44%", "92.28%", "0.9741", "91.60%", "0.9709"]
        ]
    )

    doc.add_paragraph()
    add_image_with_caption(doc, "10_confusion_matrices.png", "Hình 8: Ma trận nhầm lẫn (Confusion Matrix) trên tập Validation.", width_inch=6.2)
    add_image_with_caption(doc, "11_roc_curves.png", "Hình 9: Đường cong ROC so sánh năng lực phân tách giữa Cây và Rừng trên Validation.", width_inch=5.5)
    add_image_with_caption(doc, "13_test_confusion_matrices.png", "Hình 10: Ma trận nhầm lẫn cuối cùng trên tập Test độc lập (chạy duy nhất 1 lần).", width_inch=6.2)

    doc.add_paragraph(
        "Nhận xét kết quả:\n"
        "• Về độ chính xác: Cả Cây quyết định cắt tỉa (91.47%) và Rừng ngẫu nhiên (91.60%) đều vượt trội hoàn toàn so với Baseline (57.22%) "
        "và xấp xỉ ngang nhau trên tập Test.\n"
        "• Về ROC-AUC: Rừng ngẫu nhiên vượt trội hơn đáng kể (0.9709 so với 0.9344 của Cây đơn) nhờ khả năng tổng hợp nhiều cây con "
        "cho ra phổ xác suất liên tục mượt mà, hạn chế bậc thang nhảy cóc.\n"
        "• Về thời gian huấn luyện: Cây quyết định chỉ mất 0.036 giây, nhanh gấp gần 10 lần so với Rừng ngẫu nhiên (0.320 giây)."
    )

    add_image_with_caption(doc, "08_feature_importance_tree.png", "Hình 11: Mức độ quan trọng đặc trưng (Feature Importance) của Cây quyết định.", width_inch=5.0)
    add_image_with_caption(doc, "09_feature_importance_rf.png", "Hình 12: Mức độ quan trọng đặc trưng của Rừng ngẫu nhiên.", width_inch=5.0)

    doc.add_heading("5.4. Kết quả Thí nghiệm 4: Phân tích lỗi chi tiết (Error Analysis)", level=2)
    doc.add_paragraph(
        "Trên tập Validation (762 mẫu), mô hình Rừng ngẫu nhiên chỉ mắc phải 50 lỗi (tỷ lệ lỗi 6.6%). Phân bố lỗi hai chiều rất cân bằng:\n"
        "• 27 hạt giống Cammeo bị phân loại nhầm thành Osmancik (False Negative cho Cammeo).\n"
        "• 23 hạt giống Osmancik bị phân loại nhầm thành Cammeo (False Positive cho Cammeo).\n\n"
        "So sánh trung bình các đặc trưng giữa nhóm dự đoán Đúng và nhóm dự đoán Sai:"
    )

    tbl_err = doc.add_table(rows=1, cols=4)
    format_table(
        tbl_err,
        [2.2, 1.4, 1.4, 1.5],
        ["Đặc trưng", "Trung bình mẫu ĐÚNG", "Trung bình mẫu SAI", "Độ chênh lệch (Δ)"],
        [
            ["Area (pixel²)", "12,722.04", "12,773.52", "+51.48"],
            ["Perimeter (pixel)", "454.94", "458.54", "+3.60"],
            ["Major_Axis_Length (pixel)", "188.83", "191.75", "+2.92"],
            ["Minor_Axis_Length (pixel)", "86.68", "85.84", "-0.85"],
            ["Eccentricity", "0.89", "0.89", "+0.00"],
            ["Convex_Area (pixel²)", "13,010.52", "13,058.96", "+48.44"],
            ["Extent", "0.66", "0.67", "+0.01"]
        ]
    )

    doc.add_paragraph()
    add_image_with_caption(doc, "12_error_analysis.png", "Hình 13: So sánh phân phối đặc trưng giữa các mẫu dự đoán đúng và sai.", width_inch=6.2)

    doc.add_paragraph(
        "Phân tích bản chất lỗi:\n"
        "1. Vùng giao thoa hình thái học: Các mẫu sai có Major_Axis trung bình là 191.75 pixel — đúng ngay tại ngưỡng cắt quyết định (191.05 pixel). "
        "Đây là những hạt gạo lai tạo hoặc biến dị nhẹ, khiến kích thước nằm đúng tại ranh giới phân tách tự nhiên của hai giống.\n"
        "2. Độ tin cậy dự đoán (Model Confidence): Trong khi các mẫu đúng có xác suất trung bình rất cao (94.37%), thì ở các mẫu bị đoán sai, "
        "xác suất tin cậy trung bình sụt giảm mạnh chỉ còn 78.28% (nhiều mẫu xác suất chỉ đạt 51% - 55%). "
        "Phát hiện này dẫn tới giải pháp thiết kế cờ cảnh báo bất định trên ứng dụng web."
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 6: ỨNG DỤNG WEB VÀ TÍNH NĂNG MỞ RỘNG
    # ============================================================
    doc.add_heading("6. Ứng Dụng Web và Triển Khai Phục Vụ (Serving)", level=1)

    doc.add_heading("6.1. Kiến trúc hệ thống và Tách biệt Train - Serving", level=2)
    doc.add_paragraph(
        "Ứng dụng web tuân thủ chặt chẽ tiêu chuẩn kỹ thuật đề ra:\n"
        "• Tách biệt hoàn toàn offline training và online serving: Mô hình tối ưu được đóng gói thành tệp `models/final_model.pkl`. "
        "Khi khởi động, server Flask chỉ nạp mô hình một lần duy nhất vào bộ nhớ, không huấn luyện lại ở mỗi request.\n"
        "• Backend: Sử dụng Python Flask cung cấp các REST API endpoint chuẩn.\n"
        "• Frontend: Xây dựng bằng Bootstrap 5 hỗ trợ giao diện đáp ứng (responsive), hiển thị trực quan tỷ lệ xác suất và trạng thái tin cậy."
    )

    doc.add_heading("6.2. Kiểm tra tính hợp lệ dữ liệu đầu vào (Input Validation)", level=2)
    doc.add_paragraph(
        "Hệ thống kiểm soát chặt chẽ 3 tầng validation trước khi chuyển dữ liệu vào mô hình:\n"
        "1. Kiểm tra thiếu trường: Trả về mã lỗi HTTP 400 nếu người dùng bỏ trống bất kỳ trường nào trong 7 đặc trưng.\n"
        "2. Kiểm tra định dạng số: Từ chối chuỗi ký tự hoặc định dạng không hợp lệ.\n"
        "3. Kiểm tra miền quan sát [min, max]: Nếu giá trị vượt ra ngoài ngưỡng thống kê khi huấn luyện, hệ thống không âm thầm thay thế "
        "mà cảnh báo rõ ràng cho người dùng biết kết quả có thể không đáng tin cậy."
    )

    doc.add_heading("6.3. Tính năng mở rộng 1: Cảnh báo bất định (Uncertainty Detection)", level=2)
    doc.add_paragraph(
        "Theo yêu cầu mở rộng tại Mục 4 của đề tài, hệ thống tự động kiểm tra khoảng cách xác suất tới ngưỡng 0.5. "
        "Khi xác suất phân lớp rơi vào dải bất định [0.40, 0.60], giao diện hiển thị hộp cảnh báo màu đỏ với thông điệp: "
        "'⚠️ CẢNH BÁO BẤT ĐỊNH: Xác suất phân lớp xấp xỉ 50/50. Theo quy định sử dụng có trách nhiệm, hạt gạo này nằm ở vùng ranh giới hình thái, "
        "không được tự động đóng gói mà phải chuyển qua chuyên viên giám định trực tiếp.'"
    )

    doc.add_heading("6.4. Tính năng mở rộng 2: Xử lý theo lô (Batch CSV Prediction)", level=2)
    doc.add_paragraph(
        "Nhằm phục vụ vận hành thực tế tại nhà máy, ứng dụng cung cấp tab 'Phân loại theo lô':\n"
        "• Cho phép tải lên tệp CSV chứa danh sách hàng trăm mẫu hạt.\n"
        "• Hệ thống tự động phân loại, kiểm tra từng dòng, thống kê tổng số hạt Cammeo, Osmancik và đếm số lượng mẫu bất định.\n"
        "• Cung cấp nút xuất toàn bộ kết quả ra tệp CSV (`ket_qua_phan_loai_lo.csv`) kèm cờ Uncertainty_Flag."
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 7: ĐẠO ĐỨC, GIỚI HẠN VÀ SỬ DỤNG CÓ TRÁCH NHIỆM
    # ============================================================
    doc.add_heading("7. Đạo Đức, Giới Hạn và Sử Dụng Có Trách Nhiệm", level=1)

    doc.add_paragraph(
        "1. Giới hạn công nghệ: Mô hình được huấn luyện trên điều kiện ánh sáng và độ phân giải ảnh tiêu chuẩn của bộ dữ liệu gốc. "
        "Nếu camera tại cơ sở sản xuất có sai lệch về tiêu cự, kích thước pixel hoặc góc chiếu sáng, đặc trưng trích xuất sẽ bị sai lệch hệ thống.\n"
        "2. Phạm vi giống lúa: Hệ thống chỉ được thiết kế để phân biệt giữa Cammeo và Osmancik. Tuyệt đối không sử dụng mô hình "
        "cho các giống lúa khác (như Jasmine, ST25 hay Basmati) vì mô hình sẽ bắt buộc gán vào 1 trong 2 nhãn có sẵn một cách sai lệch.\n"
        "3. Trách nhiệm pháp lý: Kết quả máy học chỉ đóng vai trò hỗ trợ sàng lọc cơ học sơ bộ, không có giá trị thay thế chứng thư "
        "giám định chất lượng giống của cơ quan nhà nước có thẩm quyền.\n"
        "4. Quyền riêng tư: Dữ liệu hình ảnh hạt gạo là dữ liệu nông sản thuần túy, không chứa dữ liệu cá nhân hay thông tin định danh người dùng."
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 8: KẾT LUẬN VÀ TRẢ LỜI CÂU HỎI VẤN ĐÁP
    # ============================================================
    doc.add_heading("8. Kết Luận và Trả Lời Câu Hỏi Nghiên Cứu", level=1)

    doc.add_heading("8.1. Trả lời câu hỏi nghiên cứu", level=2)
    doc.add_paragraph(
        "Dựa trên toàn bộ kết quả thực nghiệm của 4 thí nghiệm:\n"
        "✓ Cây quyết định tối giản (đã cắt tỉa ccp_alpha) ĐỦ CHÍNH XÁC cho bài toán này (Test Accuracy 91.47% so với 91.60% của Rừng ngẫu nhiên). "
        "Đồng thời, cây đơn lẻ mang lại ưu thế tuyệt đối về tính minh bạch: toàn bộ luật phân loại có thể viết thành đúng 2 câu lệnh điều kiện đơn giản, "
        "giúp kỹ thuật viên tại nhà máy kiểm tra và hiểu tường tận lý do phân loại.\n"
        "✓ Rừng ngẫu nhiên nên được ưu tiên khi hệ thống cần kết nối với dây chuyền tự động đòi hỏi chỉ số ROC-AUC cao (0.9709) và phân bổ xác suất mượt mà "
        "để thiết lập các ngưỡng điều khiển van khí nén gạt hạt gạo."
    )

    doc.add_heading("8.2. Chuẩn bị câu hỏi vấn đáp gợi ý của giảng viên", level=2)
    doc.add_paragraph(
        "1. Câu hỏi: 'ccp_alpha làm cây thay đổi thế nào?'\n"
        "   Trả lời: ccp_alpha điều khiển mức độ phạt độ phức tạp theo công thức R_alpha(T) = R(T) + alpha*|T|. Khi alpha = 0, cây giữ nguyên kích thước đầy đủ (dễ overfitting). "
        "Khi alpha tăng dần, các nhánh có độ suy giảm tạp chất (impurity decrease) nhỏ hơn chi phí phạt sẽ bị cắt bỏ, thu nhỏ số lượng lá |T|, "
        "làm cây nông hơn và tổng quát hóa tốt hơn.\n\n"
        "2. Câu hỏi: 'Feature importance của Rừng ngẫu nhiên có thiên lệch gì?'\n"
        "   Trả lời: Mức độ quan trọng đặc trưng dựa trên MDI (Mean Decrease in Impurity) trong Rừng ngẫu nhiên có xu hướng thiên vị (bias) "
        "cho các đặc trưng liên tục có nhiều giá trị phân biệt (high cardinality) hoặc các biến có tương quan mạnh với nhau. "
        "Trong bài toán này, các biến kích thước (Area, Convex_Area, Perimeter) chia sẻ mức độ quan trọng cho nhau do đa cộng tuyến cao.\n\n"
        "3. Câu hỏi: 'Khi nào chọn cây thay vì rừng?'\n"
        "   Trả lời: Chọn Cây quyết định khi: (1) Cần khả năng giải thích tuyệt đối cho người dùng không có chuyên môn IT; "
        "(2) Thiết bị nhúng/phần cứng biên bị giới hạn tài nguyên tính toán và bộ nhớ cực thấp; "
        "(3) Khi dữ liệu có ranh giới phân tách trực giao rõ ràng và cây đơn đã đạt độ chính xác tiệm cận rừng."
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 9: PHÂN CÔNG ĐÓNG GÓP VÀ NHẬT KÝ 6 TUẦN
    # ============================================================
    doc.add_heading("9. Phân Công Đóng Góp Nhóm và Nhật Ký 6 Tuần", level=1)

    doc.add_heading("9.1. Bảng phân công trách nhiệm (Tỷ lệ đóng góp 50% - 50%)", level=2)
    tbl_team = doc.add_table(rows=1, cols=4)
    format_table(
        tbl_team,
        [1.5, 2.5, 1.8, 0.7],
        ["Thành viên", "Trách nhiệm chính đảm nhiệm", "Minh chứng Git / Artifact", "Tỷ lệ"],
        [
            ["Thành viên 1", "Xử lý dữ liệu, kiểm soát chống rò rỉ, xây dựng baseline, thuật toán cắt tỉa ccp_alpha (TN1 & TN2), lập trình REST API backend Flask.", "Commit data_processing.py, train.py, app.py core", "50%"],
            ["Thành viên 2", "Huấn luyện Random Forest, thiết kế thực nghiệm TN3 & TN4, phân tích lỗi chi tiết, xây dựng giao diện frontend web app & tính năng batch CSV.", "Commit evaluate.py, templates/index.html, figures generation", "50%"]
        ]
    )

    doc.add_paragraph()
    doc.add_heading("9.2. Nhật ký tiến độ 6 tuần", level=2)
    tbl_log = doc.add_table(rows=1, cols=4)
    format_table(
        tbl_log,
        [0.8, 1.6, 2.8, 1.3],
        ["Tuần", "Trọng tâm công việc", "Kết quả đầu ra đạt được", "Người phụ trách"],
        [
            ["Tuần 1", "Xác nhận đề tài, bối cảnh", "Hoàn thành Project Brief, thiết lập cấu trúc repo và Data README.", "Cả 2 thành viên"],
            ["Tuần 2", "Dữ liệu & Baseline", "Làm sạch dữ liệu, EDA, chốt Stratified Split 60/20/20, xây dựng Dummy Baseline.", "Thành viên 1"],
            ["Tuần 3", "Mô hình chính & Pruning", "Hoàn thiện TN1 (depth curve) và TN2 (ccp_alpha pruning 5-fold CV).", "Thành viên 1 & 2"],
            ["Tuần 4", "Đánh giá & Phân tích lỗi", "Thực hiện TN3 (so sánh cây vs rừng), TN4 (phân tích lỗi), đánh giá test 1 lần duy nhất.", "Thành viên 2"],
            ["Tuần 5", "Tích hợp Ứng dụng Web", "Hoàn thiện Web Flask phục vụ model offline, validation 3 tầng, thêm tính năng Batch CSV.", "Thành viên 1 & 2"],
            ["Tuần 6", "Hoàn thiện & Thử nghiệm sạch", "Thử nghiệm tái lập từ môi trường sạch, hoàn thiện Báo cáo DOCX, Slide 11 trang.", "Cả 2 thành viên"]
        ]
    )

    doc.add_paragraph()

    # ============================================================
    # CHƯƠNG 10: TÀI LIỆU THAM KHẢO VÀ PHỤ LỤC TÁI LẬP
    # ============================================================
    doc.add_heading("10. Tài Liệu Tham Khảo và Hướng Dẫn Tái Lập", level=1)

    doc.add_heading("10.1. Tài liệu tham khảo", level=2)
    doc.add_paragraph(
        "[1] Tài liệu phương pháp: Chuyên đề Cây quyết định, cắt tỉa (Pruning) và Rừng ngẫu nhiên (Random Forest).\n"
        "[2] Cinar, I. and Koklu, M. (2019). Classification of Rice Varieties Using Artificial Intelligence Methods. "
        "International Journal of Intelligent Systems and Applications in Engineering, vol. 7, no. 3, pp. 188-194. DOI: 10.18201/ijisae.2019355381.\n"
        "[3] UCI Machine Learning Repository: Rice (Cammeo and Osmancik) Dataset. DOI: 10.24432/C5MW4Z. URL: https://archive.ics.uci.edu/dataset/545/rice+cammeo+and+osmancik.\n"
        "[4] Scikit-learn Documentation: Decision Trees & Post pruning decision trees with cost complexity pruning. "
        "URL: https://scikit-learn.org/stable/modules/tree.html#minimal-cost-complexity-pruning."
    )

    doc.add_heading("10.2. Phụ lục: Hướng dẫn tái lập toàn bộ pipeline từ máy mới", level=2)
    doc.add_paragraph(
        "Để tái lập toàn bộ kết quả từ môi trường máy tính hoàn toàn sạch:\n"
        "Bước 1: Tạo và kích hoạt môi trường ảo Python:\n"
        "   python -m venv venv\n"
        "   venv\\Scripts\\activate  (trên Windows) hoặc source venv/bin/activate (trên Linux/macOS)\n\n"
        "Bước 2: Cài đặt các thư viện phụ thuộc:\n"
        "   pip install -r requirements.txt\n\n"
        "Bước 3: Chạy lần lượt các bước của pipeline:\n"
        "   python src/data_processing.py  # Xử lý, EDA và chia dữ liệu (tạo figures 01-04)\n"
        "   python src/train.py            # Huấn luyện baseline, TN1, TN2, TN3 (tạo figures 05-09)\n"
        "   python src/evaluate.py         # Đánh giá, TN4, test set 1 lần (tạo figures 10-13)\n\n"
        "Bước 4: Khởi động ứng dụng web:\n"
        "   python app/app.py\n"
        "   Mở trình duyệt truy cập: http://localhost:5000"
    )

    doc.save(DOCX_PATH)
    print(f"Report DOCX saved successfully to: {DOCX_PATH}")


if __name__ == "__main__":
    generate_document()
