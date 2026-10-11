# 🌾 CẨM NANG HƯỚNG DẪN THUYẾT TRÌNH & BẢO VỆ ĐỒ ÁN
## Project 17: Phân Loại Hai Giống Gạo (Cammeo vs. Osmancik) Bằng Học Máy

---

## 🎯 PHẦN 1: BÀI NÓI MẪU 2 PHÚT (KHI ĐỨNG TRƯỚC THẦY CÔ)

> *"Kính thưa quý Thầy/Cô, em xin phép trình bày về Đồ án Học máy cơ bản: **Phân loại giống lúa Cammeo và Osmancik từ hình dạng hạt gạo**.*
>
> * **Bối cảnh thực tế:** Trong dây chuyền chế biến nông sản xuất khẩu, việc phân loại thủ công bằng mắt thường rất chậm, tốn nhân lực và dễ sai sót. Mục tiêu của đồ án là xây dựng một hệ thống thị giác máy tính kết hợp học máy, cho phép phân loại tức thì 2 giống lúa từ 7 thông số hình học (như diện tích, chu vi, chiều dài trục chính...).
> * **Phương pháp thực nghiệm:** Nhóm đã triển khai thực nghiệm so sánh giữa mô hình tối giản dễ giải thích là **Cây quyết định (Decision Tree)** và mô hình tổ hợp mạnh mẽ là **Rừng ngẫu nhiên (Random Forest)** trên tập dữ liệu 3,810 hạt gạo (UCI Dataset).
> * **Kết quả nổi bật:** 
>   1. Cây quyết định sau khi cắt tỉa cành bằng kỹ thuật Cost-Complexity Pruning ($ccp\_alpha$) đạt độ chính xác **91.47%** trên tập Test độc lập. Mô hình này có tính minh bạch cực cao: chỉ cần nhìn vào quy tắc chiều dài trục chính lớn hơn $191.05\text{px}$ là đã phân loại đúng đại đa số hạt gạo.
>   2. Rừng ngẫu nhiên đạt độ chính xác **91.60%** và chỉ số ROC-AUC vượt trội **0.9709**, cung cấp xác suất dự đoán mượt mà.
> * **Triển khai ứng dụng:** Nhóm đã đóng gói mô hình thành một Web App hoàn chỉnh bằng Flask, hỗ trợ dự đoán từng hạt, xử lý file CSV theo lô, đồng thời tích hợp tính năng **Cảnh báo bất định (Uncertainty Detection)** đối với các mẫu nằm ở vùng ranh giới $50/50$ để chuyển giao cho chuyên viên thẩm định. Em xin mời Thầy/Cô theo dõi phần demo trực tiếp."*

---

## 💡 PHẦN 2: Ý NGHĨA 7 THÔNG SỐ HÌNH HỌC (DỄ NHỚ)

1. **Area (Diện tích):** Số pixel bao phủ bề mặt hạt gạo (Cammeo thường có diện tích lớn hơn Osmancik).
2. **Perimeter (Chu vi):** Chiều dài đường viền quanh hạt gạo.
3. **Major Axis Length (Chiều dài trục chính):** Độ dài lớn nhất từ đầu đến đuôi hạt gạo. **(Đây là đặc trưng quan trọng nhất để phân biệt 2 giống)**.
4. **Minor Axis Length (Chiều rộng trục phụ):** Bề ngang lớn nhất của hạt gạo.
5. **Eccentricity (Độ lệch tâm / độ thuôn dài):** Đo xem hạt gạo dẹt hay tròn (giá trị từ 0 đến 1; càng gần 1 hạt càng thuôn dài, Cammeo thuôn dài hơn).
6. **Convex Area (Diện tích đa giác bao lồi):** Diện tích phần đa giác lồi nhỏ nhất chứa trọn hạt gạo.
7. **Extent (Tỷ lệ lấp đầy khung chữ nhật bao quanh):** Diện tích hạt chia cho diện tích hình chữ nhật nhỏ nhất bao quanh nó.

---

## ❓ PHẦN 3: BỘ 10 CÂU HỎI THẦY CÔ HAY HỎI NHẤT & CÂU TRẢ LỜI MẪU

### Câu 1: Tại sao lại chia tập dữ liệu theo tỷ lệ 60% Train / 20% Val / 20% Test mà không phải 80/20?
* **Trả lời:** Chia 3 tập để phục vụ đúng quy trình học máy chuẩn:
  * **Train (60%):** Dùng để huấn luyện các trọng số của mô hình.
  * **Validation (20%):** Dùng để dò tìm siêu tham số tối ưu (như độ sâu `max_depth` và hệ số cắt tỉa `ccp_alpha`) mà không làm lộ dữ liệu kiểm thử.
  * **Test (20%):** Đề thi bí mật, chỉ mở ra đánh giá đúng 1 lần cuối cùng sau khi đã chốt mô hình để đảm bảo kết quả khách quan nhất.

### Câu 2: Stratified Split là gì? Tại sao phải dùng Stratified?
* **Trả lời:** Stratified Split là kỹ thuật phân tầng khi chia dữ liệu, đảm bảo tỷ lệ giữa 2 giống gạo (khoảng 43% Cammeo và 57% Osmancik) được giữ nguyên vẹn như nhau ở cả 3 tập Train, Val và Test. Điều này tránh tình trạng một tập nào đó bị lệch lớp hoặc thiếu hụt mẫu của một giống.

### Câu 3: Kỹ thuật cắt tỉa ccp_alpha (Cost-Complexity Pruning) hoạt động thế nào?
* **Trả lời:** Nếu để cây quyết định phát triển tự nhiên, cây sẽ mọc rất nhiều nhánh con cho đến khi phân loại đúng 100% tập Train, dẫn tới hiện tượng **Overfitting (học vẹt)**. Kỹ thuật `ccp_alpha` thêm một lượng phạt tỷ lệ với số lượng lá của cây:
  $$Cost = R(T) + \alpha \times |T|$$
  (với $R(T)$ là tỷ lệ lỗi, $|T|$ là số nút lá). Giá trị $\alpha$ càng lớn thì cây càng bị phạt nặng và bị cắt bớt nhánh rườm rà. Nhóm em dùng 5-fold Cross-Validation trên tập Train để chọn ra giá trị $\alpha = 0.008024$ tối ưu nhất.

### Câu 4: Baseline trong bài này là gì? Tại sao phải có Baseline?
* **Trả lời:** Nhóm sử dụng mô hình Baseline là `DummyClassifier(strategy="most_frequent")` — nghĩa là không học gì cả, bất kỳ hạt gạo nào cũng đoán là giống phổ biến nhất (Osmancik). 
* Độ chính xác của Baseline là **57.22%**. Baseline là thước đo tối thiểu: một mô hình AI thực sự có giá trị chỉ khi vượt trội rõ rệt so với con số 57.22% này (mô hình của nhóm đạt **91.60%**).

### Câu 5: Điểm khác nhau giữa Decision Tree và Random Forest trong bài toán này là gì?
* **Trả lời:**
  * **Decision Tree:** Là một cây đơn lẻ, thời gian chạy cực nhanh ($0.013s$) và có **khả năng diễn giải tuyệt vời** (đưa ra luật `if-else` rõ ràng).
  * **Random Forest:** Là một tập hợp 100 cây quyết định, mỗi cây học trên một tập mẫu con ngẫu nhiên (Bootstrap) và chọn đặc trưng ngẫu nhiên. Random Forest hoạt động như "hội đồng chuyên gia bỏ phiếu", cho đường cong ROC-AUC mượt hơn ($0.9709$), hạn chế phương sai (giảm rủi ro đoán sai ở các hạt dị tật).

### Câu 6: Feature nào đóng vai trò quan trọng nhất trong việc phân loại?
* **Trả lời:** Đặc trưng quan trọng nhất là `Major_Axis_Length` (chiều dài trục chính), tiếp theo là `Perimeter` (chu vi) và `Area` (diện tích). Giống Cammeo có chiều dài trung bình vượt trội so với Osmancik.

### Câu 7: Ma trận nhầm lẫn (Confusion Matrix) trên tập Test nói lên điều gì?
* **Trả lời:** Trên 762 mẫu kiểm thử:
  * Cammeo: Đúng 284 mẫu, nhầm sang Osmancik 42 mẫu.
  * Osmancik: Đúng 414 mẫu, nhầm sang Cammeo 22 mẫu.
  * Tỷ lệ chính xác tổng thể đạt $91.60\%$.

### Câu 8: Nhóm đã làm phân tích lỗi (Error Analysis) như thế nào? Tại sao mô hình đoán sai?
* **Trả lời:** Nhóm đã lọc ra 50 hạt gạo bị đoán sai và vẽ biểu đồ phân phối đối chứng. Kết quả cho thấy các hạt bị đoán sai đều là những hạt có kích thước rơi vào vùng giao thoa (chiều dài xấp xỉ $191\text{px}$), hoặc hạt bị lép, vỡ đầu khiến hình dạng của Cammeo bị ngắn lại như Osmancik. Xác suất tin cậy của mô hình ở các hạt bị sai cũng thấp hơn hẳn so với hạt đúng ($0.78$ vs $0.94$).

### Câu 9: Tính năng "Cảnh báo bất định (Uncertainty Detection)" trên Web App có ý nghĩa gì trong thực tế?
* **Trả lời:** Đây là nguyên tắc đạo đức và an toàn AI trong công nghiệp. Khi một hạt gạo có xác suất dự đoán lơ lửng quanh mức $40\% - 60\%$ (AI đang phân vân), hệ thống Web không tự tiện gán nhãn mà phát tín hiệu cờ vàng/đỏ: *"Mẫu bất định, cần chuyên viên kiểm tra trực tiếp"*. Điều này ngăn ngừa việc đóng gói nhầm giống gạo cao cấp.

### Câu 10: Ứng dụng Web này có chạy offline được không hay cần kết nối Internet?
* **Trả lời:** Ứng dụng chạy hoàn toàn offline $100\%$. File mô hình `final_model.pkl` đã được đóng gói sẵn và nạp thẳng vào bộ nhớ RAM khi bật server, không phụ thuộc vào Internet hay bất kỳ dịch vụ đám mây bên ngoài nào.
