# K26CNTT2_IT108_Session10_TH
Bài Tập Tổng Hợp
# Hồ Sơ Đồ Án Hệ Thống Quản Lý Thư Viện Trường (Session 11 - BTVN)

## Phần 1. Cấu Trúc Đặc Tả SRS (Heading Styles & Auto TOC)

### 1. Khung tiêu đề SRS phân cấp 3 cấp hoàn chỉnh

Dưới đây là cây phân cấp các tiêu đề được gán nhãn Heading Styles trong file Word `SRS_ThuVien.docx`:

# 1. Giới thiệu dự án (Heading 1)
## 1.1. Mục tiêu (Heading 2)
## 1.2. Đối tượng sử dụng (Heading 2)

# 2. Chức năng hệ thống (Heading 1)
## 2.1. Quản lý mượn sách (Heading 2)
### 2.1.1. Luồng mượn sách qua thẻ sinh viên (Heading 3)
## 2.2. Quản lý trả sách (Heading 2)

# 3. Yêu cầu bảo mật (Heading 1)

---

### 2. Mô tả nội dung câu Functional & Non-functional Requirements

- **Yêu cầu chức năng (Functional Requirements - Đặt dưới mục 2. Chức năng hệ thống):**
  - Hệ thống phải cho phép độc giả tra cứu danh mục sách theo thời gian thực và tự động cập nhật trạng thái khả dụng của từng đầu sách.
  - Ban quản lý thư viện có thể trích xuất báo cáo thống kê danh sách độc giả mượn quá hạn và tự động tính toán tiền phạt theo chu kỳ hàng tuần.

- **Yêu cầu phi chức năng (Non-functional Requirements - Đặt dưới mục 3. Yêu cầu bảo mật):**
  - Mọi dữ liệu tài khoản và thông tin thẻ sinh viên phải được mã hóa mã nguồn theo chuẩn mã hóa AES-256 để chống nguy cơ rò rỉ thông tin cá nhân.
  - Hệ thống phải đảm bảo thời gian hoạt động (Uptime) đạt 99.9% và thực hiện tự động sao lưu dữ liệu bộ nhớ đệm (Cache Backup) vào lúc 00:00 mỗi ngày.

---

### 3. Quy trình chèn Auto TOC & Cập nhật số trang (Update Field)

1. **Chèn Auto TOC:** Tại trang đầu tiên của file Word, chọn thẻ `References` ➔ Chọn `Table of Contents` ➔ Chọn `Automatic Table 1`.
2. **Đẩy trang & Cập nhật số trang:** 
   - Đặt con trỏ chuột trước mục `3. Yêu cầu bảo mật`, nhấn `Ctrl + Enter` để ngắt trang (Page Break), đẩy mục 3 sang trang tiếp theo.
   - Nhấp chuột phải vào bảng Mục lục tự động ➔ Chọn **`Update Field`** ➔ Chọn **`Update entire table`** ➔ Nhấn **OK**. Số trang của Mục 3 sẽ ngay lập tức được hệ thống cập nhật tự động từ trang 1 sang trang 2 chuẩn xác.

---

## Phần 2. Chuẩn Hóa Dữ Liệu & Thao Tác AutoFilter

### 1. Bảng phân tích 3 lỗi vi phạm 5 Quy tắc Vàng Database

| Lỗi # | Vị trí (Hàng) | Quy tắc bị vi phạm | Mô tả lỗi chi tiết & Hậu quả |
| :---: | :---: | :--- | :--- |
| **1** | Hàng 1 (A1:E1) | **Quy tắc 1:** Hàng đầu tiên phải là Header, không merge ô | Gộp ô Merge Cell làm tiêu đề lớn khiến Excel nhận diện sai hàng Header khi bật AutoFilter. |
| **2** | Hàng 5 | **Quy tắc 4:** Không có dòng hoặc cột trống | Dòng trống hoàn toàn làm đứt gãy vùng quét (Range) của bộ lọc AutoFilter, bỏ sót dữ liệu phía dưới. |
| **3** | Hàng 6 | **Quy tắc 5:** Dòng tổng cộng không nằm trong vùng dữ liệu thô | Nhập dòng "Tổng cộng" xen giữa làm sai lệch bản chất dòng dữ liệu độc lập (Record) của sinh viên. |

---

### 2. Bảng dữ liệu chuẩn hóa tái thiết kế

Bảng dữ liệu `PhieuMuon` sau khi loại bỏ ô gộp, xóa dòng trống và chuyển dòng tính tổng ra ngoài:

| MSSV | Họ Tên | Mã Sách | Ngày Mượn | Tình Trạng |
| :--- | :--- | :--- | :--- | :--- |
| SV001 | Nguyễn Văn An | S001 | 01/10/2026 | Đang mượn |
| SV002 | Trần Thị Bình | S002 | 03/10/2026 | Quá hạn |
| SV003 | Lê Văn Cường | S003 | 10/10/2026 | Đã trả |
| SV004 | Phạm Thu Hà | S099 | 12/10/2026 | Đang mượn |

*(Dòng "Tổng số phiếu mượn: 4" được đặt ở hàng 7, nằm hoàn toàn bên ngoài bảng dữ liệu thô)*.

---

### 3. Định nghĩa Record và Field trong bảng dữ liệu

- **Ví dụ 1 Record (Bản ghi / Dòng dữ liệu):** 
  - *Dòng 3:* `SV002 | Trần Thị Bình | S002 | 03/10/2026 | Quá hạn`
  - *Ý nghĩa:* Tập hợp toàn bộ các thuộc tính mô tả cho đúng 1 sự kiện mượn sách cụ thể của sinh viên Trần Thị Bình.
- **Ví dụ 1 Field (Trường / Cột dữ liệu):** 
  - *Cột E (`Tình Trạng`):* `[Đang mượn, Quá hạn, Đã trả, Đang mượn]`
  - *Ý nghĩa:* Trữ lượng thông tin chứa duy nhất 1 loại thuộc tính về trạng thái mượn sách của tất cả các bản ghi trong hệ thống.

---

### 4. Thao tác bật AutoFilter và lọc trạng thái "Quá hạn"

1. Bôi đen vùng bảng dữ liệu đã sửa từ Hàng 1 (`MSSV` đến `Tình Trạng`).
2. Nhấn tổ hợp phím tắt **`Ctrl + Shift + L`** để kích hoạt tính năng **AutoFilter**.
3. Nhấp chuột vào nút mũi tên lọc tại cột **Tình Trạng** (Cột E).
4. Bỏ tích chọn `(Select All)`, chỉ tích chọn duy nhất giá trị **`Quá hạn`** ➔ Nhấn **OK**.
5. **Kết quả:** Hệ thống lọc ra đúng 1 bản ghi của sinh viên **Trần Thị Bình (SV002)**.

---

## Phần 3. Nối Dữ Liệu Bằng Hàm VLOOKUP

### 1. Công thức VLOOKUP tự động lấy Tên Sách

Thêm cột `Tên Sách` vào Cột D (bên cạnh cột Mã Sách). Nhập công thức tại ô D2:

```excel
=VLOOKUP(C2, DanhMucSach!A$2:C$4, 2, FALSE)

```

Tham số:

C2: Giá trị tra cứu (Mã Sách S001).

DanhMucSach!A$2:C$4: Vùng bảng tra cứu cố định bằng dấu $.

2: Lấy dữ liệu ở Cột 2 (Tên Sách).

FALSE: Tìm kiếm chính xác 100%.

2. Phân tích lỗi đối với Mã Sách S099 (Hàng 5)
Loại lỗi trả về: Lỗi #N/A (Not Available).

Nguyên nhân: Mã sách S099 ở hàng 5 không tồn tại trong vùng dữ liệu tra cứu DanhMucSach!A$2:C$4. Khi không tìm thấy giá trị khớp chính xác 100%, hàm VLOOKUP với tham số FALSE sẽ lập tức báo lỗi #N/A.

3. Công thức hoàn chỉnh bọc IFERROR
Viết lại công thức bọc hàm IFERROR tại ô D2 và copy xuống dòng D5:
```exel
=IFERROR(VLOOKUP(C2, DanhMucSach!A$2:C$4, 2, FALSE), "Không tìm thấy sách")
```

Kết quả xử lý: Tại dòng chứa mã S099, ô Tên Sách sẽ hiển thị chuỗi chữ "Không tìm thấy sách" thay vì hiển thị lỗi rác #N/A.

Phần 4. Slide Báo Cáo Tiến Độ (Rule 10/20/30)
1. Thiết kế nội dung 1 Slide báo cáo duy nhất
Assertion Title (Tiêu đề Kết luận):

"Tỷ lệ mượn sách đạt 75% hoạt động ổn định với 1 trường hợp phát sinh quá hạn cần xử lý"

Content Area (Con số & Kết luận nổi bật):

4 Phiếu mượn tổng cộng được ghi nhận trong tháng 10.

50% sinh viên đang mượn (2 phiếu), 25% đã trả (1 phiếu), 25% quá hạn (1 phiếu).

100% dữ liệu phiếu mượn và danh mục sách đã được làm sạch và nối thành công qua hàm VLOOKUP.

Loại biểu đồ minh họa & Lý do lựa chọn:
 Biểu đồ bánh khuyết (Donut Chart).

Lý do: Biểu đồ Donut thể hiện trực quan mối quan hệ tỷ lệ phần trăm giữa các bộ phận với tổng thể (2 phiếu Đang mượn / 1 phiếu Quá hạn / 1 phiếu Đã trả trên Tổng 4 phiếu) mà không làm rối mắt người xem.

2. Phân tích áp dụng Rule 10/20/30
Quy tắc Rule 10/20/30: Khuyên dùng tối đa 10 slides, thời gian trình bày trong 20 phút, và cỡ chữ nhỏ nhất là 30pt.

Tỷ trọng áp dụng thực tế:

Slide tiến độ này chiếm 1/10 slides (10% dung lượng bộ slide báo cáo đồ án).

Thời gian trình bày dự kiến cho slide này là 2 phút (trên tổng 20 phút thuyết trình đồ án trước lớp), vừa đủ để làm nổi bật kết quả xử lý dữ liệu và phương hướng tiếp theo.






