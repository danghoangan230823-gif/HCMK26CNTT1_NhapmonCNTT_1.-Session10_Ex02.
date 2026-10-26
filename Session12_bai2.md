# Bài 2 : Chuẩn bị hồ sơ dự án trước khi triển khai hệ thống
## - **Phần 1. Phân tích:**
Cấu trúc gồm 4 mục Heading 1 và đúng 10 mục Heading 2:
```
SRS – PTIT-LEARNING (Title)
├── 1. Giới thiệu (Heading 1)
│   ├── 1.1 Mục đích (Heading 2)
│   ├── 1.2 Phạm vi (Heading 2)
│   └── 1.3 Thuật ngữ và viết tắt (Heading 2)
├── 2. Mô tả tổng quan (Heading 1)
│   ├── 2.1 Đối tượng sử dụng (Heading 2)
│   └── 2.2 Ràng buộc và giả định (Heading 2)
├── 3. Yêu cầu chức năng (Heading 1)
│   ├── 3.1 FR-01: Tạo và lên lịch lớp học (Heading 2)
│   ├── 3.2 FR-02: Tham gia lớp học (Heading 2)
│   └── 3.3 FR-03: Điểm danh và chuyên cần (Heading 2)
└── 4. Yêu cầu phi chức năng (Heading 1)
    ├── 4.1 Hiệu năng (Heading 2)
    └── 4.2 Bảo mật (Heading 2)
```

**Thống kê:** 4 Heading 1, đúng 10 Heading 2 (3 + 2 + 3 + 2).

**Kiểm soát:** mở Word → View → Navigation Pane để kiểm tra đủ 10 mục cấp 2 và kéo thả đổi thứ tự khi cần.
## - **Phần 2. Triển khai:** File báo cáo hoàn chỉnh.
### 1. Đặc tả yêu cầu (Word - SRS)
Cấu trúc Navigation Pane: Sử dụng 5 Heading cấp 1 (1. Giới thiệu, 2. Tổng quan hệ thống, 3. Yêu cầu chức năng, 4. Yêu cầu phi chức năng, 5. Kế hoạch triển khai) kết hợp với các Heading cấp 2 để phân rã chi tiết module "Học trực tuyến".

Yêu cầu chức năng: Tập trung đặc tả phân vùng Lớp học ảo (quản lý phòng học, điểm danh tự động, phát video/livestream độ trễ thấp).

### 2. Quản lý & Lọc dữ liệu (Excel)
Cấu trúc dữ liệu: File gồm 200 bản ghi sinh viên/khóa học với dòng tiêu đề (Header Row) duy nhất ở hàng đầu tiên, không chứa dòng tổng cộng xen kẽ giữa dữ liệu.

Thao tác lọc (Ctrl + Shift + L): Lọc nhanh danh sách sinh viên thuộc khoa CNTT có điểm chuyên cần dưới 7.0.

Xử lý lỗi lọc tên: Sử dụng hàm =TRIM(C2) hoặc =TRIM(SUBSTITUTE(C2, CHAR(160), " ")) để loại bỏ các khoảng trắng thừa/ký tự ẩn giúp bộ lọc nhận diện chính xác dữ liệu.

### 3. Thiết kế giao diện (PowerPoint)
Combo màu tối: Nền Navy (#0A192F), chữ trắng (#FFFFFF) kết hợp Accent cam (#FF6B35) làm điểm nhấn, đảm bảo độ tương phản cao (Contrast Ratio).

Bố cục 3 vùng:

Vùng 1 (Header): Tiêu đề và logo nhận diện.

Vùng 2 (Body): Flowchart lộ trình học tập trực tuyến.

Vùng 3 (Footer): Thông tin phiên bản và số trang.
## Phần 3. Kiểm thử: Xử lý lỗi bộ lọc dữ liệu (AutoFilter)
Vấn đề: Khi sử dụng Ctrl + Shift + L để lọc tên sinh viên nhưng không hiển thị kết quả (dù dữ liệu có trên bảng).

Nguyên nhân: Dữ liệu tồn tại các khoảng trắng thừa (leading/trailing spaces) hoặc ký tự ngắt dòng ẩn (CHAR(160)).

Cách khắc phục:

### 1.Tạo cột phụ sử dụng hàm chuẩn hóa:
=TRIM(SUBSTITUTE(C2, CHAR(160), " "))
### 2.
Dùng cột phụ này để tiến hành AutoFilter hoặc copy đè giá trị đã làm sạch (Paste as Values) vào cột chính.