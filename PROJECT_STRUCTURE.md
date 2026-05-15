# DUP-SEEKER: Technical Specification & Context

Tài liệu này cung cấp cái nhìn chi tiết về kiến trúc và logic nghiệp vụ của dự án DUP-SEEKER để phục vụ việc bảo trì hoặc phát triển bởi AI.

## 1. Tổng quan hệ thống
- **Mục tiêu:** Phát hiện ảnh trùng lặp trong file `.xlsx` bằng cách trích xuất dữ liệu media thô và so sánh MD5 hash.
- **Công nghệ:** Python 3.12, CustomTkinter (UI), TkinterDnD (Drag & Drop), Zipfile/XML (Excel Parsing).

## 2. Kiến trúc mã nguồn
Dự án được chia thành 2 phần chính trong cùng một tệp hoặc module:

### A. Backend Logic (ExcelProcessor)
- **Cơ chế hoạt động:** File Excel được xử lý như một tệp ZIP.
- **Luồng trích xuất:**
    1. Đọc `xl/workbook.xml` để lấy tên các Sheet.
    2. Duy cập `xl/worksheets/_rels/sheet{n}.xml.rels` để tìm liên kết đến tệp `drawing{n}.xml`.
    3. Đọc `xl/drawings/drawing{n}.xml` để xác định vị trí ảnh (Anchors: `twoCellAnchor`, `oneCellAnchor`, `absoluteAnchor`).
    4. Ánh xạ `rId` từ tệp drawing sang tệp media thực tế trong `xl/media/`.
- **Hệ thống băm:** Sử dụng `hashlib.md5(img_data).hexdigest()` để tạo định danh duy nhất cho nội dung ảnh.

### B. Frontend Architecture (DuplicateApp)
- **Framework:** `customtkinter` với chế độ "Light Mode".
- **Responsive:** Tính năng tính toán kích thước cửa sổ dựa trên `winfo_screenwidth/height`.
- **Đa luồng:** Sử dụng `threading.Thread` cho hàm `scan_logic` để giữ cho giao diện luôn phản hồi (Main loop không bị block).
- **Trạng thái (State):** Quản lý qua các thuộc tính của class như `self.selected_files`, `self.is_scanning`.

## 3. Quy chuẩn UI/UX (Seoul Corporate Blue)
- **Primary Color:** `#0284C7` (Azure Blue).
- **Sidebar:** `#0F172A` (Deep Slate).
- **Corner Radius:** `12px` - `16px`.
- **Font:** Ưu tiên Sans-serif (Segoe UI, Inter).

## 4. Lưu ý khi phát triển tiếp
- **Excel Formats:** Hiện tại chỉ hỗ trợ `.xlsx`.
- **Image Formats:** Hỗ trợ tất cả định dạng ảnh mà Excel cho phép nhúng (png, jpeg, emf, wmf).
- **Xử lý lỗi:** Đã có cơ chế `try-except` cho từng file để đảm bảo một file lỗi không làm dừng toàn bộ tiến trình quét.

## 5. Quy trình đóng gói
Sử dụng PyInstaller với các tham số:
- `--onedir` để tối ưu tốc độ load.
- `--add-data` để đính kèm thư mục `assets` và các binary của `tkinterdnd2`, `customtkinter`.
