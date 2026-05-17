# 🚀 DUP-SEEKER: Premium Excel Image Integrity Suite

> [🇺🇸 English / English Edition](README.md)

---

**DUP-SEEKER** là một giải pháp phần mềm Desktop chuyên dụng cấp doanh nghiệp giúp tự động hóa quá trình quét, phát hiện và kiểm tra ảnh trùng lặp trong hàng loạt báo cáo Excel (`.xlsx`), bảo vệ tuyệt đối tính trung thực của dữ liệu và bố cục báo cáo kỹ thuật.

![GitHub release (latest by date)](https://img.shields.io/badge/release-v2.1--Final-emerald.svg?style=flat-square)
![Python](https://img.shields.io/badge/python-3.12-sky.svg?style=flat-square)
![UI](https://img.shields.io/badge/UI-CustomTkinter-indigo.svg?style=flat-square)
![License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat-square)

---

## 🌟 Điểm Nhất Vượt Trội (Tính Năng Đáng Giá Nhất)

*   **🔌 Liên kết Sâu COM & Tự động Định vị:** Không chỉ mở file thông thường, DUP-SEEKER kết nối trực tiếp vào hệ thống **Microsoft Excel** bằng Windows COM. Khi nhấp nút **`OPEN`**, phần mềm tự động kích hoạt Excel, mở đúng **Sheet**, và **bôi đen/tự động cuộn màn hình** tập trung vào chính xác **Tọa độ ô (Cell Address)** chứa bức ảnh trùng lặp đó!
*   **⚡ Động cơ Băm Kép Siêu Tốc (MD5 + pHash):** Mạng lưới đối khớp ảnh kép kết hợp cùng **Cơ sở dữ liệu đệm SQLite vĩnh viễn** giúp quét và phân tích hàng nghìn bức ảnh chỉ trong vài giây mà không cần tính toán lại dữ liệu cũ.
*   **💎 Giao diện Doanh nghiệp Đỉnh cao:** Giao diện CustomTkinter sang trọng với tông xanh Slate-Blue chủ đạo:
    *   **Cây thư mục phân cấp kẻ sọc Zebra** với cột Tên file được giới hạn trần `380px` chống vỡ giao diện.
    *   **Bảng lưới Spreadsheet Grid** hiển thị danh sách ảnh trùng ở góc phải sắc nét, dễ nhìn với các nút mở trực tiếp tiện lợi.
    *   **Đồng bộ Icon dưới Taskbar & Khởi động Maximize cố định** mang lại cảm giác của ứng dụng thương mại hoàn thiện.

---

## 🛠️ Hướng dẫn Cài đặt & Sử dụng

> [!NOTE]
> **Không cần cài đặt!** DUP-SEEKER hoạt động hoàn toàn độc lập dưới dạng tệp chạy di động (Portable).

### Cách 1: Sử dụng bản EXE đóng gói sẵn (Cho người dùng cuối)
1. Tải về và giải nén tệp tin **`DUP-SEEKER_v2.1_Final.zip`**.
2. Nhấp đúp chuột vào tệp [DUP-SEEKER.exe](file:///e:/project/check_dup_img_excel/dist/DUP-SEEKER/DUP-SEEKER.exe) để khởi chạy phần mềm ngay lập tức.

### Cách 2: Chạy từ mã nguồn Python (Cho lập trình viên)
1. Cài đặt các thư viện phụ thuộc:
   ```bash
   pip install customtkinter pillow tkinterdnd2-universal imagehash xlsxwriter pywin32
   ```
2. Khởi chạy ứng dụng:
   ```bash
   python app.py
   ```

---

## 👥 Nhóm phát triển & Bản quyền
*   **Tác giả:** [ductai.nguyen](https://github.com/ductai-nguyen)
*   **Bản quyền:** Phát hành dưới giấy phép MIT License. Bảo lưu mọi quyền đối với bộ nhận diện thương hiệu DUP-SEEKER.
