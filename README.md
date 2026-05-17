# 🚀 DUP-SEEKER: Premium Excel Image Integrity Suite

**DUP-SEEKER** là một giải pháp Desktop cao cấp (Enterprise-Grade) được thiết kế để tự động hóa việc kiểm tra tính toàn vẹn của hình ảnh trong hàng loạt báo cáo Excel (.xlsx). Bằng cách kết hợp **Cơ chế phân tích XML lõi cực nhanh**, **Mạng lưới Băm Ảnh kép (MD5 + pHash)**, **Bộ nhớ đệm SQLite vĩnh viễn** và **COM Automation sâu với Microsoft Excel**, ứng dụng giúp phát hiện và định vị chính xác 100% các hình ảnh trùng lặp được tái sử dụng sai quy định trong môi trường báo cáo doanh nghiệp.

![GitHub release (latest by date)](https://img.shields.io/badge/release-v2.1--Final-emerald.svg?style=flat-square)
![Python](https://img.shields.io/badge/python-3.12-sky.svg?style=flat-square)
![UI](https://img.shields.io/badge/UI-CustomTkinter-indigo.svg?style=flat-square)
![License](https://img.shields.io/badge/license-MIT-blue.svg?style=flat-square)

---

## 📸 Kiến trúc Luồng Dữ liệu (System Workflow)

Mô hình dưới đây mô tả hành trình xử lý từ khi kéo thả tệp tin Excel vào ứng dụng cho đến khi định vị chính xác vị trí ảnh trùng lặp trên Microsoft Excel:

```mermaid
graph TD
    %% Định nghĩa các lớp node
    A[📂 Drop Files / Folders] --> B(🧵 Background Worker Thread)
    B --> C{🗄️ SQLite MD5 permanent Cache}
    
    %% Quá trình Phân tích Zip/XML
    C -->|Chưa có Hash| D[⚡ Fast Pass 1: Raw ZIP & XML Parser]
    D --> E[📍 Extract Anchor Coordinates: Sheet/Cell/Row/Col]
    E --> F[✨ MD5 Hash of raw media data]
    
    %% Bộ lọc trùng lặp
    F --> G{📊 Duplicate Filter}
    G -->|Trùng lặp| H[🖥️ Hierarchical Treeview List]
    G -->|Nghi vấn| I[👁️ Pass 2: AI Vision pHash Engine]
    
    %% Cơ chế Hiển thị
    H --> J[🔍 Parent-Child Sets with Zebra Striping]
    H --> K[📊 Right Panel: Interactive Spreadsheet Grid]
    
    %%COM Deep Linking
    K -->|Click OPEN| L[🖨️ Windows COM Automation]
    L --> M[🟢 Auto Open workbook, active target sheet & focus Cell address]
```

---

## ✨ Các Tính Năng Đỉnh Cao (Feature Spotlight)

### 📊 1. Bảng lưới Lớp chi tiết Nhóm Trùng (Right Panel Spreadsheet Grid)
*   **Trực quan hóa dạng bảng:** Hiển thị danh sách tất cả các vị trí trùng lặp của `SET` dưới dạng bảng lưới (Spreadsheet Grid) mini sắc nét với đầy đủ kẻ sọc Zebra sang trọng, không bị cắt cụt chữ.
*   **Hành động tức thì:** Nút **`OPEN`** nhỏ gọn cho từng dòng giúp mở trực tiếp tệp Excel tương ứng chỉ với một click.

### 🛡️ 2. Thuật toán Chống vỡ giao diện (Filename Capping)
*   **Tự động co giãn an toàn:** Cột hiển thị tên tệp tin được giới hạn tối đa `380px`, tự động rút gọn bằng dấu `...` nếu tên tệp quá dài, bảo vệ 100% không gian hiển thị của các cột thông tin phụ (`Sheet`, `Cell Address`, `Occurrences`).
*   **Căn chỉnh tối giản:** Cấu trúc [Tên File] (Căn trái) [Sheet] (Căn trái) [Cell] (Căn giữa) chuẩn mực như một báo cáo tài chính chuyên nghiệp.

### 🔌 3. COM Deep-Linking (Liên kết sâu tới Excel)
*   Không chỉ mở file thông thường, DUP-SEEKER sử dụng cơ chế **Windows COM Interface** để:
    1. Khởi động hoặc móc nối trực tiếp vào phiên làm việc đang chạy của **Microsoft Excel**.
    2. Kích hoạt đúng Workbook và chuyển tiếp đến đúng **Sheet** chứa ảnh trùng.
    3. **Tự động chọn (Select) và di chuyển màn hình** tập trung vào đúng ô tọa độ (`Cell Address`) chứa hình ảnh đó!

### 🎨 4. Tích hợp Biểu tượng Hệ thống (Taskbar & Titlebar Icon)
*   Sử dụng Windows **AppUserModelID** để đăng ký tiến trình riêng biệt với Windows Shell. Nhờ đó, biểu tượng `assets/app.ico` luôn hiển thị đồng bộ dưới Taskbar và góc tiêu đề cửa sổ, mang lại cảm giác của một sản phẩm thương mại cao cấp.

### ⚡ 5. Tự động Phóng to Toàn màn hình (Startup Maximized)
*   Tự động phóng to tối đa cửa sổ ngay khi khởi động bằng cơ chế **Trễ luồng giao diện (150ms)**, giải quyết triệt để lỗi CustomTkinter tự động co rút lại do tính toán DPI Scaling của Windows.

---

## 🛠️ Hướng dẫn Cài đặt & Chạy từ nguồn

Nếu bạn là nhà phát triển và muốn chạy ứng dụng từ mã nguồn Python:

### 1. Cài đặt các thư viện phụ thuộc:
```bash
pip install customtkinter pillow tkinterdnd2-universal imagehash xlsxwriter pywin32
```

### 2. Khởi chạy ứng dụng:
```bash
python app.py
```

---

## 📦 Quy trình Đóng gói thành tệp .EXE

Ứng dụng được đóng gói chuyên nghiệp bằng **PyInstaller** dựa trên cấu hình tệp `.spec` đi kèm, đảm bảo đóng gói đầy đủ các thư viện kéo thả phức tạp (`tkinterdnd2`) và các tài nguyên giao diện:

```bash
# Thực hiện biên dịch ứng dụng sang thư mục phân phối Windows
pyinstaller DUP-SEEKER.spec --noconfirm
```

> [!TIP]
> Sản phẩm sau khi đóng gói sẽ nằm tại thư mục `dist/DUP-SEEKER`. Để chuyển giao, bạn chỉ cần nén (Zip) thư mục này lại thành **`DUP-SEEKER_v2.1_Final.zip`** là có thể gửi cho người dùng cuối chạy trực tiếp không cần cài đặt Python.

---

## 👥 Nhóm phát triển & Bản quyền
*   **Tác giả:** [ductai.nguyen](https://github.com/ductai-nguyen)
*   **Bản quyền:** Phát hành dưới giấy phép MIT License. Bảo lưu mọi quyền đối với bộ nhận diện thương hiệu DUP-SEEKER.
