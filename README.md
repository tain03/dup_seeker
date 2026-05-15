# 🌸 DUP-SEEKER: Premium Excel Image Integrity Suite

**DUP-SEEKER** là một ứng dụng Desktop chuyên nghiệp được thiết kế để tự động hóa việc kiểm tra tính toàn vẹn của hình ảnh trong hàng loạt báo cáo Excel. Ứng dụng giúp phát hiện các hình ảnh bị sử dụng lặp lại (duplicate) giữa các báo cáo, đảm bảo tính trung thực và chất lượng của dữ liệu báo cáo kỹ thuật.

![License](https://img.shields.io/badge/license-MIT-blue.svg)
![Python](https://img.shields.io/badge/python-3.10+-blue.svg)
![UI](https://img.shields.io/badge/UI-CustomTkinter-indigo.svg)

## ✨ Tính năng nổi bật

- 🔍 **Quét Ảnh Chính Xác:** Sử dụng thuật toán băm MD5 để so sánh nội dung hình ảnh, đảm bảo phát hiện chính xác 100% các ảnh trùng lặp.
- 📂 **Xử Lý Hàng Loạt:** Hỗ trợ kéo và thả (Drag & Drop) hàng chục file Excel cùng lúc.
- 📍 **Định Vị Chính Xác:** Chỉ rõ vị trí ảnh trùng tại **File nào, Sheet nào và Ô (Cell) nào**. Hỗ trợ cả ảnh lơ lửng (Floating Anchors).
- ⚙️ **Tùy Biến Linh Hoạt:** Cho phép loại trừ các Sheet mẫu (Template) để tránh báo lỗi nhầm.
- 🏢 **Giao Diện Đẳng Cấp:** Thiết kế theo phong cách **Seoul Corporate Blue** (Samsung/LG style), tối ưu cho môi trường doanh nghiệp.
- 🚀 **Hiệu Năng Cao:** Xử lý đa luồng (Multi-threading), không gây treo ứng dụng khi xử lý hàng nghìn ảnh.

## 🛠️ Cài đặt

Nếu bạn muốn chạy từ mã nguồn:

1. Cài đặt các thư viện cần thiết:
```bash
pip install customtkinter pillow tkinterdnd2-universal
```

2. Chạy ứng dụng:
```bash
python app.py
```

## 📦 Đóng gói (Release)

Ứng dụng có thể được đóng gói thành file `.exe` duy nhất bằng PyInstaller:
```bash
pyinstaller --noconsole --onedir --icon="assets/app.ico" --name "DUP-SEEKER" app.py
```

## 📖 Hướng dẫn sử dụng

1. **Chọn file:** Kéo thả các file Excel vào ứng dụng hoặc nhấn nút "SELECT DATA SOURCE".
2. **Cấu hình:** Nhập tên các Sheet cần bỏ qua vào ô "EXCLUSION SETS".
3. **Thực hiện:** Nhấn "EXECUTE AUDIT" và đợi trong giây lát.
4. **Xử lý:** Xem kết quả, nhấn "VIEW FILE" để mở trực tiếp file Excel chứa lỗi để chỉnh sửa.

---
**Developed by:** [ductai.nguyen](https://github.com/ductai-nguyen)
