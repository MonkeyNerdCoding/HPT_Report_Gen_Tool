# 📊 OracleHC Report Generator

Công cụ hỗ trợ tạo báo cáo Oracle Health Check tự động với giao diện người dùng (GUI) thân thiện.

---

## 🛠 Hướng dẫn Phát triển (Development)

Trong quá trình phát triển và chỉnh sửa mã nguồn, bạn có thể khởi chạy ứng dụng trực tiếp từ Python để kiểm tra nhanh:

```bash
python gui.py
```

---

## 🚀 Hướng dẫn Đóng gói (Build Release)

Khi ứng dụng đã sẵn sàng để phát hành, hãy sử dụng **PyInstaller** để đóng gói thành file thực thi (`.exe`). Quá trình này được chia thành 2 bước:

### 1. Chạy Test Build (Dạng thư mục)
Dùng để kiểm tra xem các tệp tin assets và mapping đã được nhúng đúng chưa mà không mất nhiều thời gian nén.
```bash
pyinstaller --noconfirm --onedir --windowed --name "OracleHC Report Generator" --add-data "assets\tachnen_hpt.png;assets" --add-data "mapping\report_mapping.yaml;mapping" gui.py
```

### 2. Build Final (Dạng một file duy nhất)
Dùng để tạo ra bản phát hành cuối cùng cho người dùng, gọn nhẹ và chuyên nghiệp.
```bash
pyinstaller --clean --noconfirm --onefile --windowed --name "OracleHC Report Generator" --add-data "assets\tachnen_hpt.png;assets" --add-data "mapping\report_mapping.yaml;mapping" gui.py
```

---

## 📂 Vị trí File sau khi Build

Sau khi quá trình **Build Final** hoàn tất, bạn có thể tìm thấy file chạy duy nhất tại đường dẫn sau:

> `dist\OracleHC Report Generator.exe`

---

