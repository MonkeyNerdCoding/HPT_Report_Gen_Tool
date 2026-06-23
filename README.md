

# Switch sang brand oracle sql HC để sử dụng nhé



# OracleHC Report Generator



Ứng dụng nhỏ tạo báo cáo healthcheck Oracle từ dữ liệu trích xuất và mẫu Word.



## Chạy chương trình (phát triển)



- Cài phụ thuộc:



```bash

pip install -r requirements.txt

```



- Chạy GUI (máy phát triển):



```bash

python main.py

# hoặc

python gui.py

```



## Tạo file .exe (Windows)



Sử dụng PyInstaller để đóng gói ứng dụng thành một file thực thi. Ví dụ lệnh đã dùng:



```bash

pyinstaller --clean --noconfirm --onefile --windowed --name "OracleHC Report Generator" --add-data "assets\\tachnen_hpt.png;assets" --add-data "mapping\\report_mapping.yaml;mapping" gui.py

```



- Sau khi chạy xong, file `.exe` sẽ nằm trong thư mục `dist/` (ví dụ: `dist/OracleHC Report Generator.exe`).

- Các file tạm và gói khác nằm trong `build/`.



## Lấy file .exe / phân phối



- Không đẩy `dist/` hay `build/` lên GitHub (đã thêm vào `.gitignore`). Thay vào đó bạn có thể:

  - Tải file `.exe` lên trang Releases của GitHub (recommended), hoặc

  - Sử dụng Git LFS nếu muốn lưu binary lớn trong repo.



## Lưu ý



- Nếu cần build lại, xoá thư mục `dist/` và `build/` trước khi chạy PyInstaller.

- Nếu muốn chia sẻ bản build, upload file `.exe` trên GitHub Releases hoặc dịch vụ lưu trữ file.



---

