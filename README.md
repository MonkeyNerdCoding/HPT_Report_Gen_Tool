<p align="center">
  <img src="assets/hpt-logo.jpg" alt="HPT" width="420">
</p>

# HPT Report Generator Tool

Tool web nội bộ hỗ trợ tạo báo cáo Health Check/Tuning/Security từ dữ liệu OracleHC, SQLHealthcheck và template Word của HPT. Người dùng upload template `.docx` và gói dữ liệu `.zip`, hệ thống tự đọc dữ liệu, map placeholder, render bảng/biểu đồ và trả về file report `.docx`.

## Tính năng chính

- Giao diện web chạy local bằng FastAPI.
- Generate report OracleHC và SQLHealthcheck.
- Insert placeholder vào template Word theo mapping có sẵn.
- Quản lý lịch sử report, log xử lý và tải lại file đã generate.
- Scan/kiểm tra placeholder trong template.
- Hỗ trợ AI review nếu cấu hình API key phù hợp.

## Clone repo

```bash
git clone https://github.com/MonkeyNerdCoding/HPT_Report_Gen_Tool/tree/feature/hpt-report-generator-updates
cd HPT_Report_Gen_Tool
```

## Cài đặt

Yêu cầu Python 3.11+.

```bash
python -m venv .venv
```

Windows PowerShell:

```powershell
.\.venv\Scripts\Activate.ps1
```

Windows CMD:

```cmd
.venv\Scripts\activate.bat
```

Cài thư viện:

```bash
pip install -r requirements.txt
```

## Chạy web app

```bash
uvicorn web.app:app --reload
```

Mở trình duyệt:

```text
http://127.0.0.1:8000
```

## Cách dùng nhanh

1. Chọn mode `OracleHC` hoặc `SQLHealthcheck`.
2. Upload Word template `.docx`.
3. Upload source package `.zip`.
4. Nếu template chưa có placeholder, dùng chức năng insert placeholder trước.
5. Bấm generate và tải file report `.docx` sau khi xử lý xong.

## Cấu trúc repo

- `web/`: giao diện web, API FastAPI, static assets và templates.
- `mapping/`: file mapping placeholder với dữ liệu OracleHC/SQLHealthcheck.
- `extraction/`: parser đọc bảng và biểu đồ từ source HTML.
- `rendering/`: logic render bảng/biểu đồ vào Word.
- `sql_healthcheck/`: xử lý dữ liệu SQLHealthcheck.
- `placeholder_inserter.py`: insert placeholder vào template Word.
- `app_logic.py`: orchestration chính cho generate report.
- `data/`, `runtime_jobs/`: dữ liệu runtime local, không commit lên GitHub.

## Ghi chú

- Không upload dữ liệu khách hàng, file report sinh ra, `.env`, log runtime hoặc cache lên GitHub.
- Nếu dùng AI review, tạo `.env` từ `.env.example` và cấu hình API key cần thiết.
