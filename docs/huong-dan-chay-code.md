# Hướng dẫn chạy chương trình KisorDoc-AI từ file runner.py

Tài liệu này hướng dẫn chi tiết cách thiết lập môi trường và chạy ứng dụng KisorDoc-AI thông qua entry point `runner.py` (khởi chạy song song Gradio UI và FastAPI API).

---

## 1. Yêu cầu hệ thống và môi trường
* **Hệ điều hành:** Windows (khuyên dùng vì tính năng tự động mở thư mục sử dụng các API Windows).
* **Phiên bản Python:** Python 3.11 trở lên.

---

## 2. Các bước khởi chạy ứng dụng

### Bước 2.1: Chuẩn bị môi trường ảo (Virtual Environment)
Khuyên dùng môi trường ảo để cô lập các thư viện của dự án, tránh xung đột hệ thống:
1. Mở cửa sổ dòng lệnh (Terminal / PowerShell / Command Prompt) tại thư mục dự án `D:\Antigravity\KisorDoc`.
2. Tạo môi trường ảo (nếu chưa có):
   ```bash
   python -m venv .venv
   ```
3. Kích hoạt môi trường ảo:
   - **Trên PowerShell:**
     ```powershell
     .venv\Scripts\Activate.ps1
     ```
   - **Trên Command Prompt (cmd):**
     ```cmd
     .venv\Scripts\activate.bat
     ```

### Bước 2.2: Cài đặt các thư viện phụ thuộc (Dependencies)
Cài đặt tất cả các thư viện cần thiết bằng cách chạy lệnh sau khi đã kích hoạt môi trường ảo:
```bash
pip install -r requirements.txt
```

### Bước 2.3: Thiết lập cấu hình (Tùy chọn)
Tạo file `.env` ngay tại thư mục gốc của dự án (`D:\Antigravity\KisorDoc`) — copy từ mẫu `.env-example`. Chương trình tự động load file này khi khởi động. Định dạng mẫu:
   ```env
   PROJECT_PATH=D:\Antigravity\1. Thanh toan nho
   ONLINE_MODE=Disable
   DATA_FOLDER=1. Data
   TEMPLATE_FOLDER=2. Templates
   FILE_FOLDER=3. Files
   DATA_SHEET=GoiThau
   CLOSE_WORD=false
   TASK_MANAGER_PROCESS=WINWORD.exe
   EXCEPTION_SHEET=S.
   APP_NAME=KisorDoc-AI
   EXCEL_TO_WORD_WIDTH_FACTOR=90
   FILE_RETRY_DELAY=2.0
   FILE_MAX_RETRIES=3
   DANH_MUC_FILE=DanhMuc
   DEFAULT_SHOW={TT}
   DEFAULT_KEY_ID=ID
   ```

### Bước 2.4: Khởi chạy ứng dụng
Chạy entry point chính bằng lệnh:
```bash
python runner.py
```

Khi khởi chạy thành công, giao diện dòng lệnh sẽ xuất hiện thông báo:
```text
========================================
         KisorDoc dang khoi dong
----------------------------------------
  Gradio UI : http://127.0.0.1:7864
  FastAPI   : http://127.0.0.1:8000
  API Docs  : http://127.0.0.1:8000/docs
========================================
```

Chương trình khởi chạy đồng thời:
* **Gradio UI** tại `http://127.0.0.1:7864` — giao diện đồ họa thao tác trực quan (Chọn Gói thầu -> Chọn template -> Chạy & Xem log).
* **FastAPI API** tại `http://127.0.0.1:8000` — tài liệu Swagger tự động tại `/docs`, hỗ trợ các endpoint `/generate`, `/templates`, `/packages`, `/jobs/*`.

Có thể chạy độc lập từng thành phần:
```bash
python app.py                    # chỉ Gradio UI
uvicorn api:app --host 0.0.0.0 --port 8000   # chỉ FastAPI API
```

---

## 3. Khắc phục một số lỗi thường gặp

### Lỗi 1: Cổng mạng bị chiếm dụng (Nhảy cổng sang 7865 hoặc báo lỗi cổng)
* **Nguyên nhân:** Cổng mạng `7864` đang bị chiếm dụng bởi một tiến trình chạy nền khác (thường là KisorDoc-AI cũ chưa được tắt hoàn toàn).
* **Khắc phục:**
  1. **Đóng nhanh tất cả các tiến trình Python đang chạy ngầm:**
     Mở CMD hoặc PowerShell và chạy lệnh:
     ```powershell
     taskkill /f /im python.exe
     ```
  2. **Tìm và tắt chính xác tiến trình đang chiếm cổng 7864:**
     * Tìm ID tiến trình (PID) đang lắng nghe cổng 7864:
       ```powershell
       netstat -ano | findstr 7864
       ```
       *(Dòng kết quả trả về sẽ có ID số ở cuối, ví dụ: `LISTENING   12448`)*
     * Tiêu diệt tiến trình đó (thay `<PID>` bằng số vừa tìm được):
       ```powershell
       taskkill /f /pid <PID>
       ```

### Lỗi 2: `ModuleNotFoundError: No module named '...'`
* **Nguyên nhân:** Bạn chưa kích hoạt môi trường ảo `.venv` hoặc chưa cài đặt đầy đủ thư viện từ `requirements.txt`.
* **Khắc phục:** Đảm bảo dòng lệnh có ký hiệu `(.venv)` ở đầu dòng trước khi gõ lệnh chạy, sau đó chạy lại lệnh `pip install -r requirements.txt`.

### Lỗi 3: Không tìm thấy gói thầu/lỗi đọc Excel
* **Nguyên nhân:** Đường dẫn `PROJECT_PATH` trong file `.env` hoặc `Config-5.txt` đang cấu hình sai, dẫn đến việc công cụ không tìm thấy thư mục chứa file Excel `1. Data/`.
* **Khắc phục:** Kiểm tra lại đường dẫn tuyệt đối của thư mục dự án trong cấu hình.

### Lỗi 4: `Python was not found; run without arguments to install...`
* **Nguyên nhân:** Python chưa được thêm vào biến môi trường PATH của hệ thống, hoặc alias mặc định của Microsoft Store đang chặn lệnh `python`.
* **Khắc phục:** Bạn có thể áp dụng 1 trong 3 cách sau:
  1. **Tắt App Execution Aliases:** Vào **Settings > Apps > Advanced app settings > App execution aliases** (hoặc tìm kiếm cụm từ này trong Windows Search) và tắt (Toggle OFF) `python.exe` và `python3.exe`.
  2. **Sử dụng bộ chạy phụ trợ (Python Launcher):** Chạy lệnh sử dụng lệnh `py` thay vì `python`:
     ```bash
     py -m venv .venv
     py runner.py
     ```
  3. **Chạy bằng đường dẫn tuyệt đối:** Sử dụng đường dẫn trực tiếp tới thư mục cài đặt Python trên máy của bạn (Ví dụ trên máy của bạn):
     - **Trên PowerShell:**
       ```powershell
       & "C:\Users\Desktop\AppData\Local\Programs\Python\Python313\python.exe" -m venv .venv
       ```
     - **Trên Command Prompt (cmd):**
       ```cmd
       "C:\Users\Desktop\AppData\Local\Programs\Python\Python313\python.exe" -m venv .venv
       ```

---

## 4. Hướng dẫn chạy công cụ tự động tạo placeholder (Migrate Text to Placeholder)

Công cụ `migrate_text_to_placeholders.py` giúp bạn quét đệ quy các file Word mẫu (`.docx`), đối chiếu với một hàng dữ liệu mẫu trong file Excel và tự động chuyển đổi các từ khóa thô thành các placeholder `{{ Key }}` tương ứng.

### Bước 4.1: Cú pháp chạy lệnh cơ bản

Kích hoạt môi trường ảo `.venv` và thực hiện gọi lệnh sau:

```bash
python migrate_text_to_placeholders.py \
  --excel "đường_dẫn_file_excel.xlsx" \
  --row <index_dòng_mẫu> \
  --docx-dir "thư_mục_chứa_word"
```
* **`--excel`**: Đường dẫn file Excel chứa dữ liệu mẫu.
* **`--row`**: Vị trí dòng dữ liệu mẫu làm mẫu để thay (1-based index, ví dụ `--row 1` là dòng dữ liệu đầu tiên ngay dưới tiêu đề cột).
* **`--docx-dir`**: Thư mục chứa các tệp tin Word cần quét và chuyển đổi.

### Bước 4.2: Các tham số nâng cao

* **`--dry-run`**: Thực hiện chạy thử nghiệm. Script sẽ mô phỏng việc thay thế và xuất báo cáo preview mà không ghi đè bất kỳ file Word thật nào.
* **`--report-dir <path>`**: Chỉ định thư mục xuất báo cáo kiểm tra (bao gồm file HTML và file Excel). Mặc định là thư mục `reports/`.
* **`--config-sheet <tên>`**: Chỉ định tên sheet cấu hình mapping của KisorDoc. Mặc định tự tìm sheet mang tên `Config`.
* **`--case-insensitive`**: So khớp không phân biệt hoa thường khi tìm kiếm văn bản trong Word.
* **`--min-length <N>`**: Bỏ qua các giá trị mẫu ngắn hơn N ký tự (mặc định: `3`).
* **`--dense-threshold <N>`**: Ngưỡng cảnh báo nếu một từ khóa xuất hiện quá nhiều lần trong 1 file (mặc định: `15`).
* **`--include "<pattern>"` / `--exclude "<pattern>"`**: Lọc các file docx qua glob pattern (VD: `--exclude "*nhap*"`).

### Bước 4.3: Ví dụ thực tế

1. **Chạy nháp kiểm tra (Khuyên dùng trước khi làm thật):**
   ```bash
   python migrate_text_to_placeholders.py \
     --excel "D:\Antigravity\1. Thanh toan nho\1. Data\data.xlsx" \
     --row 1 \
     --docx-dir "D:\Antigravity\1. Thanh toan nho\2. Templates" \
     --dry-run \
     --case-insensitive \
     --report-dir "reports/"
   ```
   *Mở báo cáo dạng HTML `reports/dryrun_YYYYMMDD_HHMMSS.html` trên trình duyệt để kiểm tra trực quan các từ khóa được highlight Đỏ/Xanh trước khi chạy thật.*

2. **Chạy chuyển đổi chính thức:**
   ```bash
   python migrate_text_to_placeholders.py \
     --excel "D:\Antigravity\1. Thanh toan nho\1. Data\data.xlsx" \
     --row 1 \
     --docx-dir "D:\Antigravity\1. Thanh toan nho\2. Templates" \
     --case-insensitive
   ```
   *Hệ thống sẽ tự động tạo file backup dạng `*.bak.docx` kèm theo mã thời gian (timestamp) trước khi thực hiện ghi đè dữ liệu.*
