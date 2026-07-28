# Hướng dẫn chạy chương trình KisorDoc-AI từ file main.py

Tài liệu này hướng dẫn chi tiết cách thiết lập môi trường và chạy ứng dụng KisorDoc-AI trực tiếp từ file nguồn `main.py`.

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
Chương trình hỗ trợ 2 nguồn cấu hình:
1. **Config-5.txt (Mặc định):** Đọc từ thư mục hệ thống `%LOCALAPPDATA%\UiPathProjectConfigs\Config-5.txt`.
2. **File .env (Ưu tiên cao hơn):** Tạo file `.env` ngay tại thư mục gốc của dự án (`D:\Antigravity\KisorDoc`) để ghi đè cấu hình. Định dạng mẫu:
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
   ```

### Bước 2.3: Khởi chạy file main.py
Chạy tập lệnh chính bằng lệnh:
```bash
python main.py
```

Khi khởi chạy thành công, giao diện dòng lệnh sẽ xuất hiện thông báo:
```text
KisorDoc-AI running at http://127.0.0.1:7864
```
Đồng thời, hệ thống sẽ tự động mở một tab mới trên trình duyệt mặc định của bạn dẫn đến địa chỉ trên để bạn thao tác trực tiếp trên giao diện đồ họa (Gradio UI).

---

## 3. Khắc phục một số lỗi thường gặp

### Lỗi 1: `OSError: Cannot find empty port in range: 7864-7864`
* **Nguyên nhân:** Cổng mạng `7864` đang bị chiếm dụng bởi một tiến trình chạy nền khác (có thể là ứng dụng Gradio cũ chưa được tắt hoàn toàn).
* **Khắc phục:** Mở file `main.py`, tìm đến dòng định nghĩa `PORT = 7864` ở cuối file và đổi sang một cổng trống khác (ví dụ: `7865`, `7866`,...) rồi chạy lại.

### Lỗi 2: `ModuleNotFoundError: No module named '...'`
* **Nguyên nhân:** Bạn chưa kích hoạt môi trường ảo `.venv` hoặc chưa cài đặt đầy đủ thư viện từ `requirements.txt`.
* **Khắc phục:** Đảm bảo dòng lệnh có ký hiệu `(.venv)` ở đầu dòng trước khi gõ lệnh chạy, sau đó chạy lại lệnh `pip install -r requirements.txt`.

### Lỗi 3: Không tìm thấy gói thầu/lỗi đọc Excel
* **Nguyên nhân:** Đường dẫn `PROJECT_PATH` trong file `.env` hoặc `Config-5.txt` đang cấu hình sai, dẫn đến việc công cụ không tìm thấy thư mục chứa file Excel `1. Data/`.
* **Khắc phục:** Kiểm tra lại đường dẫn tuyệt đối của thư mục dự án trong cấu hình.
