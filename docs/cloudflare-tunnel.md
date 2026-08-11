## KisorDoc — Hướng dẫn cài đặt & Cloudflare Tunnel

### Tổng quan kiến trúc

```
Internet → Cloudflare Tunnel → localhost:7864 (Gradio/KisorDoc)
```

KisorDoc chạy Gradio tại `http://127.0.0.1:7864` (và FastAPI tại `http://127.0.0.1:8000`). Cloudflare Tunnel sẽ tạo một kết nối an toàn từ máy Windows của bạn đến Cloudflare, giúp truy cập từ bên ngoài không cần mở port router hay cấu hình IP tĩnh.

---

## PHẦN 1 — Cài đặt KisorDoc

### Bước 1: Cài Python

Tải Python 3.11+ từ [python.org](https://python.org). **Quan trọng:** Tick chọn **"Add Python to PATH"** khi cài đặt.

Kiểm tra sau khi cài:
```cmd
python --version
pip --version
```

### Bước 2: Clone repo

```cmd
cd C:\
git clone https://github.com/phongsun01/KisorDoc.git
cd KisorDoc
```

Nếu chưa có Git, tải tại [git-scm.com](https://git-scm.com) hoặc tải ZIP từ GitHub rồi giải nén.

### Bước 3: Tạo virtual environment (khuyến nghị)

```cmd
python -m venv .venv
.venv\Scripts\activate
```

### Bước 4: Cài dependencies

```cmd
pip install -r requirements.txt
```

Các thư viện chính được cài đặt bao gồm: `python-docx`, `openpyxl`, `duckdb`, `gradio>=5.0`, `pydantic>=2.0`, `lxml`, `pandas`, `python-dotenv`, `docxtpl`, `fastapi`, `uvicorn`, `jinja2`.

### Bước 5: Cấu hình ứng dụng qua file `.env`

KisorDoc hiện tại sử dụng file cấu hình `.env` đặt tại thư mục gốc của dự án (`C:\KisorDoc\.env`) để thiết lập các tham số chạy.

1. Tạo file `.env` tại thư mục gốc:
```cmd
copy .env-example .env
```
2. Mở file `.env` bằng Notepad hoặc bất kỳ trình soạn thảo nào và chỉnh sửa đường dẫn `PROJECT_PATH` trỏ tới thư mục chứa dữ liệu của bạn (bắt buộc):
```env
PROJECT_PATH=C:\KisorDoc\Data
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

### Bước 6: Chuẩn bị cấu trúc thư mục dữ liệu

Tạo thư mục dữ liệu theo đường dẫn đã khai báo ở `PROJECT_PATH` bên trên:

```
C:\KisorDoc\Data\
├── 1. Data\          ← đặt file Excel (.xlsx) vào đây
├── 2. Templates\     ← đặt template Word (.docx) vào đây
└── 3. Files\         ← thư mục chứa file kết quả đầu ra (tự động tạo)
```

### Bước 7: Chuẩn bị file Excel

File Excel dữ liệu (đặt trong `1. Data\`) cần chứa các sheet cấu hình tối thiểu sau:
- **GoiThau** — danh sách gói thầu (cần có các cột định danh như `TT`, `Số hiệu gói thầu`, `Tên gói thầu`, `GoiThau_ID`, `Giá gói thầu`, ...)
- **Options** — cấu hình các tùy chọn dòng/cột
- **Workflow** — danh sách template liên kết theo quy trình
- **Config** — mapping biến và các cài đặt cấu trúc dữ liệu
- **Tables** — cấu hình trích xuất và copy bảng biểu từ Excel sang Word

### Bước 8: Chuẩn bị template Word

Template Word sử dụng cú pháp Jinja2 để ánh xạ biến từ Excel. Ví dụ:
```
Tên gói thầu: {{ TenGoiThau }}
Ngày ký: {{ NgayKy|date }}
Giá trị: {{ GiaTriHopDong|number }}
```

Đặt các file `.docx` vào các thư mục tương ứng trong `2. Templates\` (ví dụ `2. Templates\Opt1\`).

### Bước 9: Chạy thử ứng dụng

Khởi chạy ứng dụng thông qua entry point `runner.py`:

```cmd
cd C:\KisorDoc
.venv\Scripts\activate
python runner.py
```

Khi chạy thành công, KisorDoc sẽ khởi động đồng thời cả 2 dịch vụ:
- **Gradio UI** tại `http://127.0.0.1:7864` (Giao diện web chính)
- **FastAPI** tại `http://127.0.0.1:8000` (Hỗ trợ tài liệu REST API Swagger tại `/docs`)

---

## PHẦN 2 — Cấu hình Cloudflare Tunnel

### Yêu cầu

- Tên miền đã được **cấu hình trong tài khoản Cloudflare** (nameservers trỏ về Cloudflare)
- Tài khoản Cloudflare hoạt động bình thường (gói Free là đủ)

### Bước 1: Cài cloudflared trên Windows

Tải `cloudflared-windows-amd64.exe` từ trang phát hành chính thức:

```
https://github.com/cloudflare/cloudflared/releases/latest
```

Đổi tên file tải về thành `cloudflared.exe` và copy vào thư mục `C:\cloudflared\`.

Thêm thư mục vào biến môi trường PATH của hệ thống: **System Properties → Environment Variables → Path → New** → thêm `C:\cloudflared`.

Kiểm tra xem lệnh đã hoạt động chưa:
```cmd
cloudflared --version
```

### Bước 2: Đăng nhập Cloudflare từ CLI

```cmd
cloudflared tunnel login
```

Trình duyệt web sẽ tự động mở, yêu cầu bạn đăng nhập và chọn tên miền muốn sử dụng cho tunnel. File chứng thực `cert.pem` sau đó sẽ tự động được lưu vào thư mục `%USERPROFILE%\.cloudflared\`.

### Bước 3: Tạo Tunnel mới

```cmd
cloudflared tunnel create kisordoc
```

Lệnh này sẽ khởi tạo tunnel mới và sinh ra file cấu hình JSON chứa credentials tại `%USERPROFILE%\.cloudflared\<tunnel-id>.json`. **Lưu lại Tunnel ID** dạng UUID này.

### Bước 4: Tạo cấu hình cho Tunnel

Tạo file văn bản `%USERPROFILE%\.cloudflared\config.yml` với nội dung sau:

```yaml
tunnel: <TUNNEL_ID_CỦA_BẠN>
credentials-file: C:\Users\<TEN_USER>\.cloudflared\<TUNNEL_ID_CỦA_BẠN>.json

ingress:
  - hostname: kisordoc.yourdomain.com
    service: http://127.0.0.1:7864
  - service: http_status:404
```

Thay thế:
- `<TUNNEL_ID_CỦA_BẠN>` → UUID nhận được ở Bước 3.
- `<TEN_USER>` → Tên thư mục người dùng Windows của bạn.
- `kisordoc.yourdomain.com` → Tên miền/Subdomain bạn muốn truy cập ứng dụng.

### Bước 5: Cấu hình DNS Route

```cmd
cloudflared tunnel route dns kisordoc kisordoc.yourdomain.com
```

Lệnh này sẽ tự động thêm một bản ghi CNAME trên Cloudflare trỏ subdomain đã chọn về Tunnel của bạn.

### Bước 6: Chạy thử Tunnel thủ công

Mở hai cửa sổ dòng lệnh độc lập:

**Cửa sổ 1 — Chạy KisorDoc:**
```cmd
cd C:\KisorDoc
.venv\Scripts\activate
python runner.py
```

**Cửa sổ 2 — Chạy Tunnel:**
```cmd
cloudflared tunnel run kisordoc
```

Giờ bạn có thể truy cập từ bất kỳ đâu qua địa chỉ `https://kisordoc.yourdomain.com`.

### Bước 7: Cài đặt Tunnel thành Windows Service (Tự chạy ngầm)

Để Cloudflare Tunnel tự động khởi chạy cùng Windows mà không cần mở CMD:

```cmd
cloudflared service install
```

Khởi động service:
```cmd
sc start cloudflared
```

Kiểm tra trạng thái hoạt động:
```cmd
sc query cloudflared
```

### Bước 8: Khởi chạy KisorDoc cùng hệ thống (Tự chọn)

Để KisorDoc tự khởi động khi máy Windows bật lên, tạo file `C:\KisorDoc\start.bat` với nội dung:

```bat
@echo off
cd /d C:\KisorDoc
call .venv\Scripts\activate
python runner.py
```

Sử dụng công cụ **NSSM** (Non-Sucking Service Manager) để cài đặt file bat này thành một Windows Service:

```cmd
# Tải nssm từ trang chủ https://nssm.cc/download
nssm install KisorDoc C:\KisorDoc\start.bat
nssm set KisorDoc AppDirectory C:\KisorDoc
nssm start KisorDoc
```

---

## Bảo mật ứng dụng khi expose ra Internet

Mặc định giao diện Gradio không kích hoạt chế độ xác thực người dùng. Khi expose ra Internet công cộng, bạn rất nên sử dụng tính năng **Cloudflare Access**:

Tru cập **Cloudflare Dashboard → Zero Trust → Access → Applications → Add an application**, chọn **Self-hosted**, điền địa chỉ tên miền `kisordoc.yourdomain.com` và cấu hình chính sách bảo mật (Policy) để giới hạn quyền truy cập (ví dụ: chỉ cho phép một danh sách email cụ thể nhận mã OTP đăng nhập).

---

## Tóm tắt luồng hoạt động

```
[Khởi động Windows]
       │
       ├─► Dịch vụ cloudflared ──► Kết nối an toàn đến Cloudflare
       └─► Dịch vụ KisorDoc (NSSM) ──► Chạy ứng dụng tại localhost:7864 & :8000
       
[Người dùng truy cập]
   kisordoc.yourdomain.com ──► Cloudflare (Access / OTP) ──► Cloudflare Tunnel ──► localhost:7864
```
