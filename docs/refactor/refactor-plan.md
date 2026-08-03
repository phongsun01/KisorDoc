Rõ rồi. Để lên plan đầy đủ, tôi fetch lại repo một lần nữa để map đúng file nào đi đâu:OK, đã có đủ thông tin. Để lên plan:---

## Quyết định kiến trúc đã thống nhất

| Quyết định | Giải pháp |
|---|---|
| **Input schema** | Sử dụng Pydantic `GenerateRequest` (bao gồm `config_row_range`) |
| **Progress tracking** | Sử dụng Callback Pattern `on_progress: Callable[[dict], None]` thay vì Generator |
| **Entry Gradio** | Đổi tên `main.py` thành `app.py` |
| **Start command** | Sử dụng `python runner.py` để chạy song song cả FastAPI và Gradio |
| **Output schema** | Sử dụng Pydantic `GenerateResult` để kiểm soát chặt chẽ kiểu trả về |

---

## Chi tiết từng bước

### Bước 1 — Tách core library và đóng gói `kisordoc/`

**Mục tiêu:** Tạo package `kisordoc/` chứa toàn bộ logic xử lý nghiệp vụ độc lập. `engine.py` là điểm vào duy nhất cho mọi caller.

**Cấu trúc thư mục mong muốn:**
```
KisorDoc/
├── kisordoc/                  # Package core library
│   ├── __init__.py
│   ├── config.py              # Đọc config, hỗ trợ biến môi trường
│   ├── dataset.py             # Kết nối DuckDB, load Excel
│   ├── engine.py              # Hàm generate_documents chính và các Pydantic models
│   ├── filters.py             # Các custom filters (date, date_long, number, upper)
│   ├── merger.py              # Logic mail merge và thay thế text placeholders
│   ├── table_copier.py        # Logic copy bảng biểu từ Excel sang Word
│   └── file_utils.py          # Dọn dẹp thư mục, rename file, xử lý file locked
├── app.py                     # Gradio Web UI (đổi tên từ main.py)
├── api.py                     # FastAPI backend web service
├── runner.py                  # Script runner chạy song song app.py và api.py
├── requirements.txt           # Danh sách thư viện phụ thuộc
└── tests/                     # Thư mục pytest kiểm thử tự động
```

#### Định nghĩa Pydantic Models trong `kisordoc/engine.py`:

```python
from pydantic import BaseModel
from typing import List, Optional, Dict, Any

class GenerateRequest(BaseModel):
    excel_path: Optional[str] = None
    option: str                                 # "Opt1"
    package_id: str                             # "MS26-01"  
    templates: List[str]                        # ["BaoCao", "TuTrinh"]
    output_dir: Optional[str] = None
    dry_run: bool = False
    config_row_range: Optional[str] = None      # Ví dụ: "2-97"

class FileResult(BaseModel):
    template_name: str
    status: str                                 # "SUCCESS", "WARNING", "ERROR", "LOCKED"
    output_path: Optional[str] = None
    message: str

class GenerateResult(BaseModel):
    status: str                                 # "SUCCESS", "PARTIAL", "FAILED"
    elapsed_seconds: float
    files: List[FileResult]
    summary: str
```

#### Signature của API Core Engine:
```python
from typing import Callable, Optional

def generate_documents(
    request: GenerateRequest,
    on_progress: Optional[Callable[[Dict[str, Any]], None]] = None
) -> GenerateResult:
    """Hàm xử lý mail merge và copy table chính của KisorDoc."""
    ...
```

---

### Bước 2 — Viết FastAPI (`api.py`)

Tạo FastAPI app chạy song song, cung cấp các endpoints:
- `POST /generate`: Nhận `GenerateRequest`, thực hiện render thông qua engine và trả về `GenerateResult`.
- `GET /templates`: Trả về danh sách các template khả dụng dựa trên cấu hình.
- `GET /packages`: Trả về danh sách các gói thầu đang có trong database.

---

### Bước 3 — Tạo runner song song (`runner.py`)

Runner sẽ khởi chạy song song cả Gradio app (`app.py`) và FastAPI (`api.py` sử dụng uvicorn) trên các cổng khác nhau bằng thread hoặc subprocess, giúp người dùng khởi chạy toàn bộ hệ thống bằng 1 lệnh duy nhất:
```bash
python runner.py
```

---

### Lộ trình thực hiện

| Thứ tự | Việc làm | Trạng thái |
|---|---|---|
| 1 | Tạo thư mục `kisordoc/`, di chuyển logic và định nghĩa Pydantic models vào `engine.py` | ⬜ Chưa bắt đầu |
| 2 | Cập nhật Gradio sang `app.py` để import và gọi thông qua `engine.py` dùng Callback | ⬜ Chưa bắt đầu |
| 3 | Viết FastAPI `api.py` sử dụng `engine.py` | ⬜ Chưa bắt đầu |
| 4 | Tạo `runner.py` chạy song song cả 2 dịch vụ | ⬜ Chưa bắt đầu |
| 5 | Viết pytest cơ bản trong thư mục `tests/` | ⬜ Chưa bắt đầu |