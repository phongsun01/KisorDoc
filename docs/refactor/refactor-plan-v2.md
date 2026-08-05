## Refactor Plan — `app.py` (1780 lines)

### Mục tiêu

```
app.py (1780 lines)  →  4 files trong kisorlib/ + app.py mỏng (~350 lines)
```

---

### Sơ đồ file mới

```
kisorlib/
├── utils.py          (mới) ~150 lines  — pure functions, không dependency
├── service.py        (mới) ~450 lines  — business logic, nhận (config, ds)
├── batch.py          (mới) ~600 lines  — batch execution
└── ... (các file hiện có giữ nguyên)

app.py               ~350 lines  — chỉ còn init + Gradio UI
```

> ⚠️ Đổi tên `runner.py` → `batch.py` để tránh trùng với `runner.py` ở root (process launcher).

---

### Chi tiết từng file

#### `kisorlib/utils.py` (~150 lines)

Hàm pure, không import `config`/`ds`, dễ unit test độc lập.
Import duy nhất: `re`, `math`.

| Hàm | Line hiện tại | Ghi chú |
|-----|--------------|---------|
| `_str()` | 35 | |
| `clean_config_key()` | 44 | |
| `safe_format()` | 172 | có `import re` nội bộ → chuyển lên top-level |
| `_parse_price()` | 670 | — đã dời xuống dưới so với plan cũ |
| `_parse_row_range()` | 132 | có `import re` nội bộ → chuyển lên top-level |
| `_parse_repeat_key_id()` | 496 | — đã dời so với plan cũ |
| `_parse_repeat_sheet_config()` | 402 | |
| `_JOIN_RE` + `parse_join_expression()` | 316–363 | |
| `resolve_sheet_query()` | 369 | |
| `_safe_eval_condition()` | 234 | helper của `check_condition` — tách cùng |
| `import ast as _ast` | 232 | đi kèm `_safe_eval_condition` |

> **Lưu ý mới:** `_parse_price()` hiện ở line 670, không phải 565 như plan cũ.
> `safe_format()` và `_parse_row_range()` có `import re` nội bộ trong thân hàm
> → khi tách ra `utils.py` cần chuyển thành top-level import.

---

#### `kisorlib/service.py` (~450 lines)

Class `KisorService` nhận `config` + `ds` qua constructor — xóa global state.

```python
class KisorService:
    def __init__(self, config: AppConfig, ds: DataSet): ...
```

| Method | Line hiện tại | Ghi chú |
|--------|--------------|---------|
| `get_options()` | 99 | |
| `get_option_config()` | 104 | dùng `config.DataSheet`, `config.DefaultShow`, `config.DefaultKeyId` |
| `get_config_for_option()` | 148 | gọi `ds.query_rows` |
| `get_all_option_templates()` | 161 | |
| `check_condition()` | 184 | gọi `_safe_eval_condition` từ `utils.py` |
| `get_packages()` | 540 | — đã dời so với plan cũ |
| `get_package_details()` | 588 | |
| `get_package_excel_file()` | 386 | |
| `get_workflow_templates()` | 617 | |
| `get_repeat_members()` | 438 | |
| `register_temporary_tcgttd()` | 497 | |
| `run_preview()` | 690 | có `import openpyxl` nội bộ → chuyển top-level |

> **Lưu ý mới:**
> - `run_preview()` hiện ở line 690, không phải 585. Có `import openpyxl` inline → chuyển top-level khi tách.
> - **Fix run_preview key_id composite đồng thời khi move:**
>   ```python
>   # Trước:
>   goi_thau_id = _str(selected_pkg.get(key_id))
> 
>   # Sau:
>   left_key, _ = _parse_repeat_key_id(key_id)
>   goi_thau_id = _str(selected_pkg.get(left_key))
>   ```
> - `check_condition()` phụ thuộc `_parse_price()` và `_safe_eval_condition()` — cả hai phải export từ `utils.py` trước.
> - `get_workflow_templates()` gọi `check_condition()` → phải là method cùng class, không tách riêng.

Import: `kisorlib/utils.py`, `kisorlib/config.py`, `kisorlib/dataset.py`, `pandas`.

---

#### `kisorlib/batch.py` (~600 lines)

Không import Gradio — nhận `progress_cb: Callable | None` thay vì `gr.Progress`.

| Symbol | Line hiện tại | Ghi chú |
|--------|--------------|---------|
| `write_with_retry()` | 820 | đã normalize tuple, không cần sửa thêm |
| `IncrementalRunLogger` | 846 | đã có `record_ok/record_error/record_locked` — giữ nguyên |
| `run_batch()` | 928 | async generator — xem breaking changes |
| `run_retry_batch()` | 1695 | hiện là closure bên trong `create_ui()` — cần **tách ra** thành top-level function |

> **Lưu ý mới — `run_retry_batch` là closure:**
> Hiện tại `run_retry_batch` (line 1695) nằm bên trong `create_ui()` và gọi thẳng
> `run_batch()` (cùng scope). Sau refactor cần tách thành top-level trong `batch.py`:
>
> ```python
> # batch.py
> async def run_retry_batch(retry_state, progress_cb=None):
>     ...
>     async for r in run_batch(..., progress_cb=progress_cb):
>         yield r
> ```
>
> `app.py` sẽ wrap lại:
> ```python
> async def _ui_run_retry(retry_state, progress=gr.Progress()):
>     async for r in batch.run_retry_batch(retry_state, progress_cb=progress):
>         yield r
> ```

Import: `kisorlib/utils.py`, `kisorlib/service.py`, `kisorlib/merger.py`,
`kisorlib/table_copier.py`, `kisorlib/file_utils.py`, `shutil`, `docxtpl`.

---

#### `app.py` (~350 lines còn lại)

| Phần | Lines hiện tại | Ghi chú |
|------|---------------|---------|
| Imports | 1–30 | Thu gọn: bỏ `re`, `math`, `pd`, `DocxTemplate`, `shutil`... |
| `init()` | 78–96 | Giữ, khởi tạo `service` thay vì `config`/`ds` riêng |
| `create_ui()` | 1274–1760 | Đổi tất cả call sang `service.*` và `batch.*` |
| `__main__` | 1763–1780 | Giữ nguyên — `runner.py` ở root đã dùng `create_ui()` qua import |

> **Lưu ý mới — `__main__` trong `app.py`:**
> `runner.py` (root) gọi `from app import create_ui` — không đụng đến `__main__`.
> Giữ `__main__` trong `app.py` để vẫn chạy được `python app.py` độc lập.
> `warnings.filterwarnings` và `sys.stdout.reconfigure` (line 16–18) chuyển sang
> `runner.py` (root) vì đó là entry point thực sự, nhưng giữ lại trong `app.py`
> như fallback khi chạy trực tiếp.

Global state:
```python
# Trước:
config: AppConfig | None = None
ds: DataSet | None = None
ui_labels = {}

# Sau:
service: KisorService | None = None
ui_labels: dict = {}
```

---

### Dependency graph sau refactor

```
runner.py (root — process launcher, KHÔNG đụng đến)
├── app.py
│    ├── kisorlib/batch.py
│    │    ├── kisorlib/service.py
│    │    │    ├── kisorlib/utils.py       (no deps ngoài stdlib)
│    │    │    ├── kisorlib/dataset.py
│    │    │    └── kisorlib/config.py
│    │    ├── kisorlib/merger.py
│    │    ├── kisorlib/table_copier.py
│    │    └── kisorlib/file_utils.py
│    └── kisorlib/service.py              (trực tiếp, cho UI callbacks)
└── api.py → kisorlib/engine.py           (KHÔNG bị ảnh hưởng)
```

Không có circular import.

---

### Thứ tự thực hiện

```
Bước 1  → kisorlib/utils.py  (tạo + commit)
          - Copy các hàm pure
          - Chuyển import re/math/ast lên top-level

Bước 0b → engine.py dedup    (commit riêng)
          - engine.py: xóa _clean_config_key, _get_option_config_from_ds, _build_context
          - Import các hàm tương đương từ utils và service

tests   → tests/test_utils.py + tests/test_filters.py (commit riêng)
          - Viết unit test cho các hàm pure/filters:
            + _parse_price: None, NaN, "1.500.000", ""
            + clean_config_key: .Date.Long, .upper, | trong key
            + parse_join_expression: <*>, <*, *>, *, SELECT passthrough
            + _parse_repeat_key_id: không có |, có |, khoảng trắng thừa
            + filter_number: "1.500", "1.123", "1,500,000", 1500000.0

Bước 2  → kisorlib/service.py (fix run_preview luôn trong bước này)
          - Tạo class KisorService(config, ds)
          - Move các hàm business logic vào
          - Fix lỗi run_preview key_id composite đồng thời khi move
          - Chuyển import openpyxl (run_preview) lên top-level

Bước 3  → kisorlib/batch.py
          - Move write_with_retry, IncrementalRunLogger, run_batch
          - Tách run_retry_batch ra khỏi closure create_ui()
          - Đổi gr.Progress → progress_cb: Callable | None

Bước 4  → app.py mỏng
          - Xóa code đã tách khỏi app.py, thêm import tương ứng
          - Thay global config/ds bằng service
          - Đổi tất cả call sang service.* và batch.*
          - Wrap progress_cb cho Gradio
          - Thu gọn imports
```

Mỗi bước commit riêng — `app.py` gốc vẫn chạy cho đến bước 4.

---

### Các thay đổi breaking cần chú ý

| # | Vấn đề | Giải pháp |
|---|--------|-----------|
| 1 | Global `config`, `ds` bị xóa | `service.config`, `service.ds` |
| 2 | `run_batch` nhận `gr.Progress` | Đổi thành `progress_cb: Callable \| None` |
| 3 | Callbacks UI gọi function trực tiếp | Đổi sang `service.get_packages(...)` |
| 4 | `_sel` dict trong closure | Giữ nguyên trong `app.py`, không tách |
| 5 | `ui_labels` global | Chuyển vào `init()`, truyền vào `create_ui(service, ui_labels)` |
| 6 | `run_retry_batch` là closure | Tách thành top-level trong `batch.py` (mới) |
| 7 | `import re` nội bộ trong hàm | Chuyển top-level khi tách sang `utils.py` |
| 8 | `import openpyxl` nội bộ trong `run_preview` | Chuyển top-level khi tách sang `service.py` |
