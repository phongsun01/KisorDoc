Rõ rồi. Tóm lại 3 lý do crash lần trước đã được fix hết trong `engine.py` mới:

- ✅ **Lý do A** (`copy_tables_for_file` signature) — đã fix #5
- ✅ **Lý do B** (`query_row`, `query_option` không tồn tại) — đã fix #3
- ✅ **Lý do C** (Repeat mode) — đây là vấn đề còn lại, `engine.py` hiện **chưa hỗ trợ Repeat mode**

Trước khi viết, tôi cần đọc lại đúng phần Repeat mode trong `app.py` hiện tại:Đọc xong. Tôi thấy vấn đề cốt lõi rõ ràng hơn bây giờ.

---

## Đánh giá thực tế

Sau khi đọc kỹ `run_batch`, tôi thấy **bug #2 thực ra không nên sửa lúc này**. Đây là lý do:

### `run_batch` và `engine.py` có logic khác nhau về cơ bản

| | `run_batch` (app.py) | `engine.py` |
|---|---|---|
| Copy file | `copy_templates_to_output()` — copy trước, merge sau | `shutil.copy2` + merge trong cùng loop |
| File output | Dùng tên file gốc của template (`.docx`), `rename_output` đổi sau | Giống |
| Retry | Copy chỉ các file lỗi, **không xóa folder** | Không có |
| Repeat mode | Register `TCGTTD` động, rebuild context per-member | Không có |
| Missing placeholder | Quét bằng `DocxTemplate.get_undeclared_template_variables()` | Không có |
| Write retry | `write_with_retry()` 3 lần khi file bị lock | Không có |
| Logger | `IncrementalRunLogger` ghi file `.log` | Không có |
| `used_names` | Track để tránh trùng tên output | Không có |

`engine.py` hiện tại thiếu **7 tính năng** so với `run_batch`. Nếu thay `run_batch` bằng `generate_documents()` sẽ mất hết 7 cái này.

### Kết luận thực tế

Bug #2 không phải "engine là dead code — xấu về kiến trúc". Nó phản ánh thực tế rằng **`engine.py` chưa đủ trưởng thành để thay thế `run_batch`**. Ép thay thế bây giờ sẽ mất tính năng quan trọng (retry, repeat mode, missing placeholder warning, logging).

**Đề xuất:** Để `engine.py` phục vụ `api.py` trước — đó là use case phù hợp với khả năng hiện tại của nó. Song song đó, từng bước bổ sung các tính năng còn thiếu vào `engine.py`, đến khi đủ thì mới chuyển `run_batch`.

### Lộ trình bổ sung nếu muốn làm đúng

| Bước | Tính năng cần thêm vào engine.py | Độ phức tạp |
|---|---|---|
| 1 | `write_with_retry` (file lock) | Thấp |
| 2 | Missing placeholder scan | Thấp |
| 3 | `IncrementalRunLogger` | Thấp |
| 4 | `used_names` tracking | Thấp |
| 5 | Retry mode (không xóa folder) | Trung bình |
| 6 | Repeat mode (TCGTTD dynamic) | Cao — phụ thuộc DuckDB session |

---

## Cập nhật kết quả xử lý (v4.0.0 & v4.0.1)

Các vấn đề kiến trúc và lỗi trong danh sách trên đã được giải quyết triệt để trong đợt tái cấu trúc phiên bản 4.0.0 & 4.0.1:

1. **Tái cấu trúc mã nguồn app.py**:
   - `run_batch`, `run_retry_batch`, `IncrementalRunLogger`, `write_with_retry` đã được di chuyển toàn bộ từ `app.py` sang `kisorlib/batch.py` và `kisorlib/service.py`.
   - `app.py` được tinh giản từ ~1780 dòng xuống còn ~360 dòng, loại bỏ được global state và chỉ còn làm nhiệm vụ kết nối UI Gradio.

2. **Khắc phục các lỗi vận hành & Hardcode**:
   - **Sửa lỗi Preview ở chế độ Repeat**: Sửa lỗi composite key_id trong `run_preview`.
   - **Khắc phục hardcode trong api.py**: Chuyển đổi các cấu hình đường dẫn `"1. Data"` và truy vấn bảng `"GoiThau"` sang sử dụng cấu hình động `cfg.data_path` và `cfg.DataSheet` từ `AppConfig`. Sửa lỗi truyền tham số `excel_files` lỗi thời vào `DataSet`.
   - **Bảo mật SQL (Parameter Binding)**: Nâng cấp `DataSet.query()` hỗ trợ parameter binding `?`, giải quyết triệt để vấn đề SQL Injection và lỗi cú pháp khi dữ liệu chứa ký tự nháy đơn (`'`).

3. **Bổ sung Unit Tests tự động**:
   - Viết thêm bộ test chuyên biệt [tests/test_service.py](file:///D:/Antigravity/KisorDoc/tests/test_service.py) chạy trên bộ dữ liệu giả lập in-memory để kiểm thử `KisorService`, Repeat mode, đăng ký thành viên tạm thời và preview composite key.

### Bảo mật SQL (SQL Identifier Injection - Đã giải quyết ở v4.0.3)
- **Vấn đề**: Trước đây, mặc dù giá trị dữ liệu đã được tham số hóa an toàn bằng `?`, tên cột (`join_key`, `key_id`) và tên bảng (`right_sheet`) vẫn được nối chuỗi trực tiếp từ cấu hình Excel.
- **Giải pháp**: Đã bổ sung hàm `validate_sql_identifier` sử dụng regex whitelist `^[A-Za-z0-9_\s\-\.\#\u00C0-\u1EF9]+$` để kiểm tra định dạng an toàn trước khi thực hiện câu lệnh SQL, triệt tiêu hoàn toàn nguy cơ chèn mã SQL qua Excel.

### Giải quyết Dual-Pipeline (Đã giải quyết ở v4.1.0)
- **Vấn đề**: Hai nhánh chạy song song (UI dùng `batch.py` và API dùng `engine.py`) có logic sinh tài liệu trùng lặp, dễ dẫn đến lệch tính năng và lỗi hồi quy.
- **Giải pháp**: Tách toàn bộ core sinh file đồng bộ (sync core) sang `kisorlib/generator.py` làm Single Source of Truth. UI (`batch.py`) và API (`engine.py`) đóng vai trò các adapter mỏng gọi xuống core. Toàn bộ 7 tính năng (retry, repeat, log, placeholders,...) nay đã được chia sẻ chung. Đồng thời tách các hàm SQL sang `kisorlib/sql_join.py`.
