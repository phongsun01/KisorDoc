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
