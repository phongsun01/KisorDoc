# Hướng dẫn sử dụng chức năng Migrate Template (KisorDoc)

Tài liệu này hướng dẫn chi tiết cách sử dụng chức năng **Migrate Template** để tự động chuyển đổi cú pháp placeholder cũ (như `<<TenBien>>`, `<<TenBien.Date>>`, `{DanhMuc}`) sang cú pháp tiêu chuẩn của **Jinja2** (như `{{TenBien}}`, `{{TenBien_Date|date}}`, `{{DanhMuc}}`).

---

## 1. Tổng quan & Quy tắc chuyển đổi

Chức năng Migrate Template giúp đồng bộ hóa các file template `.docx` cũ lên chuẩn cú pháp Jinja2 mới được hỗ trợ bởi KisorDoc v5.0.0+. 

Các quy tắc chuyển đổi được áp dụng tự động bao gồm:

| Cú pháp cũ | Cú pháp Jinja2 mới | Mô tả |
|---|---|---|
| `<<TenBien>>` | `{{TenBien}}` | Placeholder cơ bản |
| `<<TenBien.Date>>` | `{{TenBien_Date\|date}}` | Thêm hậu tố `_Date` và bộ lọc `\|date` |
| `<<TenBien.Date.Long>>` | `{{TenBien_Date\|date_long}}` | Thêm hậu tố `_Date` và bộ lọc `\|date_long` |
| `<<TenBien.Day/Month/Year>>` | `{{TenBien_Date\|day/month/year}}` | Tương ứng các bộ lọc thời gian tách lẻ |
| `<<TenBien.Number>>` | `{{TenBien\|number}}` | Định dạng số |
| `<<TenBien.Chu>>` | `{{TenBien\|num2text}}` | Chuyển số thành chữ |
| `<<TenBien.Upper>>` | `{{TenBien\|upper}}` | Viết hoa |
| `{DanhMuc}` | `{{DanhMuc}}` | Dấu ngoặc đơn trong bảng biểu tự động chuyển sang ngoặc kép |

---

## 2. Cách chạy qua giao diện đồ họa (Gradio UI)

Giao diện đồ họa là phương pháp trực quan và dễ sử dụng nhất cho người dùng thông thường.

### Các bước thực hiện:

1. **Khởi động ứng dụng**:
   Kích hoạt môi trường ảo và chạy file `runner.py`:
   ```bash
   .venv\Scripts\activate
   python runner.py
   ```
2. **Truy cập Giao diện**:
   Mở trình duyệt và truy cập địa chỉ `http://127.0.0.1:7864`.
3. **Mở Tab Migrate**:
   Chọn tab **"3. Migrate Template"** trên thanh điều hướng.
4. **Điền thông số**:
   - **Thư mục template**: Nhập đường dẫn thư mục chứa các file `.docx` cần migrate (ví dụ: `D:\Antigravity\1. Thanh toan nho\2. Templates`). Nếu để trống, hệ thống sẽ sử dụng thư mục mặc định từ cấu hình `.env` (`TEMPLATE_FOLDER`).
   - **Quét tất cả thư mục con**: Tích chọn nếu muốn tìm kiếm đệ quy trong các thư mục con.
   - **Tạo bản backup (.bak.docx)**: Nên tích chọn để đề phòng lỗi phát sinh (file backup sẽ được lưu trữ trong thư mục `bak/` cùng cấp với file gốc).
5. **Chạy kiểm tra (Dry-run)**:
   - Nhấn nút **"🔍 Phân tích (Dry-run)"**.
   - Hệ thống sẽ liệt kê trước các file cần thay đổi cùng chi tiết các placeholder cũ và mới trong ô **Kết quả** mà không làm thay đổi file thực tế.
6. **Chạy Migrate thật**:
   - Nhấn nút **"⚡ Migrate thật"** sau khi kiểm tra kết quả dry-run chuẩn xác.
   - Hệ thống sẽ sao lưu và cập nhật trực tiếp cấu trúc XML của các file `.docx`.

---

## 3. Cách sử dụng qua API Python (Lập trình)

Nếu bạn muốn tích hợp chức năng này vào các pipeline tự động hóa hoặc script riêng, bạn có thể gọi thư viện `kisorlib.migrator`.

```python
from pathlib import Path
from kisorlib.migrator import migrate_folder, format_summary

# 1. Định nghĩa thư mục template cần xử lý
template_dir = Path("D:/Antigravity/KisorDoc/docs/Data")

# 2. Thực hiện Dry-run để quét thay đổi
results = migrate_folder(
    folder=template_dir,
    dry_run=True,
    backup=True,
    recursive=True
)

# 3. Định dạng kết quả và in ra console
summary = format_summary(results, dry_run=True)
print(summary)

# 4. Khi muốn áp dụng thật: dry_run=False
# results = migrate_folder(template_dir, dry_run=False, backup=True)
```

### Các hàm API chính trong `kisorlib/migrator.py`:
- `migrate_xml(xml: str) -> (new_xml, list[PlaceholderChange])`: Xử lý chuỗi XML thô (pure function, thích hợp cho viết test).
- `migrate_file(path: Path, dry_run: bool, backup: bool) -> FileResult`: Xử lý cho duy nhất 1 file `.docx`.
- `migrate_folder(folder: Path, ...) -> list[FileResult]`: Duyệt thư mục để migrate hàng loạt file.

---

## 4. Cơ chế an toàn và Backup

- **Tự động sao lưu**: Khi thực hiện migrate thật với tùy chọn `backup=True`, hệ thống sẽ tạo một thư mục mang tên `bak/` tại thư mục chứa file `.docx` gốc và copy file nguyên bản vào đó với định dạng `<tên_file>.bak.docx`.
- **Tránh ghi đè trùng lặp**: Các file trong thư mục `bak/` hoặc có đuôi `.bak.docx` sẽ tự động bị bỏ qua trong tất cả các lượt quét tiếp theo của bộ di trú (migrator).
- **Ghi file an toàn**: Chương trình luôn tạo ra một file nháp tạm `.tmp.docx`, thực hiện ghi dữ liệu hoàn chỉnh, sau đó mới tiến hành ghi đè thay thế file gốc. Nếu quá trình ghi bị lỗi, file `.tmp.docx` sẽ bị xóa và file gốc được giữ nguyên vẹn.
