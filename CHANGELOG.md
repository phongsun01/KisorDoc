# Changelog

## [2.0.1] - 2026-07-31

### Fixed
- Khắc phục hoàn toàn lỗi cảnh báo thiếu dữ liệu (`Warning: Placeholder ... không có data`) khi chạy các template đã được migrate bằng cách ánh xạ đồng thời cả khóa gốc và khóa hậu tố `_Date` (tương thích ngược hoàn toàn với cả template mới `{{Dexuat}}` và template migrate cũ `{{Dexuat_Date}}`).

## [2.0.0] - 2026-07-30

### Added
- **Cấu hình Option động (Dynamic Options):** Hỗ trợ khai báo các cột `Sheet`, `Show`, `KeyId` trong sheet `Options` để tự động hóa định dạng nhãn hiển thị và tên sheet nguồn dữ liệu chính.
- **Lọc điều kiện động (Dynamic Conditions):** Bổ sung cột `Condition` trong sheet `Workflow` hỗ trợ cú pháp ngoặc nhọn `{Tên cột/Tên biến}` và tự động parse chuỗi số Excel (ví dụ: `150.000.000` -> `150000000`) khi so sánh logic trong Python `eval`.
- **Phân vùng Config theo Option (Config Range):** Bổ sung cột `Config` trong sheet `Options` cho phép tách biệt các vùng ánh xạ trong sheet Config (ví dụ: `2-97`, `99-253`) tránh xung đột dữ liệu.
- **Tập lệnh di chuyển nâng cao:** Thêm script `migrate_modifiers.py` hỗ trợ nâng cấp đồng bộ toàn bộ modifier (date, day, month, year, number, chu/text) sang cú pháp bộ lọc Jinja2 `|` chuẩn xác.

### Changed
- Cải tiến hiệu năng nạp file Excel trong DuckDB dataset bằng cơ chế cache `query_rows` theo range dòng, giúp chỉ mở file 1 lần khi nhiều Option dùng chung phân vùng.
- Tự động di chuyển toàn bộ file `.bak.docx` sang thư mục con `bak/` tương ứng của từng thư mục quy trình.

## [1.9.0] - 2026-07-29

### Added
- Cấu hình hệ số chuyển đổi độ rộng cột từ Excel sang Word bằng biến môi trường EXCEL_TO_WORD_WIDTH_FACTOR (mặc định = 90) trong file .env và AppConfig.
- Tích hợp 6 đề xuất tính năng nâng cao (F1 - F6) vào tài liệu PRD chính thức [PRD-WordBatchProcessor.md](docs/PRD-WordBatchProcessor.md).
- Thêm **Phase 8** và mục **Các lỗi cần sửa sau** vào tài liệu [CHECKLIST-KisorDoc-AI.md](docs/CHECKLIST-KisorDoc-AI.md).

### Changed
- Cập nhật hàm un_batch trong main.py để định dạng log kết quả xuống dòng (join bằng \n) giúp hiển thị rõ ràng trên giao diện Gradio Textbox thay vì hiển thị dạng danh sách thô.
- Hợp nhất tài liệu PRD bằng cách gộp file PRD-Enhancements.md vào PRD-WordBatchProcessor.md và xóa file PRD-Enhancements.md cũ.

### Fixed
- Sửa lỗi lệch thẻ XML (Opening and ending tag mismatch: body line 2 and tc) gây hỏng file Word bằng cách cập nhật regex di chuyển template an toàn hơn, không quét xuyên qua các tag cấu trúc XML quan trọng (p, 	c, 	r, 	bl).
- Khôi phục và di chuyển thành công toàn bộ 36 file template gốc sang cấu trúc template {{}} mới an toàn.
- Sửa lỗi tìm kiếm sheet Excel khi tên sheet chứa khoảng trắng ngoài mong muốn (ví dụ: ' S.DoDa' thay vì 'S.DoDa') bằng cách so khớp strip whitespace.
- Sửa lỗi crash khi định dạng ngày do dữ liệu chứa giá trị NaT (Not a Time) trong Pandas.
- Cải thiện nút mở thư mục output sử dụng subprocess.Popen đáng tin cậy hơn trên Windows từ Gradio background thread.
