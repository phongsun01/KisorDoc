# Changelog

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
