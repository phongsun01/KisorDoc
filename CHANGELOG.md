# Changelog - KisorDoc-AI

## v1.1.0 (2026-07-28)

- **Di chuyển cú pháp Placeholder:** Chuyển đổi toàn bộ template Word sang định dạng ngoặc nhọn kép `{{ }}` tương thích với Jinja2 và thư viện `docxtpl`.
- **Phân tách các biến trùng lặp (Date):** Tách riêng các biến định dạng ngày (ví dụ `<<KHLCNT_TTr.Date>>` -> `{{KHLCNT_TTr_Date|date}}`) và các biến văn bản (ví dụ `<<KHLCNT_TTr>>` -> `{{KHLCNT_TTr}}`) để tránh xung đột ghi đè dữ liệu.
- **Chuẩn hóa biến Bảng:** Chuyển đổi toàn bộ `{DanhMuc}` và `{DanhMucKoGia}` trong Word sang `{{DanhMuc}}` và `{{DanhMucKoGia}}`.
- **Sửa lỗi biến chứa dấu chấm:** Tự động tổ chức lại từ điển dữ liệu phẳng (flat dictionary) từ Excel thành cấu trúc lồng nhau (nested dictionary) để Jinja2 hiển thị đúng các biến như `{{KHLCNT_TTr.Dvi}}` và `{{DuToan.NguoiLap}}` mà không bị báo lỗi `UndefinedError`.
- **Cải tiến cấu hình (.env):** Cho phép file `.env` ghi đè toàn bộ các trường cấu hình trong file `Config-5.txt`.
- **Cải tiến tính năng Mở thư mục Output:** Giải quyết tuyệt đối đường dẫn tuyệt vời trên hệ điều hành Windows, bổ sung kiểm tra tồn tại và bắt lỗi ngoại lệ khi mở thư mục.

## v1.0.0 (2026-07-27)

- Phiên bản đầu tiên chuyển đổi từ bot UiPath sang mã nguồn Python.
