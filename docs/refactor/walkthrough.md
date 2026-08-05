Báo cáo kết quả Refactor
Chúng tôi đã hoàn thành việc tái cấu trúc (refactoring) file 
app.py
 theo đúng lộ trình đã thống nhất tại 
refactor-plan-v2.md
 và kế hoạch triển khai.

Các thay đổi đã hoàn thành
Chi tiết các file được chỉnh sửa hoặc tạo mới:

1. Hàm thuần túy (Pure Utilities)
[NEW] 
utils.py
: Tách tất cả các hàm tính toán thuần túy (_str, clean_config_key, safe_format, _parse_price, _parse_row_range, _parse_repeat_key_id, _parse_repeat_sheet_config, parse_join_expression, resolve_sheet_query, _safe_eval_condition) ra khỏi app.py. Chỉ sử dụng các thư viện chuẩn của Python (math, re, ast).
2. Lớp dịch vụ nghiệp vụ (Business Logic Service)
[NEW] 
service.py
: Định nghĩa lớp KisorService để đóng gói toàn bộ logic truy vấn cơ sở dữ liệu và cấu hình quy trình. Loại bỏ trạng thái toàn cục (config và ds) thông qua cơ chế tiêm phụ thuộc (dependency injection) vào constructor.
Sửa lỗi Preview: Sửa lỗi độ phân giải ID gói thầu dạng composite trong KisorService.run_preview thông qua việc dùng hàm _parse_repeat_key_id để lấy đúng khóa join (tương tự như logic trong hàm chạy batch).
3. Logic chạy batch (Batch Execution)
[NEW] 
batch.py
: Tách các hàm chạy sinh văn bản hàng loạt (run_batch và run_retry_batch) cùng log ghi nhận lỗi. Module này độc lập hoàn toàn với thư viện Gradio thông qua việc thay thế gr.Progress bằng hàm callback truyền vào.
4. Loại bỏ trùng lặp trong Engine
[MODIFY] 
engine.py
: Xóa các hàm clone trùng lặp (_clean_config_key, _get_option_config_from_ds) và chuyển sang import trực tiếp từ các module tương ứng của kisorlib.utils và kisorlib.service.
5. Rút gọn giao diện UI
[MODIFY] 
app.py
: Thu gọn file từ ~1780 dòng xuống còn ~360 dòng, chỉ còn giữ lại logic khởi tạo, layout giao diện Gradio UI và các event wiring gọi trực tiếp qua KisorService và module batch.
6. Bổ sung các Unit Test
[NEW] 
test_utils.py
: Viết các test case kiểm tra edge case của các hàm pure-function (_parse_price, clean_config_key, _parse_repeat_key_id, _parse_row_range, _safe_eval_condition).
[NEW] 
test_filters.py
: Viết các test case cho bộ lọc định dạng số (filter_number).
Kiểm thử & Xác minh
Kết quả kiểm thử tự động
Chúng tôi đã chạy toàn bộ các test case bằng pytest:

bash

tests\test_engine.py .                                                   [ 12%]
tests\test_filters.py .                                                  [ 25%]
tests\test_utils.py ......                                               [100%]
============================== 8 passed in 0.95s ==============================
Kiểm tra thủ công
Đảm bảo việc import app hoạt động trơn tru không lỗi cú pháp hoặc thiếu thư viện.
Xác nhận hàm create_ui() khởi tạo thành công tất cả các block Gradio và kết nối event handler chính xác.