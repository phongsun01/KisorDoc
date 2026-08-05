Kế hoạch triển khai Refactor app.py sang các Module kisorlib
Tái cấu trúc file 
app.py
 dài 1780 dòng thành các file nhỏ hơn, áp dụng nguyên lý đơn trách nhiệm (Single Responsibility) bên trong gói kisorlib, dọn dẹp mã nguồn trùng lặp trong 
engine.py
, và bổ sung unit test.

Các thay đổi đề xuất
Chúng tôi sẽ tiến hành refactor theo 6 bước có cấu trúc như sau:

Bước 1: Tạo các hàm tiện ích thuần túy
[NEW] 
utils.py
Tạo kisorlib/utils.py chứa các hàm pure function được tách ra từ app.py:

_str
clean_config_key
safe_format
_parse_price
_parse_row_range
_parse_repeat_key_id
_parse_repeat_sheet_config
parse_join_expression (và _JOIN_RE)
resolve_sheet_query
_safe_eval_condition (hàm phụ trợ để kiểm tra điều kiện lọc)
Bước 2: Dọn dẹp mã nguồn trùng lặp trong engine.py (Bước 0b)
[MODIFY] 
engine.py
Loại bỏ các hàm phụ trợ bị trùng lặp trong engine.py bằng cách import các phiên bản sạch từ kisorlib/utils.py và kisorlib/service.py:

Xóa _clean_config_key (dùng utils.clean_config_key)
Xóa _get_option_config_from_ds
Loại bỏ các khối logic trùng lặp.
Bước 3: Bổ sung các Unit Test
[NEW] 
test_utils.py
Viết các test case bằng pytest cho các hàm thuần túy trong utils.py:

_parse_price (xử lý None, NaN, "1.500.000", "")
clean_config_key (xử lý các hậu tố như .Date.Long, .upper, |)
parse_join_expression (xử lý các toán tử <*>, <*, *>, *, SELECT passthrough)
_parse_repeat_key_id (xử lý khi có hoặc không có dấu |, khoảng trắng thừa)
[NEW] 
test_filters.py
Viết test case cho bộ lọc filter_number (xử lý các định dạng "1.500", "1.123", "1,500,000", 1500000.0).

Bước 4: Tách lớp dịch vụ nghiệp vụ
[NEW] 
service.py
Tạo lớp KisorService nhận config và ds qua constructor để loại bỏ trạng thái toàn cục:

get_options
get_option_config
get_config_for_option
get_all_option_templates
check_condition
get_packages
get_package_details
get_package_excel_file
get_workflow_templates
get_repeat_members
register_temporary_tcgttd
run_preview (Sửa lỗi run_preview composite key_id để giải quyết chính xác ID gói thầu).
Bước 5: Tách logic chạy batch
[NEW] 
batch.py
Tạo module chạy sinh tài liệu hàng loạt bằng cách di chuyển các hàm:

write_with_retry
IncrementalRunLogger
run_batch (nhận service và dùng progress_cb: Callable | None thay vì gr.Progress)
run_retry_batch (tách thành hàm top-level)
Bước 6: Đơn giản hóa app.py
[MODIFY] 
app.py
Xóa bỏ toàn bộ mã nguồn đã được tách ra.
Dọn dẹp các thư viện import không còn sử dụng.
Khởi tạo service và import batch.
Kết nối các event handler của Gradio UI để gọi qua các phương thức của service và batch.
Giữ lại logic khởi động init() và khởi tạo UI.
Kế hoạch kiểm thử & Xác minh
Kiểm thử tự động
Chạy lệnh pytest tests/ để đảm bảo tất cả các test case (cũ và mới) đều vượt qua.
Kiểm tra thủ công
Khởi động ứng dụng thông qua lệnh python runner.py.
Thực hiện chạy thử UI Preview, chạy thực tế (full run), chạy lại (retry) trên nhiều quy trình khác nhau (đặc biệt là chế độ lặp Repeat) để đảm bảo tính ổn định và chính xác.