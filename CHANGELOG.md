# Changelog

## [4.0.0] - 2026-08-05

### Changed
- **Tái cấu trúc mã nguồn app.py toàn diện:** Phân rã file `app.py` khổng lồ dài ~1780 dòng thành các module nhỏ, đơn nhiệm bên trong thư mục `kisorlib/` để nâng cao chất lượng code và dễ bảo trì:
  - [kisorlib/utils.py](file:///D:/Antigravity/KisorDoc/kisorlib/utils.py): Các hàm tiện ích thuần túy (pure functions).
  - [kisorlib/service.py](file:///D:/Antigravity/KisorDoc/kisorlib/service.py): Lớp nghiệp vụ `KisorService` loại bỏ global state `config`/`ds` thông qua Dependency Injection.
  - [kisorlib/batch.py](file:///D:/Antigravity/KisorDoc/kisorlib/batch.py): Hàm sinh tài liệu hàng loạt `run_batch` và `run_retry_batch` độc lập hoàn toàn với thư viện Gradio.
  - [app.py](file:///D:/Antigravity/KisorDoc/app.py): Rút gọn xuống chỉ còn ~360 dòng đóng vai trò kết nối UI Gradio.
- **Dọn dẹp mã trùng lặp trong engine.py:** Xóa bỏ các hàm clone trùng lặp trước đó (`_clean_config_key`, `_get_option_config_from_ds`) trong `kisorlib/engine.py` và tái sử dụng trực tiếp từ `utils.py` và `service.py`.

### Added
- **Unit Tests tự động:** Bổ sung thư mục kiểm thử `tests/` với các bộ test tự động [tests/test_utils.py](file:///D:/Antigravity/KisorDoc/tests/test_utils.py) và [tests/test_filters.py](file:///D:/Antigravity/KisorDoc/tests/test_filters.py) giúp chạy hồi quy nhanh chóng.

### Fixed
- **Sửa lỗi Preview ở chế độ Repeat:** Khắc phục triệt để lỗi phân tích sai ID gói thầu dạng composite (`key_id` chứa `|`) trong chức năng Preview (`run_preview`).

## [3.2.3] - 2026-08-04

### Fixed
- **Tương thích hoàn toàn parser cho Join Expression:** Viết lại hàm `_parse_repeat_sheet_config` sử dụng cùng bảng ánh xạ toán tử `_OP_MAP` với `parse_join_expression`. Fix triệt để lỗi phân tích sai tên sheet trái/phải (`left_sheet`, `right_sheet`) đối với các cú pháp join phức tạp như `<*>`, `<*`, `*>`, `*`.
- **An toàn hóa contract write_with_retry:** Chuẩn hóa kiểu trả về của hàm `write_with_retry` luôn trả về tuple `(bool, str)`. Ngăn chặn hoàn toàn lỗi runtime `TypeError: cannot unpack non-iterable` khi hàm con bên trong không trả về tuple (ví dụ: `do_copy` trả về `None`).
- **Khắc phục stale data trong chế độ Repeat:** Thêm guard clause kiểm tra kết quả trả về của hàm `register_temporary_tcgttd` trước khi thực hiện câu lệnh SQL chính. Bỏ qua và ghi log `SKIP` cho thành viên bị lỗi thay vì truy vấn đè lên dữ liệu cũ của thành viên trước đó.
- **Đồng bộ split KeyId:** Cập nhật các hàm thay đổi dữ liệu trên giao diện (`on_package_change`, `on_group_change`) để phân tách KeyId thông qua `_parse_repeat_key_id` trước khi truy xuất giá trị từ dict, sửa lỗi lấy rỗng `goi_thau_id` khi cấu hình KeyId dạng ghép.
- **Loại bỏ triệt để hardcode còn lại:** Chuyển đổi toàn bộ các fallback gán cứng tên sheet `"GoiThau"`, tên cột `"ID"`, dạng show `"{TT}"` sang sử dụng giá trị cấu hình tương ứng trong `AppConfig` (`config.DataSheet`, `config.DefaultKeyId`, `config.DefaultShow`). Giao diện chọn nhóm lặp cũng được động hóa không còn cứng `"Tổ chuyên gia" / "Tổ thẩm định"`.

## [3.2.2] - 2026-08-04

### Added
- **Động hóa hoàn toàn cấu hình `KeyId` ghép `|` cho chế độ Repeat:** Hỗ trợ cấu hình `KeyId` dạng ghép bằng dấu `|` (Ví dụ: `GoiThau_ID | CCCD`), tự động phân tách thành `left_key` (khóa bảng chính) và `right_key` (khóa bảng con) để JOIN query chính xác tuyệt đối, tránh trùng lặp họ tên thành viên khi xử lý lặp.
- **Trích xuất tên cột Họ tên động từ cột `Show`:** Thay vì gán cứng `"Họ và tên"`, chương trình tự động trích xuất tên cột định danh thành viên từ phần bên phải của cột `Show` (sau dấu `|`) để đưa giá trị chuẩn vào file template Word.
- **Nhân bản bảng DuckDB tránh ghi đè dữ liệu gốc:** Tự động tạo bản sao dự phòng `_Goc` cho toàn bộ các bảng trong DuckDB khi khởi động, giúp các quy trình xử lý lặp song song hoặc tuần lặp không bị ghi đè hay mất dữ liệu gốc của sheet Excel.

### Changed
- **Default Show Format an toàn:** Chuyển giá trị mặc định của `show` trong `get_option_config` khi không có cấu hình thành `"{TT}"` thay vì hardcode cột của gói thầu cụ thể.
- **Động hóa `left_sheet` cho Repeat:** Cập nhật `get_packages` để truy xuất bảng chính thông qua `left_sheet` lấy từ cột `Sheet` trong Options thay vì cứng nhắc `"GoiThau"`.
- **Động hóa `DANH_MUC_FILE` từ `.env`:** Di chuyển cấu hình tên file Danh Mục dự án thành biến môi trường `DANH_MUC_FILE` để thuận tiện tùy biến.
- **Vô hiệu hóa tự động bật trình duyệt web:** Tắt cơ chế tự động gọi `webbrowser.open` tại `runner.py` và `app.py` khi khởi động/nạp lại code để tránh mở tab rác trên trình duyệt của người dùng.

## [3.2.1] - 2026-08-03

### Changed
- **Tái cấu trúc thư viện dùng chung (Patch-v8):** Di chuyển các class/hàm helper `NestedVal` và `make_nested_dict` từ `app.py` vào module dùng chung `kisorlib/app_helpers.py` để cả `app.py` và `engine.py` cùng chia sẻ, giảm thiểu trùng lặp mã và tăng độ ổn định của hệ thống.
- Cập nhật và tối ưu hóa an toàn relative imports bên trong core library `kisorlib`.

## [3.2.0] - 2026-08-03

### Changed
- **Đổi tên Core Package (`kisorlib`):** Đổi tên toàn bộ thư mục thư viện dùng chung từ `kisordoc/` thành `kisorlib/` để tăng tính chuyên nghiệp, đồng thời cập nhật toàn bộ import trong `app.py`, `api.py`, `runner.py`, và `tests/test_engine.py`.
- **Định dạng Date khoảng trắng chừa trống:** Cố định **3 khoảng trắng** cho ngày và tháng khi bị trống dữ liệu trên Excel (Ví dụ: `"ngày   tháng 07 năm 2026"`, `"ngày   tháng   năm 2026"`).

## [3.1.0] - 2026-08-03

### Added
- **Xử lý ngày tháng trống một phần (Chừa khoảng trống ghi tay):** Bổ sung logic xử lý cho filter ngày tháng (như `.Date.Long`), tự động định dạng các chuỗi ngày tháng chứa dấu gạch chéo `/` nhưng bị khuyết thông tin ngày hoặc tháng (VD: `"   /07/2026"`, `"  /   /2026"`) thành `"ngày   tháng 07 năm 2026"` và `"ngày   tháng   năm 2026"`, mặc định sử dụng 3 khoảng trắng cho phần bị trống để ghi tay sau.
- **Cải tiến giao diện:** Cấu hình cột template bên phải luôn hiển thị (`visible=True`) ngay từ đầu để tránh co giãn layout, đồng thời giữ nguyên logic cập nhật danh sách động khi người dùng chọn Quy trình & Gói thầu.

### Fixed
- Khắc phục lỗi `Binder Error` khi chạy quy trình lặp `Repeat` (`Opt6`) do câu lệnh SQL Join bị rỗng lúc chưa nạp danh sách thành viên tạm thời.
- Khắc phục lỗi thiếu thư viện `os` khi nhấn nút mở thư mục log/output trên giao diện.

## [3.0.0] - 2026-08-03

### Added
- **Core Library Packaging (`kisordoc/`):** Tái cấu trúc đóng gói toàn bộ logic xử lý nghiệp vụ (`config.py`, `dataset.py`, `table_copier.py`, `merger.py`, `file_utils.py`, `filters.py`) vào package `kisordoc/`.
- **Core Engine API (`kisordoc/engine.py`):** Xây dựng điểm vào duy nhất (Public API) cho các tác vụ mail-merge và dry-run sử dụng Pydantic models (`GenerateRequest`/`GenerateResult`) và cơ chế callback tiến trình `on_progress`.
- **Tích hợp FastAPI Backend (`api.py`):** Cung cấp các RESTful API endpoints `/generate`, `/templates`, `/packages` tự động sinh tài liệu Swagger.
- **Khởi chạy song song (`runner.py`):** Hỗ trợ khởi chạy đồng thời Gradio UI (`app.py` ở cổng 7864) và FastAPI API (`api.py` ở cổng 8000) thông qua thread an toàn chỉ bằng một lệnh duy nhất.
- **Tách biệt cấu hình nhãn giao diện (`ui_labels.json`):** Chuyển toàn bộ chuỗi ký tự hiển thị trên Gradio UI ra file cấu hình JSON độc lập giúp thay đổi nhãn động không cần sửa code.
- **Cải tiến giao diện chọn template:** Tự động ẩn cột chọn file template khi chưa chọn gói thầu và chỉ hiển thị sau khi đã nạp dữ liệu thành công.

### Fixed
- Khắc phục lỗi `❌ Không tìm thấy dòng dữ liệu tương ứng` khi chạy batch hoặc dry-run đối với các Option đặc thù (như Mua sắm nhỏ Opt1 lấy từ bảng `MuaSamNho`).
- Khắc phục lỗi crash do component Textbox của phiên bản Gradio mới không hỗ trợ tham số `show_copy_button`.
- Sửa lỗi nút Mở thư mục log/output trên Windows console bằng cách chuyển sang `subprocess.Popen`.

## [2.2.2] - 2026-08-03

### Added
- **Xử lý File Locked nâng cao (F6):** Bắt lỗi `PermissionError` (khi Word chiếm dụng file), tự động retry tối đa `FILE_MAX_RETRIES` lần sau mỗi `FILE_RETRY_DELAY` giây. Đánh dấu trạng thái file bị khóa dạng `🔒` thay vì `❌`.
- **Validation trước khi chạy (F1):** Chặn sớm việc chạy batch và thông báo lỗi rõ ràng nếu thiếu Option, Gói thầu, Template hoặc dữ liệu chưa sẵn sàng.
- **Dry-run / Preview Mode (F2):** Bổ sung nút "🔍 Kiểm tra" trên UI, trích xuất placeholder an toàn từ ZIP Docx, áp dụng custom filters và hiển thị bảng kết quả Preview (`gr.Dataframe`) trực quan.
- **Lưu trạng thái & Chạy lại file lỗi (F3):** Tự động phát hiện lỗi và hiển thị nút "Chạy lại file lỗi" để chỉ merge lại các file bị lỗi mà không cần xóa/xử lý lại các file thành công.
- **Export log ra file text (F4):** Ghi log incremental khi chạy, định dạng `utf-8-sig` chuẩn Notepad Windows, tự động dọn dẹp log cũ >30 ngày (giới hạn 100 file log).
- **Lộ trình Refactor codebase:** Thống nhất kế hoạch đóng gói thư viện `kisordoc/`, FastAPI endpoints, và script khởi chạy song song `runner.py`.

## [2.2.0] - 2026-07-31

### Added
- **Quy trình Lặp (Repeat Type Options):** Hỗ trợ chạy hàng loạt nhiều dòng dữ liệu cho 1 file template thông qua bộ nhận diện `Type` = `Repeat` trong sheet `Options` (Ví dụ: xuất cam kết cho từng thành viên của Tổ chuyên gia/Tổ thẩm định).
- **Bộ chọn nhóm & thành viên động:** Tích hợp radio chọn nhóm ("Tổ chuyên gia" / "Tổ thẩm định") và load danh sách thành viên động từ sheet `S.TCGTTD` của gói thầu lên checkbox để người dùng chọn người cần xuất tài liệu.
- **Liên kết động bằng Họ tên:** Tự động kết nối dữ liệu chi tiết của thành viên trong sheet `S.TCGTTD` của gói thầu với bảng dữ liệu dùng chung `TCGTTD` bằng so khớp Họ tên, sau đó gán `GoiThau_ID` động để thực hiện phép Join của hệ thống.

### Changed
- **Sửa giá trị khóa mặc định:** Thay đổi giá trị fallback mặc định của `key_id` từ `"GoiThau_ID"` thành `"ID"` giúp hệ thống linh hoạt hơn khi cấu hình.

## [2.1.0] - 2026-07-31

### Added
- **Liên kết bảng (Join Sheets):** Hỗ trợ liên kết 2 bảng bằng ký hiệu rút gọn (`<*`, `*>`, `*`, `<*>`) và trên 3 bảng bằng câu lệnh SQL trực tiếp (`SELECT ...`).
- **Gộp sheet trùng tên tự động:** Tự động gộp dòng dữ liệu từ các file Excel khác nhau khi phát hiện có sheet trùng tên (Ví dụ: `Tables` trong `Tables.xlsx` và `DanhMuc-MSSC.xlsx`).
- **Nguồn file Excel động cho Tables:** Thêm cột `File` trong sheet `Tables` để tùy biến file nguồn copy bảng (Ví dụ: `S.Oto.xlsx`).
- **Nhân bản bảng biểu tự động:** Tự động copy lặp lại bảng dữ liệu cuối cùng khi số placeholder trong Word nhiều hơn dòng cấu hình Excel.
- **Cảnh báo trùng tên cột (Column Collision Warning):** Hiển thị cảnh báo trực quan trên log UI khi các cột bị trùng tên trong quá trình Join.

### Fixed
- Khắc phục lỗi so khớp tên file chứa số thập phân (Ví dụ: `9.1 BC tham dinh ...`) khi chèn bảng biểu.
- Sửa lỗi gộp ô (merge cell) khi remap chỉ số dòng master row bị sai tọa độ trong `table_copier.py`.
- Khắc phục lỗi crash validation của Gradio (`choices=[]`) khi chuyển đổi quy trình hoặc gói thầu trên UI.
- Tối ưu hóa bỏ qua các file `S.*` lúc khởi động giúp tiết kiệm 90% dung lượng RAM và tăng tốc app.

## [2.0.1] - 2026-07-31

### Fixed
- Khắc phục hoàn toàn lỗi cảnh báo thiếu dữ liệu (`Warning: Placeholder ... không có data`) bằng cách tự động ánh xạ đuôi `.Date` thành hậu tố `_Date` trong cả logic nạp context (`clean_config_key`) của `main.py` và script `migrate_modifiers.py`.
- Sửa lỗi xung đột (collision) ghi đè giữa Số quyết định (ví dụ: `KHLCNT_QD`) và Ngày quyết định (ví dụ: `KHLCNT_QD_Date`).

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
