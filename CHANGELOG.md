# Changelog

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
