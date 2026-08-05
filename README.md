# KisorDoc-AI

Công cụ Python xử lý hàng loạt tài liệu Word (Mail Merge & Copy bảng từ Excel sang Word), được xây dựng để thay thế hoàn toàn bot UiPath cũ.

## Tính năng chính
 
1. **Mail Trộn thư (Mail Merge):**
   - Hỗ trợ trộn dữ liệu từ Excel vào Word qua template định dạng Jinja2 `{{ TenBien }}`.
   - Hỗ trợ các bộ lọc định dạng: `|date` (dd/MM/yyyy), `|date_long` (ngày... tháng... năm...), `|number` (định dạng số phân tách hàng nghìn), `|upper` (chữ in hoa), `|num2text` (đọc số thành chữ).
   - Tự động phân tách và xử lý các biến văn bản độc lập với biến ngày tháng cùng tiền tố (Ví dụ: `KHLCNT_QD` và `KHLCNT_QD_Date`).
   - Tự động phân tích các biến lồng nhau dạng dấu chấm (Ví dụ: `{{KHLCNT_TTr.Dvi}}`, `{{DuToan.NguoiLap}}`).
   - **Phân vùng Config theo Option:** Giới hạn phạm vi đọc cấu hình mapping dòng (Ví dụ: `2-97`) cho từng Option để tránh trùng lặp placeholder chéo.

2. **Copy bảng Excel sang Word:**
   - Sao chép vùng dữ liệu (Range) từ Excel và chèn vào vị trí placeholder `{{DanhMuc}}` dạng bảng Word thật.
   - Giữ nguyên định dạng gốc: gộp ô (merged cells), màu nền, viền bảng, chiều cao/chiều rộng và căn lề.
   - **Nguồn file Excel động:** Cấu hình cột `File` trong sheet `Tables` để đọc trực tiếp từ nhiều file Excel khác nhau (Ví dụ: `S.Oto.xlsx`).
   - **Nhân bản bảng tự động:** Sao chép bảng dữ liệu cuối cùng khi số placeholder trong Word lớn hơn số dòng khai báo trong Excel.

3. **Liên kết bảng thông minh (Join Sheets):**
   - Hỗ trợ liên kết 2 bảng qua ký hiệu rút gọn (Ví dụ: `GoiThau <* TCGTTD @ GoiThau_ID` tương đương LEFT JOIN).
   - Hỗ trợ liên kết từ 3 bảng trở lên qua câu lệnh truy vấn SQL trực tiếp (`SELECT ...`).
   - **Gộp sheet trùng tên:** Tự động gộp dữ liệu từ nhiều file Excel khi phát hiện sheet trùng tên trong DataSet.
   - **Cảnh báo trùng tên cột:** Hiển thị cảnh báo trực quan trên log UI khi phát hiện các cột trùng tên giữa các bảng được join.

4. **Giao diện Web Local (Gradio) + REST API (FastAPI):**
   - UI Gradio thao tác trực quan qua 3 bước: Chọn Gói thầu -> Chọn template -> Chạy & Xem log.
   - FastAPI REST API (Swagger tự động tại `http://127.0.0.1:8000/docs`) với các endpoint `/generate`, `/templates`, `/packages`, `/jobs/*`.
   - Khởi chạy đồng thời cả hai qua một lệnh duy nhất `python runner.py`.

5. **Quy trình chạy lặp hàng loạt (Repeat Mode):**
   - Hỗ trợ chạy hàng loạt nhiều dòng dữ liệu cho 1 file template thông qua bộ nhận diện `Type` = `Repeat` trong sheet `Options` (Ví dụ: xuất cam kết cho từng thành viên của Tổ chuyên gia/Tổ thẩm định).
   - Tự động hiển thị bộ chọn nhóm (Tổ chuyên gia/Tổ thẩm định) và load danh sách thành viên động từ sheet `S.TCGTTD` của gói thầu đang chọn lên checkbox.
   - Liên kết động thông tin cá nhân chi tiết bằng Họ tên và gán ID liên kết trước khi thực hiện phép Join để sinh dữ liệu.

## Cấu trúc thư mục

### Thư mục dữ liệu (tạo tại `PROJECT_PATH`)

```text
{ProjectPath}/
├── 1. Data/            # Chứa các file dữ liệu Excel (.xlsx)
├── 2. Templates/       # Chứa các template Word (.docx), chia theo Opt1/Opt2
├── 3. Files/           # Thư mục đầu ra (Output)
```

### Cấu trúc mã nguồn (kể từ v4.0.0)

```text
KisorDoc/
├── runner.py           # Entry point: chạy song song Gradio UI (7864) + FastAPI (8000)
├── app.py              # UI Gradio (~360 dòng, chỉ kết nối giao diện)
├── api.py              # FastAPI REST API, Swagger tại /docs
├── kisorlib/           # Core library (độc lập với Gradio)
│   ├── config.py       # AppConfig (Pydantic) + load_config từ .env
│   ├── service.py      # KisorService: nghiệp vụ, Dependency Injection (không global state)
│   ├── engine.py       # Public API mail-merge & dry-run (Pydantic models)
│   ├── batch.py        # run_batch / run_retry_batch hàng loạt
│   ├── dataset.py      # Nạp Excel vào DuckDB (cache, join, gộp sheet trùng tên)
│   ├── table_copier.py # Copy bảng Excel → Word (giữ định dạng, merge cell)
│   ├── merger.py       # Mail merge Jinja2 + filters
│   ├── filters.py      # Bộ lọc định dạng (|date, |number, |num2text, ...)
│   ├── file_utils.py   # Xử lý file, retry khi bị Word chiếm dụng
│   ├── utils.py        # Hàm tiện ích thuần túy (pure functions)
│   └── app_helpers.py  # Helper dùng chung cho app.py & engine.py
├── tests/              # Unit tests tự động (test_utils, test_filters, ...)
└── ui_labels.json      # Nhãn giao diện có thể tùy chỉnh
```

## Hướng dẫn cài đặt và khởi chạy

1. Cài đặt các thư viện phụ thuộc:
   ```bash
   pip install -r requirements.txt
   ```
2. Copy `.env-example` thành `.env` và điền các biến môi trường phù hợp (xem hướng dẫn trong file).
3. Chạy ứng dụng:
   ```bash
   python runner.py
   ```
   Ứng dụng sẽ khởi chạy đồng thời:
   - Gradio UI: `http://127.0.0.1:7864`
   - FastAPI & Swagger docs: `http://127.0.0.1:8000/docs`

   Nếu muốn chạy độc lập từng phần:
   ```bash
   python app.py   # chỉ Gradio UI
   uvicorn api:app --host 0.0.0.0 --port 8000   # chỉ FastAPI API
   ```
