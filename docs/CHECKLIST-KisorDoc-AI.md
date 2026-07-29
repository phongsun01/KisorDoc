# Checklist – KisorDoc-AI (Python)
**Phiên bản:** 1.0 | **Ngày:** 2026-07-27

---

## Phase 0 – Setup & Migration (làm 1 lần)

### 0.1 Môi trường
- [x] Tạo virtual env Python 3.11+
- [x] Cài dependencies: `docxtpl`, `duckdb`, `openpyxl`, `gradio`, `pydantic`, `python-docx`, `lxml`
- [x] Tạo `requirements.txt`
- [x] Tạo cấu trúc thư mục project code:
  ```
  kisor_doc_ai/
  ├── main.py           # Entry point Gradio
  ├── config.py         # Load & validate Config-5.txt
  ├── dataset.py        # DuckDB DataSet loader
  ├── merger.py         # Mail merge
  ├── table_copier.py   # Copy bảng Excel → Word
  ├── file_utils.py     # Xóa/copy folder, đặt tên output
  ├── filters.py        # Custom filters
  └── migrate.py        # Script migration <<>> → {{}} (1 lần)
  ```

### 0.2 Script migration template
- [x] Viết `migrate.py`: convert `<<TenBien>>` → `{{TenBien}}`
- [x] Map modifier: `.Date` → `|date`, `.Date.Long` → `|date_long`, `.Upper` → `|upper`, `.Number` → `|number`
- [x] Thao tác trực tiếp trong `word/document.xml`, `word/header*.xml`, `word/footer*.xml` bên trong ZIP
- [ ] Test trên 2–3 template thực tế, review thủ công kết quả
- [ ] Chạy migrate toàn bộ thư mục `2. Templates/`
- [ ] Backup bản gốc `<<>>` trước khi migrate

---

## Phase 1 – Config & DataSet

### 1.1 Config loader (`config.py`)
- [x] Đọc `Config-5.txt` từ `%LOCALAPPDATA%\UiPathProjectConfigs\`
- [x] Validate schema bằng Pydantic (các key bắt buộc: `ProjectPath`, `DataFolder`, `TemplateFolder`, `FileFolder`, `ExceptionSheet`, `AppName`)
- [x] Expose các path đã resolve: `data_path`, `template_path`, `output_path`
- [x] Báo lỗi rõ ràng nếu file không tìm thấy hoặc thiếu key

### 1.2 DataSet loader (`dataset.py`)
- [x] Cài và test `duckdb` extension `excel`
- [x] Quét tất cả `*.xlsx` trong `1. Data/`
- [x] Load từng sheet (trừ sheet bắt đầu bằng `ExceptionSheet` = `S.`) vào DuckDB in-memory
- [x] Tên table = tên sheet (xử lý ký tự đặc biệt/tiếng Việt nếu cần)
- [x] Fallback: nếu DuckDB không đọc được sheet → dùng `openpyxl` + load thủ công
- [x] Expose hàm `query(sql)` trả về list of dict

---

## Phase 2 – Mail Merge

### 2.1 Jinja2 custom filters (`filters.py`)
- [x] `filter_date(value)`: parse chuỗi ngày → `dd/MM/yyyy`
  - [x] Thử theo thứ tự: `dd/MM/yyyy` → `MM/dd/yyyy` → timestamp số
  - [x] Nếu rỗng → trả `" /MM/yyyy"` với tháng/năm hiện tại
  - [x] Nếu không parse được → Warning + trả về giá trị gốc
- [x] `filter_date_long(value)`: → `"ngày DD tháng MM năm YYYY"`
  - [x] Nếu rỗng → `"tháng MM năm YYYY"` hiện tại
- [x] `filter_number(value)`: format số dấu `.` nghìn
  - [x] Nếu không phải số → raise Exception (dừng file đó)
- [x] `upper` dùng Jinja2 built-in (không cần custom)
- [ ] Unit test cho từng filter với các edge case

### 2.2 Mail merge engine (`merger.py`)
- [x] Nhận `template_path` + `context dict` + `output_path`
- [x] Dùng python-docx load template + replace trực tiếp
- [x] Phân biệt text placeholder vs table placeholder:
  - [x] Text placeholder: `{{key}}` hoặc `{{key|filter}}` → replace bằng giá trị
  - [x] Table placeholder: `{{DanhMuc}}`, `{{DanhMucKoGia}}` → bỏ qua, xử lý sau (Phase 3)
- [x] Xử lý header, footer, tables trong document
- [x] Bắt lỗi → log + dừng file đó (không crash cả batch)

### 2.3 Context builder (trong `main.py`)
- [x] Query DuckDB sheet `Config`: map variable → tên cột GoiThau
- [x] Query DuckDB sheet `GoiThau`: lấy row theo `GoiThau_ID` đã chọn
- [x] Build context dict: `{variable_name: value}`
- [x] Log các variable trong template nhưng không có trong Config → ⚠️ Warning (todo)
- [x] Log các key trong Config không được dùng trong template → ℹ️ Info (todo)

---

## Phase 3 – Copy bảng Excel → Word

### 3.1 Parse range (`table_copier.py`)
- [x] Parse chuỗi range:
  - [x] `A1:F20` → fixed: row 1–20, col A–F
  - [x] `A1:F` → col A–F, max row = `ws.max_row`
  - [x] `A1` → max row và max col = `ws.max_row`, `ws.max_column`
- [x] Validate range hợp lệ trước khi đọc

### 3.2 Đọc Excel với openpyxl
- [x] Mở file `openpyxl.load_workbook(path, read_only=False, data_only=True)`
- [x] Đọc `ws.merged_cells` → build merged cells lookup map
- [x] Đọc `ws.column_dimensions` → chiều rộng cột
- [x] Đọc `ws.row_dimensions` → chiều cao hàng
- [x] Đọc từng ô trong range: `value`, `font`, `fill`, `border`, `alignment`, `number_format`
- [x] Xử lý `hidden_cols`: ẩn cột theo tên header hoặc index

### 3.3 Tạo bảng Word
- [x] Tạo `python-docx` Table với đúng số hàng/cột (sau khi bỏ hidden_cols)
- [x] Với từng ô:
  - [x] **Merged cells:** gridSpan (merge ngang) và vMerge (merge dọc)
  - [x] **Chiều rộng cột:** set `w:tcW` qua lxml
  - [x] **Chiều cao hàng:** set `w:trHeight` qua lxml
  - [x] **Font:** name, size, bold, italic, underline, color (RGB)
  - [x] **Fill:** màu nền ô (`w:shd` fill)
  - [x] **Border:** 4 cạnh, map Excel style → Word style
  - [x] **Căn lề:** ngang (left/center/right) + dọc (top/center/bottom)
  - [x] **Giá trị:** format theo `number_format` (ngày, tiền, %, text)
- [ ] Unit test và kiểm tra thủ công kết quả

### 3.4 Chèn bảng vào Word
- [x] Sau khi mail merge xong → mở file output bằng `python-docx`
- [x] Scan document tìm tất cả paragraph chứa `{{DanhMuc}}`, `{{DanhMucKoGia}}`
  - [x] Thu thập danh sách occurrence theo thứ tự
- [x] Query DuckDB `Tables`: lấy danh sách dòng theo `GoiThau_ID` + tên file + `Name`
- [x] Validate: số occurrence == số dòng → nếu không → ⚠️ Warning
- [x] Xử lý từ **cuối document lên đầu** (tránh lệch vị trí XML)
  - [x] Lấy XML element của paragraph chứa placeholder
  - [x] Insert bảng Word vào vị trí đó bằng `lxml`
  - [x] Xóa paragraph placeholder
- [x] Save file

---

## Phase 4 – File & Folder management (`file_utils.py`)

- [x] Xóa toàn bộ `3. Files/` trước mỗi lần chạy mới (`shutil.rmtree` + `os.makedirs`)
- [x] Copy template từ `2. Templates/{Option}/` sang `3. Files/` (chỉ file được chọn)
- [x] Đặt tên output: `{tên gốc bỏ "-Template"}-{GoiThau_ID}.docx`
  - [x] Làm sạch ký tự không hợp lệ trong tên file
  - [x] Xử lý conflict: thêm suffix `_1`, `_2`...
- [x] Hàm mở thư mục output (`os.startfile` trên Windows)

---

## Phase 5 – Giao diện Gradio (`main.py`)

### 5.1 Tab 1 – Chọn Option & Gói thầu
- [x] Radio button Option (load từ DuckDB `Options`)
- [x] Radio/dropdown Gói thầu (load từ DuckDB `GoiThau`, hiển thị `TT + Số hiệu + Tên`)
- [x] Nút "Tiếp theo" → validate có chọn cả 2 → chuyển Tab 2
- [x] Nếu chưa chọn → hiển thị lỗi

### 5.2 Tab 2 – Chọn template
- [x] Checkbox list: load từ DuckDB `Workflow`, lọc theo Option + Price + Type
- [x] Nút "Chọn tất cả" / "Bỏ chọn tất cả"
- [x] Nút "← Quay lại"
- [x] Nút "Chạy" → validate có chọn ít nhất 1 file → trigger xử lý

### 5.3 Tab 3 – Log & Kết quả
- [x] Progress bar (`gr.Progress`)
- [x] Textbox log real-time (✅ / ⚠️ / ❌ từng file)
- [x] Hiển thị tổng: X file thành công / Y file lỗi / thời gian
- [x] Nút "Mở thư mục output" (`os.startfile`)
- [x] Nút "Chạy lại" → reset về Tab 1

---

## Phase 6 – Orchestrator (pipeline chính trong `main.py`)

- [x] Hàm `run_batch(option, goi_thau_id, selected_files)`:
  1. [x] Xóa `3. Files/`
  2. [x] Copy các template được chọn sang `3. Files/`
  3. [x] Với từng file (tuần tự):
     - [x] Build context dict từ GoiThau row
     - [x] Mail merge (`merger.py`)
     - [x] Copy bảng (`table_copier.py`)
     - [x] Đổi tên file output
     - [x] Cập nhật progress + log
  4. [x] Tổng hợp báo cáo
- [x] Xử lý exception từng file riêng biệt (1 file lỗi không dừng cả batch)

---

## Phase 7 – Đóng gói

- [ ] Test toàn bộ flow trên máy dev
- [ ] Test với file Excel/Word thực tế (ít nhất 3 template phức tạp)
- [ ] Test merged cells phức tạp (merge chéo)
- [ ] Test modifier: `.Date`, `.Date.Long`, `.Number`, `.Upper`
- [ ] Test edge cases: placeholder không có data, bảng không có mapping, range auto-detect
- [ ] Đóng gói bằng `PyInstaller` → 1 file `.exe`
- [ ] Test `.exe` trên máy không có Python

---

## Thứ tự ưu tiên code

```
Phase 0 (migration script)     ✅
    ↓
Phase 1 (config + dataset)    ✅
    ↓
Phase 2 (mail merge)          ✅
    ↓
Phase 4 (file utils)          ✅
    ↓
Phase 5 (UI Gradio)           ✅
    ↓
Phase 3 (copy bảng)           ✅
    ↓
Phase 6 (orchestrator)        ✅
    ↓
Phase 7 (đóng gói + test)     ⬜ Cần test thực tế
```

---

## Ghi chú kỹ thuật cần nhớ

| Điểm | Chi tiết |
|---|---|
| `openpyxl` merged cells | Phải dùng `read_only=False` — merged_cells không có ở read_only mode |
| Insert bảng vào Word | Phải thao tác `lxml` XML trực tiếp — python-docx không có API insert tại vị trí tùy ý |
| Thứ tự insert bảng | Luôn xử lý từ cuối document lên đầu để tránh lệch vị trí XML |
| DuckDB table name | Sheet tên tiếng Việt có thể cần quote: `SELECT * FROM "Gói thầu"` |
| Placeholder syntax | `{{key\|filter}}` — migration từ `<<key.Filter>>` |
| PyInstaller + Gradio | Cần thêm `--collect-all gradio` và `--hidden-import` cho một số dependency |

---

## Phase 8 – Tính năng bổ sung (Enhancements F1-F6)

- [ ] **F1 – Validation trước khi chạy**:
  - [ ] Thêm validation đồng bộ các bước Option, Gói thầu, Template trước khi chạy.
  - [ ] Hiển thị thông báo rõ ràng prefix `⚠️` trong `status_text` thay vì popup.
- [ ] **F2 – Dry-run / Preview mode**:
  - [ ] Thêm nút "🔍 Kiểm tra" trong Tab Chọn & Chạy.
  - [ ] Scan các variable và so khớp context mà không ghi đè file thật.
  - [ ] Hiển thị kết quả dry-run chi tiết theo dạng bảng (Có data / Thiếu data / Table OK).
- [ ] **F3 – Retry cho file lỗi**:
  - [ ] Hiển thị nút "🔄 Chạy lại file lỗi" nếu kết quả chạy có file ❌.
  - [ ] Chỉ xử lý lại các file bị lỗi, không xóa/overwrite các file đã xử lý thành công.
- [ ] **F4 – Export log ra file**:
  - [ ] Tạo thư mục `logs/` trong project path.
  - [ ] Ghi log incremental, đặt tên file log theo pattern thời gian, gói thầu và option.
  - [ ] Thêm nút "📋 Mở thư mục log" trên UI.
  - [ ] Tự động dọn dẹp các log cũ hơn 30 ngày.
- [ ] **F5 – Version pin cho Config/Tables**:
  - [ ] Ghi snapshot cấu hình Config và Tables vào cuối file log.
  - [ ] So sánh cấu hình hiện tại với lần chạy trước của cùng gói thầu và hiển thị banner cảnh báo nếu có diff.
- [ ] **F6 – Xử lý file đang mở (File Locked)**:
  - [ ] Bắt riêng lỗi `PermissionError` (lỗi 13 / Locked).
  - [ ] Tự động retry 3 lần, mỗi lần cách nhau 2 giây kèm cập nhật thông báo trạng thái.
  - [ ] Đổi tên nút Chạy lại thành "🔄 Chạy lại (đã đóng file chưa?)" nếu toàn bộ lỗi là file locked.

---

## Các lỗi cần sửa sau

- [ ] **Lỗi file Word của gói thầu DoDa (MS26-04) không mở được trên MS Word**:
  - [ ] Dù cấu trúc XML sinh ra hợp lệ và python-docx vẫn load thành công, MS Word vẫn báo lỗi không mở được file. Cần điều tra sâu hơn cấu trúc XML đặc thù (các thẻ merge ngang/dọc, theme màu, hoặc kiểu đường viền đặc biệt) của gói thầu DoDa để tìm ra nguyên nhân và sửa đổi.
- [ ] **Tối ưu hóa hệ số chuyển đổi cột Excel sang Word**:
  - [ ] Hệ số chuyển đổi độ rộng cột hiện được cấu hình động qua `.env` làm biến `EXCEL_TO_WORD_WIDTH_FACTOR`. Cần kiểm chứng với nhiều bảng biểu để tìm ra hệ số tối ưu nhất cho trang giấy A4.

