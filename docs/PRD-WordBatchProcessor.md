# PRD – Word Batch Processor (KisorDoc-AI)
**Phiên bản:** 2.2  
**Ngày:** 2026-07-28  
**Trạng thái:** Production Ready (ver1.3)

---

## 1. Tổng quan

Tool Python thay thế UiPath xử lý hàng loạt file Word (`*-Template.docx`) với 2 tính năng chính:

1. **Mail Merge** – Replace placeholder `<<TenBien>>` (hoặc `<<TenBien.Modifier>>`) trong Word bằng dữ liệu từ DataSet (quét từ tất cả file Excel trong thư mục Data, trừ sheet bắt đầu bằng `S.`). Hỗ trợ modifier: `.Date`, `.Date.Long`, `.Upper`, `.Number`.
2. **Copy bảng Excel → Word** – Copy vùng range từ Excel vào đúng vị trí placeholder trong Word dưới dạng **bảng Word thật** (có thể edit trực tiếp trong Word như bảng thông thường, không phải OLE object hay ảnh), giữ tối đa định dạng gốc: merged cell, border, font, màu nền, chiều rộng cột, chiều cao hàng. Hỗ trợ ẩn cột linh hoạt theo từng loại tài liệu.

**Nguyên tắc tối ưu cốt lõi:** Scan placeholder trong Word trước → chỉ truy vấn đúng dữ liệu cần thiết (không load thừa).

---

## 2. Cấu trúc thư mục

### 2.1 Project root

Dự án làm việc với **1 thư mục gốc** (project path), cấu trúc:

```
{ProjectPath}/
├── 1. Data/                          # File Excel nguồn (tất cả file .xlsx)
│   ├── DanhMuc-MSSC.xlsx
│   └── TongHop-MSSC.xlsx
│
├── 2. Templates/                     # File Word template, phân nhóm theo Option
│   ├── Opt1/                         # Nhóm "Các giấy tờ đến bước Hợp đồng"
│   │   ├── 0. Danh muc.A-Template.docx
│   │   ├── 0. Danh muc.B-Template.docx
│   │   ├── ...
│   │   └── 14. Hop dong.XD-Template.docx
│   └── Opt2/                         # Nhóm "Nghiệm thu, Thanh toán"
│       ├── 15. BBBG-Template.docx
│       ├── 15. BBBG.XD-Template.docx
│       ├── ...
│       └── 19. QD ban giao tai san-Template.docx
│
├── 3. Files/                         # Output – tự xóa toàn bộ trước mỗi lần chạy mới
│   └── (file output được đặt ở đây sau khi xử lý)
│
└── Config-5.txt                      # File cấu hình JSON (nằm ở %LOCALAPPDATA%\UiPathProjectConfigs\)
```

> **Quy tắc đặt tên template:** `{Số thứ tự}. {Tên mô tả}-Template.docx` (VD: `3. Yeu cau bao gia-Template.docx`). Các template có đuôi `-Template.docx` mới được xử lý.

### 2.2 File cấu hình (Config-5.txt)

File JSON đặt tại `%LOCALAPPDATA%\UiPathProjectConfigs\Config-5.txt`:

| Key | Giá trị | Mô tả |
|---|---|---|
| `ProjectPath` | `D:\Antigravity\1. Thanh toan nho` | Đường dẫn thư mục dự án |
| `OnlineMode` | `Disable` | `Enable` = lấy dữ liệu từ Google Drive |
| `DataFolder` | `1. Data` | Tên thư mục chứa file Excel |
| `TemplateFolder` | `2. Templates` | Tên thư mục chứa Word template |
| `FileFolder` | `3. Files` | Tên thư mục output |
| `DataSheet` | `GoiThau` | Tên sheet chính chứa dữ liệu gói thầu |
| `CloseWord` | `false` | Đóng Word sau khi xử lý |
| `TaskManagerProcess` | `WINWORD.exe` | Tiến trình Word cần kill |
| `AgentPath` | `C:\FPT\EGP-AGENT\egp-edoc-agent.exe` | (Không dùng trong Python) |
| `ExceptionSheet` | `S.` | Tiền tố sheet bị loại khỏi DataSet |
| `AppName` | `KisorDoc-AI` | Tên ứng dụng hiển thị trên UI |

---

## 3. DataSet – Cấu trúc dữ liệu

Khi bắt đầu, tool quét **tất cả file `.xlsx`** trong thư mục Data và load **tất cả sheet** (trừ sheet có tên bắt đầu bằng `S.`) vào DataSet. Mỗi sheet trở thành 1 `DataTable`, tên sheet là tên table.

### 3.1 Danh sách sheet (table) trong DataSet

| Sheet (Table) | Nguồn file | Vai trò |
|---|---|---|
| `GoiThau` | File Excel bất kỳ | Dữ liệu chính của gói thầu (1 row = 1 gói thầu) |
| `Options` | File Excel bất kỳ | Danh sách Option cho UI chọn quy trình |
| `Workflow` | File Excel bất kỳ | Định nghĩa template theo option, giá, loại |
| `Tables` | File Excel bất kỳ | Định nghĩa copy bảng Excel → Word |
| `Config` | File Excel bất kỳ | Mapping placeholder key ↔ tên cột trong GoiThau |
| `S.*` (exception) | **Bị loại** | Sheet bắt đầu bằng `S.` không được đưa vào DataSet |

### 3.2 Sheet `GoiThau` – Dữ liệu gói thầu

1 row = 1 gói thầu. Các cột chính (tham khảo, có thể thay đổi tùy file Excel):

| Cột | Ví dụ | Ghi chú |
|---|---|---|
| `TT` | 1 | Số thứ tự |
| `GoiThau_ID` | MS26-01 | ID duy nhất |
| `GoiThau_Loai` | MSHH / XD / TUVAN / PHITUVAN | Loại gói thầu |
| `Số hiệu gói thầu` | XLNT | |
| `Chủ đầu tư` | Bệnh viện Sản Nhi tỉnh Quảng Ninh | |
| `Chủ đầu tư viết tắt` | Bệnh viện Sản-Nhi | |
| `Địa chỉ CĐT` | phường Tuần Châu, tỉnh Quảng Ninh | |
| ... | ... | Nhiều cột khác (xem Data/config.md để biết mapping) |

### 3.3 Sheet `Options` – Danh sách quy trình

| Cột | Ví dụ | Mô tả |
|---|---|---|
| `Key` | Opt1 | Mã option |
| `Value` | Các giấy tờ đến bước Hợp đồng | Tên hiển thị trên UI |

### 3.4 Sheet `Workflow` – Định nghĩa template theo option

| Cột | Ví dụ | Mô tả |
|---|---|---|
| `Key` | 1 | Thứ tự |
| `Option` | Opt1 | Thuộc nhóm option nào (để trống = chung cả 2 option) |
| `Name` | Bảng kê hồ sơ-trên 50 triệu | Tên hiển thị checkbox |
| `File` | 0. Danh muc.A-Template | Tên file (không cần đuôi `.docx`, không cần `-Template`) |
| `Price` | 50.000.000 | Giá tối thiểu |
| `PriceMax` | 500.000.000 | Giá tối đa |
| `Type` | ALL / XD / MSHH / TUVAN / PHITUVAN | Lọc theo loại gói thầu |

> **Quy tắc lọc:** Tool lấy giá gói thầu (`dblPrice`), chỉ hiển thị các template có `Price <= dblPrice <= PriceMax` và `Type` = `ALL` hoặc trùng với `GoiThau_Loai` của gói thầu đã chọn.

### 3.5 Sheet `Tables` – Định nghĩa copy bảng Excel → Word

| Cột | Ví dụ | Mô tả |
|---|---|---|
| `GoiThau_ID` | MS26-01 | ID gói thầu |
| `Word` | Yeu cau bao gia | Tên file Word (không đuôi, không `-Template`) để match |
| `Sheet` | S.XLNT / (để trống) | Tên sheet nguồn chứa bảng cần copy (có tiền tố `S.` — đây là sheet exception data) |
| `Range` | A1:E2 hoặc A1:F5 | Vùng range cần copy |
| `Hide` | (để trống) | Cột bị ẩn (tên header) |
| `Name` | {DanhMucKoGia} hoặc {DanhMuc} | Placeholder trong Word đánh dấu vị trí chèn bảng |

> **Lưu ý:** Bảng nguồn nằm trong sheet `S.*` (exception sheet), nhưng tool vẫn đọc trực tiếp từ file Excel để copy — không qua DataSet.

### 3.6 Sheet `Config` – Mapping placeholder ↔ cột dữ liệu

| Cột | Ví dụ | Mô tả |
|---|---|---|
| `Key` | `<<CDT_Ten>>` | Placeholder key trong Word |
| `Value` | `Chủ đầu tư` | Tên cột trong sheet GoiThau để lấy giá trị |

Các modifier được xử lý tự động:
- `<<TenBien.Date>>` → parse cột `TenBien` thành ngày `dd/MM/yyyy`
- `<<TenBien.Date.Long>>` → parse thành "ngày dd tháng MM năm yyyy"
- `<<TenBien.Upper>>` → uppercase giá trị cột `TenBien`
- `<<TenBien.Number>>` → format số với dấu `.` nghìn

### 3.7 Sheet exception (`S.*`)

Các sheet bắt đầu bằng `S.` không được load vào DataSet. Chúng chứa dữ liệu bảng cố định dùng cho copy vào Word (thông qua sheet `Tables`). VD: `S.XLNT`, `S.HC2026`, `S.DoDa`, `S.Sipap`, `S.ThueLS`, `S.BDGettinge`, `S.BaoHiem`, `S.BaoHiem`...

---

## 4. Placeholder

- **Định dạng trong Word template:** `{{TenBien}}` hoặc `{{TenBien|modifier}}` — chuyển từ `<<>>` sang `{{}}` để tương thích `docxtpl`/Jinja2
- **Migration:** Các template `.docx` hiện tại dùng `<<TenBien>>` → cần **convert 1 lần** sang `{{TenBien}}` (script tự động)
- **Vị trí hỗ trợ:** body paragraphs, bảng (kể cả nested table), header, footer — `docxtpl` xử lý tất cả
- **Hai loại placeholder:**
  - **Text placeholder:** `{{HoTen}}`, `{{NgayKy|date}}` → replace bằng giá trị text (kèm Jinja2 filter)
  - **Table placeholder:** `{{BangBaoGia}}` → toàn bộ đoạn (paragraph) chứa placeholder bị thay thế bằng bảng Word

> ⚠️ **Phương án thay thế nếu không muốn đổi template:** Giữ nguyên `<<TenBien>>` và tự viết replace với `python-docx` (như thiết kế cũ). Tuy nhiên đổi sang `{{}}` sẽ tận dụng được toàn bộ sức mạnh `docxtpl`: xử lý runs split, filter, loop trong bảng.

### 4.1 Modifier System (Jinja2 Filter)

Modifier được implement dưới dạng **custom Jinja2 filter** trong `docxtpl`. Syntax đổi từ `<<TenBien.Modifier>>` sang `{{TenBien|modifier}}`.

| Modifier cũ (UiPath) | Syntax mới (Jinja2) | Hành vi | Ví dụ output |
|---|---|---|---|
| `<<NgayKy.Date>>` | `{{NgayKy\|date}}` | Format ngày `dd/MM/yyyy`. Rỗng → ` /MM/yyyy` hiện tại | `01/07/2026` |
| `<<NgayKy.Date.Long>>` | `{{NgayKy\|date_long}}` | Ngày tiếng Việt. Rỗng → tháng/năm hiện tại | `ngày 01 tháng 07 năm 2026` |
| `<<TenCongTy.Upper>>` | `{{TenCongTy\|upper}}` | Chữ hoa (Jinja2 built-in, không cần custom) | `CÔNG TY ABC` |
| `<<GiaTri.Number>>` | `{{GiaTri\|number}}` | Dấu `.` nghìn, không thập phân | `1.500.000` |

**Implementation:**
```python
from docxtpl import DocxTemplate
import jinja2

def filter_date(value):
    # parse dd/MM/yyyy hoặc MM/dd/yyyy, fallback ngày hiện tại
    ...

def filter_date_long(value):
    # "ngày DD tháng MM năm YYYY"
    ...

def filter_number(value):
    # format 1.500.000, raise nếu không phải số
    ...

tpl = DocxTemplate("template.docx")
jenv = jinja2.Environment()
jenv.filters["date"]      = filter_date
jenv.filters["date_long"] = filter_date_long
jenv.filters["number"]    = filter_number
# "upper" là built-in của Jinja2, không cần đăng ký

tpl.render(context, jenv)
tpl.save("output.docx")
```

**Xử lý lỗi:**
- `|date` mà không parse được → `⚠️ Warning` + giữ nguyên giá trị gốc
- `|number` mà không phải số → `❌ Lỗi` + dừng file đó
- `|date` mà rỗng → lấy ngày hiện tại

---

## 5. Quy trình chạy

> **Luồng duy nhất:** 1 lần chạy = 1 gói thầu (chọn từ danh sách) → chọn template → xử lý.

**Luồng chi tiết:**
```
[1] Load DataSet: duckdb excel extension quét tất cả .xlsx trong 1. Data/
    → Mỗi sheet (trừ S.*) thành 1 DuckDB table:
      GoiThau, Options, Workflow, Tables, Config
    (Sheet S.* bỏ qua ở bước này — đọc riêng khi copy bảng)

[2] UI – Chọn Option (Opt1 / Opt2):
    • Hiển thị radio button từ DuckDB table Options
    • Nếu không chọn → báo lỗi "Select Option and click Submit"

[3] UI – Chọn Gói thầu:
    • Query: SELECT TT, GoiThau_ID, "Số hiệu gói thầu", "Tên gói thầu" FROM GoiThau
    • Hiển thị dạng: "{TT}. {Số hiệu gói thầu} - {Tên gói thầu}"

[4] UI – Chọn template file:
    • Query từ DuckDB table Workflow:
      WHERE Option = ? AND Price <= ? AND PriceMax >= ? AND (Type = 'ALL' OR Type = ?)
    • Hiển thị checkbox list (có "Chọn tất cả" + "Bỏ chọn tất cả")

[5] Xác nhận chạy
        ↓
Xóa toàn bộ 3. Files/
        ↓
Copy template từ 2. Templates/{Option}/ sang 3. Files/
  (chỉ copy file được checkbox)
        ↓
Với từng file Word trong 3. Files/ (tuần tự):
  • docxtpl scan template → lấy danh sách biến {{ }} cần thiết
  • Query DuckDB table Config để map key → tên cột trong GoiThau
  • Lấy giá trị từ row GoiThau đã chọn → build context dict
  • docxtpl render với Jinja2 custom filters (date, date_long, number)
  • Xử lý table placeholder: tra DuckDB table Tables
    → đọc sheet S.* bằng openpyxl → copy bảng (python-docx/lxml)
  • Save file (đổi tên: bỏ "-Template" + thêm GoiThau_ID)
        ↓
Báo cáo kết quả (runtime + open output folder nếu user chọn)
```

> **Lưu ý về đổi tên file output:** File copy từ Templates sang Files vẫn giữ tên `-Template.docx`. Sau khi xử lý xong → đổi tên thành `{tên}-{GoiThau_ID}.docx`.

---

## 6. Tính năng Copy bảng Excel → Word (chi tiết kỹ thuật)

### 6.1 Mục tiêu
Tái tạo kết quả tương đương **Paste Special → Microsoft Excel Worksheet Object** của UiPath nhưng dưới dạng **bảng Word thật** (có thể chỉnh sửa sau khi xuất), không phải OLE object hay ảnh.

### 6.2 Các thuộc tính định dạng cần giữ nguyên

| Nhóm | Thuộc tính | Nguồn từ Excel |
|---|---|---|
| **Cấu trúc** | Merged cells (span hàng/cột) | `cell.merged_cells`, `MergedCell` |
| **Cột** | Chiều rộng cột | `column_dimensions[col].width` |
| **Hàng** | Chiều cao hàng | `row_dimensions[row].height` |
| **Font** | Tên font, cỡ chữ, bold, italic, underline, màu chữ | `cell.font.*` |
| **Nền** | Màu nền ô (`fgColor`) | `cell.fill.fgColor` |
| **Border** | 4 cạnh: style, màu | `cell.border.{top,bottom,left,right}` |
| **Căn lề** | Ngang (left/center/right), dọc (top/center/bottom) | `cell.alignment.*` |
| **Wrap text** | Xuống dòng trong ô | `cell.alignment.wrap_text` |
| **Số** | Format hiển thị (ngày, tiền, %) | `cell.number_format` + giá trị thực |

### 6.3 Xử lý Merged Cells

Merged cells là phần phức tạp nhất, cần xử lý đặc biệt:

```
Excel merged range A1:C2 (ô A1 span 3 cột, 2 hàng)
        ↓
Word: ô A1 → grid_span=3 (merge ngang), vMerge=restart (merge dọc)
      ô B1, C1 → bỏ qua (thuộc merged range)
      ô A2 → vMerge=continue
      ô B2, C2 → bỏ qua (thuộc merged range)
```

**Thuật toán:**
1. Dùng `openpyxl` lấy toàn bộ `merged_cells` của sheet
2. Xây dựng lookup map: `(row, col) → (master_row, master_col, row_span, col_span)`
3. Khi render từng ô:
   - Nếu là master cell: tạo ô với `grid_span` và `vMerge=restart`
   - Nếu thuộc merged range nhưng không phải master: tạo ô trống với `vMerge=continue` (merge dọc) hoặc bỏ qua (merge ngang — Word tự xử lý qua `grid_span`)
   - Nếu là ô bình thường: tạo ô bình thường

### 6.4 Xử lý Hidden Columns

- Các cột trong `hidden_cols` bị loại khỏi range trước khi tạo bảng Word
- Nếu cột bị ẩn nằm trong merged range: toàn bộ merged range đó bị bỏ (không split)
- Chiều rộng các cột còn lại giữ nguyên tỷ lệ gốc từ Excel

### 6.5 Chuyển đổi đơn vị

| Thuộc tính | Đơn vị Excel | Đơn vị Word | Công thức |
|---|---|---|---|
| Chiều rộng cột | character units | DXA (twips) | `width * 96 * 914400 / 72 / 10000` (xấp xỉ `width * 7 * 20`) |
| Chiều cao hàng | points | DXA (twips) | `height * 20` |
| Cỡ chữ | points | Half-points | `size * 2` |

### 6.6 Border mapping Excel → Word

| Excel border style | Word border style |
|---|---|
| `thin` | `single` (1pt) |
| `medium` | `single` (2pt) |
| `thick` | `single` (3pt) |
| `double` | `double` |
| `dashed` | `dashed` |
| `dotted` | `dotted` |
| `hair` | `single` (0.5pt) |
| `None` / không có | `none` |

### 6.7 Format số

Giá trị ô trong Excel được format theo `number_format` trước khi ghi vào Word:

| Loại | Ví dụ Excel value | number_format | Hiển thị trong Word |
|---|---|---|---|
| Ngày | `46023` (serial) | `DD/MM/YYYY` | `01/07/2026` |
| Tiền | `1500000` | `#,##0` | `1,500,000` |
| Phần trăm | `0.15` | `0.00%` | `15.00%` |
| Text | `"ABC"` | `@` | `ABC` |

### 6.8 Luồng xử lý đầy đủ khi gặp table placeholder

```
docxtpl scan template → phát hiện {{BangChiTiet}} (table placeholder)
        ↓
Query DuckDB table Tables:
  WHERE GoiThau_ID = ? AND Word LIKE '%{tên file hiện tại}%'
  AND Name = '{BangChiTiet}'
  ORDER BY rowid  ← giữ thứ tự dòng trong sheet Tables

→ Lấy danh sách các occurrence theo thứ tự: dòng N → lần xuất hiện N
        ↓
Xử lý tuần tự từng occurrence (từ cuối document lên đầu
để tránh lệch vị trí XML sau mỗi lần insert):
  - Đọc file Excel trong 1. Data/ → sheet S.* tương ứng (openpyxl)
  - Parse Range (A1:F20 hoặc A1:F hoặc A1)
  - Lấy merged_cells, column_dimensions, row_dimensions
  - Xây dựng merged cells lookup map
  - Áp dụng cột Hide (ẩn cột chỉ định)
  - Tạo bảng Word với đầy đủ định dạng
  - Chèn bảng vào đúng vị trí {{BangChiTiet}}, xóa paragraph placeholder
        ↓
⚠️ Cảnh báo nếu số occurrence trong Word ≠ số dòng trong Tables
   (thừa hoặc thiếu dòng mapping)
```

---

## 7. Đặt tên file output

| Quy tắc | Chi tiết |
|---|---|
| Pattern | `{tên gốc bỏ "-Template"}-{GoiThau_ID}.docx` (VD: `5. QD phe duyet KH mua sam-MS26-01.docx`) |
| Làm sạch | Bỏ ký tự không hợp lệ: `/ \ : * ? " < > \|` |
| Cột id | Cố định = `GoiThau_ID` |
| Conflict | Tự thêm suffix `_1`, `_2`... nếu trùng tên trong cùng 1 lần chạy |
| Thư mục | Luôn xuất ra `Files/` (không tạo subfolder) |

---

## 8. Giao diện (Web app local – Gradio)

### Tab 1: Chọn quy trình & Gói thầu

**Tính năng:**
- Radio button chọn Option (Opt1 / Opt2) từ sheet Options
- Radio button chọn Gói thầu từ sheet GoiThau, sorted by TT (thứ tự)
- **✨ FIX #9:** Preview thông tin gói thầu (Tên CĐT, Giá, Loại, Số hiệu) hiển thị khi chọn
- Nút "📥 Tiếp theo" → load template và hiển thị feedback (✅ hoặc ❌)

**UX improvements (ver1.3):**
- Package preview tự update khi user chọn gói thầu khác
- Hiển thị trạng thái load ("✅ Đã tải 8 template từ '...'")

### Tab 2: Chọn file template

**Tính năng:**
- **✨ FIX #10:** Checkbox list các template đã lọc (theo option, giá, type), label hiển thị số template: "**Chọn template cần xử lý** (8 file)"
- Nút "✓ Chọn tất cả" / "✗ Bỏ chọn tất cả"
- Nút "🚀 Chạy" (size=lg, variant=primary)

**UX improvements (ver1.3):**
- Label tự update khi checkbox thay đổi
- Không cần nút "Back" — user chỉ cần click Tab 1 nếu muốn quay lại

### Tab 3: Log & Kết quả

**Tính năng:**
- **✨ FIX #4:** Chi tiết kết quả hiển thị trong Textbox (15-20 lines) thay vì Dataframe
  - Mỗi dòng: `✅ {Tên file} → {File output}` hoặc `❌ {Tên file}: {Lý do lỗi}`
- Status box: "Hoàn thành 5/8 file trong 12.3s"
- **✨ FIX #12:** Nút "📂 Mở thư mục output" chỉ visible sau khi chạy thành công
- **✨ FIX #8:** Nút "← Chạy lại" reset tất cả form + clear kết quả cũ → quay Tab 1

**UX improvements (ver1.3):**
- Textbox dễ copy log, clear khi chạy lại
- Nút open folder chỉ hiện lúc cần (tránh nhầm lẫn khi folder rỗng)

---

## 8.1 UX Improvements (ver1.3)

### Validation & Error Handling

**FIX #2 – Input Validation before Run:**
- User clicks "🚀 Chạy" → validation function checks:
  - ❌ Option không chọn → "❌ Vui lòng chọn quy trình"
  - ❌ Gói thầu không chọn → "❌ Vui lòng chọn gói thầu"
  - ❌ Template không chọn → "❌ Vui lòng chọn ít nhất 1 template"
- Nếu lỗi → hiển thị thông báo trong status box, không chạy

**FIX #3 – Processing State:**
- Khi nhấn "Chạy" → nút disable (trigger_mode="once")
- Tránh click 2 lần → 2 batch chạy song song → file corrupt
- Sau khi xong → nút re-enable

**FIX #9 – Package Data Preview:**
- Khi user chọn gói thầu → preview hiển thị:
  ```
  Tên CĐT: Bệnh viện Sản Nhi tỉnh Quảng Ninh
  Giá: 150.000.000
  Loại: XD
  Số hiệu: XLNT
  ```
- Tránh nhầm lẫn chọn gói sai

**FIX #5 – Package Sorting:**
- Danh sách gói thầu sort by TT (số thứ tự) — dễ tìm theo thứ tự

---

## 8.2 Bug Fixes (ver1.3+)

**NaT (Not a Time) Handling:**
- Fix lỗi `ValueError: NaTType does not support strftime` khi datetime field rỗng trong Excel
- Xử lý: check `pd.isna()` trước khi format ngày → nếu NaT → set giá trị = ""

---



| Loại | Nội dung |
|---|---|
| ✅ Thành công | File input → file output + thời gian xử lý |
| ❌ Lỗi | Tên file + lý do cụ thể |
| ⚠️ Warning | Placeholder `{{key}}` trong Word không có data tương ứng trong Config |
| ⚠️ Warning | Dòng trong sheet `Tables` không match file nào được chạy |
| ⚠️ Warning | Số lần xuất hiện `{{BangX}}` trong Word ≠ số dòng trong sheet `Tables` |
| ℹ️ Info | Key trong Config không được dùng trong template |

---

## 10. Cấu hình (Config-5.txt)

File JSON tại `%LOCALAPPDATA%\UiPathProjectConfigs\Config-5.txt`:

```json
{
  "ProjectPath": "D:\\Antigravity\\1. Thanh toan nho",
  "OnlineMode": "Disable",
  "DataFolder": "1. Data",
  "TemplateFolder": "2. Templates",
  "FileFolder": "3. Files",
  "DataSheet": "GoiThau",
  "CloseWord": "false",
  "TaskManagerProcess": "WINWORD.exe",
  "ExceptionSheet": "S.",
  "AppName": "KisorDoc-AI"
}
```

---

## 11. Stack kỹ thuật

| Thành phần | Thư viện | Phiên bản | Lý do |
|---|---|---|---|
| Mail merge Word | `docxtpl` | 0.20+ | Xử lý runs bị split sẵn; dùng Jinja2 filter cho modifier; không cần tự parse XML |
| Đọc/ghi Word (copy bảng) | `python-docx` + `lxml` | - | lxml cho insert bảng đúng vị trí placeholder, xử lý vMerge/gridSpan |
| Load DataSet (sheet data) | `duckdb` + extension `excel` | - | Đọc `.xlsx` trực tiếp bằng SQL, không cần openpyxl cho phần này; nhanh, ít RAM |
| Đọc Excel (copy bảng `S.*`) | `openpyxl` | read_only=False | Cần `merged_cells`, `column_dimensions`, `row_dimensions`, `cell.font` — duckdb không có |
| Giao diện | `Gradio` | 4.0+ | Web local, không cần cài Node/React; async support cho progress bar |
| Cấu hình | `.env` + `pydantic` | - | Validate schema config; dotenv load từ `.env` (không cần cấu hình hệ thống) |
| Đóng gói | `PyInstaller` | - | 1 file `.exe` cho deployment |
| Async framework | `asyncio` | - | Support async function trong Gradio (`async def run_batch`) |
| Pandas | `pandas` | - | Handle NaT (Not a Time), data manipulation |

> ⚠️ **Lưu ý phân chia openpyxl vs duckdb:**
> - **DuckDB excel extension** → dùng cho sheet data thông thường (GoiThau, Options, Workflow, Config, Tables): đọc nhanh, query SQL JOIN dễ.
> - **openpyxl** → chỉ dùng cho sheet `S.*` khi copy bảng sang Word: cần đọc đầy đủ formatting (merged cells, font, border, fill) mà duckdb không cung cấp.
> - `openpyxl` phải dùng `read_only=False` vì `merged_cells` không available ở read_only mode.

---

## 12. Rủi ro kỹ thuật & phương án dự phòng

| Rủi ro | Mức độ | Phương án |
|---|---|---|
| Merged cells phức tạp (merge chéo, lồng nhau) | Cao | Test kỹ với file Excel thực tế; fallback: render ô trống thay vì crash |
| Placeholder `{{key}}` bị Word split thành nhiều `<w:r>` runs | Thấp | `docxtpl` xử lý tự động — đây là lý do chính chọn thư viện này |
| Modifier `\|date` với nhiều định dạng ngày khác nhau trong Excel | Trung bình | Thử parse theo thứ tự: `dd/MM/yyyy` → `MM/dd/yyyy` → `MM/dd/yyyy hh:mm:ss` |
| Border style Excel không map được sang Word | Thấp | Dùng `single` làm fallback |
| File Excel dùng color theme (không phải RGB trực tiếp) | Trung bình | Resolve theme color qua `openpyxl` theme parser; fallback: bỏ qua màu nền |
| Chèn bảng đúng vị trí placeholder trong XML | Cao | Thao tác trực tiếp trên `lxml` element tree sau khi `docxtpl` render |
| DuckDB excel extension không đọc được sheet có tên tiếng Việt/đặc biệt | Trung bình | Test trước; fallback: dùng `openpyxl` đọc data sheet nếu cần |
| Migration template từ `<<>>` sang `{{}}` | Một lần | Script tự động find-replace trong XML của docx; cần review thủ công sau |

---

## 13. Script migration template (1 lần)

Khi chuyển từ UiPath sang Python, cần convert tất cả template từ `<<TenBien>>` → `{{TenBien}}` và `<<TenBien.Date>>` → `{{TenBien|date}}`:

```python
import zipfile, re, shutil, os

MODIFIER_MAP = {
    r"<<(\w+)\.Date\.Long>>": r"{{\1|date_long}}",
    r"<<(\w+)\.Date>>":       r"{{\1|date}}",
    r"<<(\w+)\.Upper>>":      r"{{\1|upper}}",
    r"<<(\w+)\.Number>>":     r"{{\1|number}}",
    r"<<(\w+)>>":             r"{{\1}}",         # plain, phải để cuối
}

def migrate_template(src_path, dst_path):
    """Convert <<>> → {{}} trong document.xml của .docx"""
    shutil.copy2(src_path, dst_path)
    # Đọc/ghi trực tiếp vào zip (docx là zip)
    with zipfile.ZipFile(dst_path, 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')
    for pattern, replacement in MODIFIER_MAP.items():
        xml = re.sub(pattern, replacement, xml)
    # Ghi lại vào zip
    ...
```

---

## 13. Kế thừa từ UiPath (KisorDoc-AI)

Dự án này là phiên bản Python của bot UiPath "KisorDoc-AI" (Word13.12_Windows). Các logic kế thừa:

| UiPath (cũ) | Python (mới) | Ghi chú |
|---|---|---|
| `Main.xaml` – Flowchart | Gradio web app | Không cần Flowchart |
| `ValidateConfig.xaml` – JSON config | Config-5.txt (giữ nguyên) | Giữ nguyên format và đường dẫn |
| `CollectDataSet.xaml` – Quét Excel | `duckdb` excel extension: `SELECT * FROM '*.xlsx'` | Nhanh hơn, không cần loop từng file |
| `CheckFileFolder.xaml` – Xóa/tạo folder | `shutil.rmtree` + `os.makedirs` | |
| `ProcessWordFileSilent.xaml` – Read Text + Regex | `docxtpl` scan template tìm `{{ }}` | docxtpl xử lý runs split tự động |
| `ProcessWordFileSilent.xaml` – Replace Text | `docxtpl` render với Jinja2 filter | Modifier `.Date/.Upper/.Number` → custom Jinja2 filter |
| `ProcessWordFileSilent.xaml` – Copy bảng | `openpyxl` → `python-docx`/`lxml` | Không cần VBA/COM |
| `VBA.txt` – Ẩn/hiện cột | Cột `Hide` trong sheet `Tables` | |
| `NPOI` assembly – đọc Excel | `duckdb` excel extension | Thay hoàn toàn |
| `PublicBidding.xaml` – Đăng tải web | **BỎ** (không dùng nữa) | |

**PRD hoàn chỉnh — sẵn sàng code.**
