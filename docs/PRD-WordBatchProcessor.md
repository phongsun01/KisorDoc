# PRD – Word Batch Processor (KisorDoc-AI)
**Phiên bản:** 5.2.2  
**Ngày:** 2026-08-12  
**Trạng thái:** Production Ready (ver5.2.2)


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
| `DANH_MUC_FILE` | `DanhMuc` | Tên file danh mục dự án động (sử dụng thay thế so khớp cứng) |

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
| `Sheet` | S.GoiThau hoặc S.MayMuon | Tên sheet dữ liệu chính (KHÔNG dùng tiền tố `S.` để app load vào bộ nhớ) |
| `Show` | `{TT}. {Số hiệu gói thầu} - {Tên gói thầu}` | Định dạng nhãn hiển thị cho gói thầu (hỗ trợ `{Tên Cột/Tên Biến}`) |
| `KeyId` | GoiThau_ID | Cột khoá chính duy nhất của dữ liệu (mặc định là `ID` nếu trống) |
| `Config` | 2-97 | Vùng dòng trong sheet Config thuộc về Option này (để trống = đọc toàn bộ sheet) |
| `Type` | Repeat | Loại quy trình. Nếu là `Repeat` thì hiểu là chạy loop nhiều dòng (ví dụ thành viên) cho 1 template. |
| `SortCol` | TT | Tên cột dùng để sắp xếp danh sách hiển thị theo thứ tự số nguyên. Để trống = không sắp xếp (an toàn với mọi sheet). |

#### Cấu hình nâng cao trong chế độ Repeat:
1. **Khớp nối Sheet (Join sheet):** Cột `Sheet` trong Options hỗ trợ biểu thức kết hợp bảng dạng `LeftSheet * RightSheet @ JoinKey` (Ví dụ: `GoiThau * TCGTTD @ GoiThau_ID`).
   - `LeftSheet` (ví dụ `GoiThau`): Chứa thông tin chung của gói thầu.
   - `RightSheet` (ví dụ `TCGTTD`): Chứa danh sách thành viên/đối tượng lặp.
   - `JoinKey` (ví dụ `GoiThau_ID`): Khóa liên kết dữ liệu giữa hai bảng.
2. **Khóa chính ghép (Composite KeyId):** Cột `KeyId` hỗ trợ dạng ghép bằng dấu `|` (Ví dụ: `GoiThau_ID | CCCD`) để định vị chính xác khóa bảng chính (`left_key` như `GoiThau_ID`) và khóa bảng con (`right_key` như `CCCD`), giúp tránh lỗi trùng họ tên thành viên khi lặp.
3. **Hiển thị nhãn (Composite Show):** Cột `Show` hỗ trợ dạng ghép bằng dấu `|` (Ví dụ: `{TT}. {Số hiệu gói thầu} - {Tên gói thầu} | {Họ và tên} - {CCCD}`).
   - Phần trước dấu `|` dùng để hiển thị gói thầu trên giao diện.
   - Phần sau dấu `|` dùng để hiển thị và phân biệt danh sách thành viên lặp. Cột đầu tiên trong dấu `{}` của phần này sẽ tự động được trích xuất làm tên cột chứa tên thành viên (Họ tên) để đưa giá trị thô vào template Word.

### 3.4 Sheet `Workflow` – Định nghĩa template theo option

| Cột | Ví dụ | Mô tả |
|---|---|---|
| `Key` | 1 | Thứ tự |
| `Option` | Opt1 | Thuộc nhóm option nào (để trống = chung cả 2 option) |
| `Name` | Bảng kê hồ sơ-trên 50 triệu | Tên hiển thị checkbox |
| `File` | 0. Danh muc.A-Template | Tên file (không cần đuôi `.docx`, không cần `-Template`) |
| `Price` | 50.000.000 | Giá tối thiểu (tương thích ngược) |
| `PriceMax` | 500.000.000 | Giá tối đa (tương thích ngược) |
| `Type` | ALL / XD / MSHH / TUVAN / PHITUVAN | Lọc theo loại gói thầu (tương thích ngược) |
| `Condition` | `{Giá gói thầu} <= 500000000 and {GoiThau_Loai} == 'MSHH'` | Biểu thức logic Python lọc template động (ưu tiên cao nhất) |

> **Quy tắc lọc:** 
> 1. Nếu có cột `Condition` và giá trị không trống/ALL: Biểu thức logic trong `{}` sẽ được tự động parse số tĩnh và đối chiếu giá trị thực tế của gói thầu để xác định hiển thị.
> 2. Nếu không có `Condition` hoặc trống: Fallback tự động lọc theo `Price` (`Price <= dblPrice <= PriceMax`) và `Type` (`ALL` hoặc trùng loại gói thầu).

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
| `<<GiaTri.Chu>>` | `{{GiaTri\|num2text}}` | Chuyển đổi số sang dạng chữ tiếng Việt | `Một triệu năm trăm nghìn` |
| | `{{Ngay\|day}}` | Chỉ lấy phần ngày (2 chữ số) | `30` |
| | `{{Ngay\|month}}` | Chỉ lấy phần tháng (2 chữ số) | `07` |
| | `{{Ngay\|year}}` | Chỉ lấy phần năm (4 chữ số) | `2026` |
| | `{{Ngay\|add_days(5)}}` | Cộng thêm N ngày | `04/08/2026` |
| | `{{Ngay\|add_months(2)}}` | Cộng thêm N tháng | `30/09/2026` |
| | `{{Ngay1\|date_diff(Ngay2)}}` | Tính chênh lệch ngày giữa hai mốc thời gian | `15` |
| | `{{Ngay\|quarter}}` | Lấy ký hiệu Quý trong năm (tiếng Việt) | `Quý III/2026` |
| | `{{Ngay\|weekday}}` | Lấy tên thứ trong tuần (tiếng Việt) | `Thứ Năm` |
| | `{{Ngay\|date_text}}` | Đọc toàn bộ ngày thành chữ tiếng Việt | `Ngày ba mươi tháng bảy năm hai nghìn không trăm hai mươi sáu` |
| | `{{now}}` | Biến context toàn cục lấy ngày giờ hiện tại | *(datetime object)* |

**Implementation:**
```python
from docxtpl import DocxTemplate
import jinja2

# Các hàm filter được đăng ký:
# - filter_date, filter_date_long, filter_number, filter_num2text
# - filter_day, filter_month, filter_year, filter_add_days
# - filter_add_months, filter_date_diff, filter_quarter, filter_weekday, filter_date_text

tpl = DocxTemplate("template.docx")
jenv = jinja2.Environment()
jenv.filters["date"]       = filter_date
jenv.filters["date_long"]  = filter_date_long
jenv.filters["number"]     = filter_number
jenv.filters["num2text"]   = filter_num2text
jenv.filters["day"]        = filter_day
jenv.filters["month"]      = filter_month
jenv.filters["year"]       = filter_year
jenv.filters["add_days"]   = filter_add_days
jenv.filters["add_months"] = filter_add_months
jenv.filters["date_diff"]  = filter_date_diff
jenv.filters["quarter"]    = filter_quarter
jenv.filters["weekday"]    = filter_weekday
jenv.filters["date_text"]  = filter_date_text

# Tự động chèn biến "now" vào context
context["now"] = datetime.now()

tpl.render(context, jenv)
tpl.save("output.docx")
```

**Xử lý lỗi & Định dạng đặc thù:**
- `|date` hoặc `|date_long` mà không parse được → `⚠️ Warning` + giữ nguyên giá trị gốc
- `|number` mà không phải số → `❌ Lỗi` + dừng file đó
- `|date` hoặc `|date_long` mà rỗng → lấy ngày hiện tại
- **Xử lý ngày tháng trống một phần (Chừa khoảng trống ghi tay)**:
  - Đối với các bộ lọc ngày tháng (như `|date_long`), nếu chuỗi ngày tháng chứa dấu gạch chéo `/` nhưng có phần ngày hoặc tháng bị bỏ trống (chỉ chứa khoảng trắng):
    - Nếu ngày trống (VD: `"   /07/2026"`): Tự động định dạng thành `"ngày   tháng 07 năm 2026"` (mặc định để 3 khoảng trắng trước ngày).
    - Nếu cả ngày và tháng đều trống (VD: `"  /   /2026"`): Tự động định dạng thành `"ngày   tháng   năm 2026"` (mặc định để 3 khoảng trắng trước cả ngày và tháng).

---

## 5. Quy trình chạy

> **Luồng duy nhất:** 1 lần chạy = 1 gói thầu (chọn từ danh sách) → chọn template → xử lý.

**Luồng chi tiết:**
```
[1] Load DataSet: duckdb excel extension quét tất cả .xlsx trong 1. Data/
    → Mỗi sheet (trừ S.*) thành 1 DuckDB table:
      GoiThau, Options, Workflow, Tables, Config
    → Tạo bản sao lưu DuckDB cho tất cả các bảng với hậu tố "_Goc" (Ví dụ: TCGTTD_Goc) để giữ nguyên dữ liệu gốc khi thực hiện đăng ký bảng tạm lặp.

[2] UI – Chọn Option (Opt1 / Opt2):
    • Hiển thị radio button từ DuckDB table Options

[3] UI – Chọn Gói thầu:
    • Query: SELECT * FROM {LeftSheet} (với LeftSheet lấy từ cấu hình Sheet của Option)
    • Hiển thị gói thầu theo định dạng Show (phần trước dấu |)

[4] UI – Chọn đối tượng lặp / template file:
    • Nếu Option thuộc loại "Repeat":
        - Lấy danh sách thành viên/đối tượng từ bảng con {RightSheet}_Goc WHERE {JoinKey} = {goi_thau_id}.
        - Hiển thị danh sách checkbox thành viên theo định dạng Show (phần sau dấu |).
    • Nếu Option thông thường:
        - Query Workflow lọc template theo Condition và Price.
        - Hiển thị checkbox list template file.

[5] Xác nhận chạy
        ↓
Xóa toàn bộ 3. Files/ (hoặc giữ lại nếu đang chạy chế độ Retry sửa lỗi)
        ↓
Trong chế độ Repeat (Lặp đối tượng):
  • Đối với từng thành viên được tick chọn (tuần tự):
    - Đăng ký dòng dữ liệu của thành viên đó vào bảng tạm {RightSheet} trong DuckDB để thực hiện JOIN query.
    - Lấy template khớp với Workflow.
    - Build context dict (tự động map cột sang key qua Config, trích xuất tên Họ tên động).
    - Render và ghi đè bảng từ Excel (Tables) thông qua key_id và danh_muc_file.
    - Save file output cho từng đối tượng lặp.
Trong chế độ thông thường:
  • Với từng file Word trong 3. Files/ (tuần tự):
    - docxtpl scan template → lấy danh sách biến {{ }} cần thiết
    - Query DuckDB table Config để map key → tên cột
    - Lấy giá trị từ row đã chọn → build context dict
    - Xử lý table placeholder từ Tables.
    - Save file (đổi tên: bỏ "-Template").
        ↓
Báo cáo kết quả (hiển thị trạng thái log chi tiết, nút mở thư mục log/output)

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

---

## 14. PRD – KisorDoc Enhancements (Bổ sung nâng cao)

### Tổng quan

6 tính năng bổ sung nâng cao trải nghiệm sử dụng KisorDoc sau khi core pipeline ổn định. Các tính năng được thiết kế độc lập — có thể implement theo bất kỳ thứ tự nào mà không ảnh hưởng nhau.

| # | Tính năng | Ưu tiên | Độ phức tạp |
|---|---|---|---|
| F1 | Validation trước khi chạy | 🔴 Cao | Thấp |
| F2 | Dry-run / Preview mode | 🔴 Cao | Trung bình |
| F3 | Retry cho file lỗi | 🟡 Trung bình | Thấp |
| F4 | Export log ra file | 🟡 Trung bình | Thấp |
| F5 | Version pin Config/Tables | 🟢 Thấp | Trung bình |
| F6 | Xử lý file đang mở (locked) | 🔴 Cao | Thấp |

---

### F1 – Validation trước khi chạy

#### Vấn đề
Hiện tại nếu người dùng nhấn "Chạy" mà chưa chọn đủ thông tin (Option, Gói thầu, template), `run_batch()` nhận input rỗng và chạy với behavior không xác định — có thể crash, tạo file sai tên, hoặc xóa `3. Files/` mà không tạo file mới.

#### Mục tiêu
Chặn sớm và thông báo rõ ràng trước khi bất kỳ thao tác file nào xảy ra.

#### Yêu cầu

**F1-01:** Khi nhấn "Chạy", kiểm tra tuần tự theo thứ tự:

| Bước | Điều kiện | Thông báo lỗi |
|---|---|---|
| 1 | Option đã được chọn | "Vui lòng chọn quy trình trước" |
| 2 | Gói thầu đã được chọn | "Vui lòng chọn gói thầu trước" |
| 3 | Ít nhất 1 template được check | "Vui lòng chọn ít nhất 1 file template" |
| 4 | Thư mục `2. Templates/{Option}/` tồn tại | "Không tìm thấy thư mục Templates/{Option}" |
| 5 | Tất cả file template được chọn tồn tại trên disk | "Không tìm thấy file: {tên file}" |
| 6 | DataSet đã load thành công (ds không phải None) | "Dữ liệu chưa được tải. Khởi động lại app." |

**F1-02:** Nếu bất kỳ bước nào fail → dừng ngay, hiển thị thông báo lỗi ở `status_text`, không xóa `3. Files/`, không copy file nào.

**F1-03:** Validation chạy đồng bộ (không phải async) để kết quả hiển thị ngay lập tức trước khi progress bar xuất hiện.

**F1-04:** Thông báo lỗi hiển thị ở `status_text` với prefix `⚠️` — không dùng popup/modal (Gradio không hỗ trợ tốt).

#### Triển khai

```python
def validate_inputs(option_key, package_label, selected_templates) -> tuple[bool, str]:
    """Trả về (is_valid, error_message). is_valid=True nếu tất cả điều kiện thỏa."""
    if not option_key:
        return False, "⚠️ Vui lòng chọn quy trình trước"
    if not package_label:
        return False, "⚠️ Vui lòng chọn gói thầu trước"
    if not selected_templates:
        return False, "⚠️ Vui lòng chọn ít nhất 1 file template"
    if ds is None:
        return False, "⚠️ Dữ liệu chưa được tải. Vui lòng khởi động lại app."
    # Kiểm tra file tồn tại
    opt = option_key.split(":")[0].strip()
    template_dir = config.template_path / opt
    if not template_dir.exists():
        return False, f"⚠️ Không tìm thấy thư mục: {template_dir}"
    return True, ""

# Trong run_batch():
is_valid, err_msg = validate_inputs(option_key, package_label, selected_templates)
if not is_valid:
    yield [], err_msg
    return
```

---

### F2 – Dry-run / Preview mode

#### Vấn đề
Không có cách nào biết liệu data trong Excel có đủ cho tất cả placeholder trong template không, cho đến khi chạy thật và nhận được lỗi hoặc file output trống. Với gói thầu có 10+ file template, việc discover lỗi config sau khi chạy rất tốn thời gian.

#### Mục tiêu
Cho phép người dùng "xem trước" kết quả merge mà không tạo file thật — chỉ cần biết placeholder nào được điền, placeholder nào thiếu data.

#### Yêu cầu

**F2-01:** Thêm nút "🔍 Kiểm tra" (bên cạnh nút "🚀 Chạy") trong Tab 2.

**F2-02:** Khi nhấn "Kiểm tra", thực hiện validation F1 trước. Nếu pass → tiến hành dry-run.

**F2-03:** Dry-run thực hiện cho từng template được chọn:
- Scan tất cả placeholder `{{key}}` trong file Word (dùng `docxtpl.get_undeclared_template_variables()`)
- Build context dict như bình thường (query DuckDB)
- So sánh: key nào có trong context, key nào thiếu
- Với table placeholder: kiểm tra có dòng tương ứng trong sheet `Tables` không, range có hợp lệ không
- **Không** tạo file output, **không** xóa `3. Files/`

**F2-04:** Kết quả dry-run hiển thị dạng bảng trong Tab 3:

| File template | Placeholder | Trạng thái | Giá trị (preview) |
|---|---|---|---|
| 3. Yeu cau bao gia | `{{HoTen}}` | ✅ Có data | Nguyễn Văn A |
| 3. Yeu cau bao gia | `{{NgayKy\|date}}` | ✅ Có data | 01/07/2026 |
| 3. Yeu cau bao gia | `{{CDT_DiaChi}}` | ❌ Thiếu data | – |
| 3. Yeu cau bao gia | `{{DanhMuc}}` | ✅ Table OK | Sheet S.DanhMuc, A1:F18 |
| 14. Hop dong | `{{SoHD}}` | ⚠️ Rỗng | "" |

**F2-05:** Tổng kết sau dry-run:
```
✅ 3 file sẵn sàng chạy
⚠️ 1 file có warning (giá trị rỗng)
❌ 1 file thiếu data (sẽ lỗi nếu chạy thật)
```

**F2-06:** Sau khi xem kết quả dry-run, người dùng có thể nhấn "🚀 Chạy" ngay — không cần chọn lại từ đầu.

**F2-07:** Preview value bị truncate nếu dài hơn 50 ký tự — hiển thị `"Nguyễn Văn A..."`.

#### UI

```
Tab 2:
[🔍 Kiểm tra]  [🚀 Chạy]   ← 2 nút nằm cạnh nhau

Tab 3 (sau dry-run):
  Chế độ: Kiểm tra (không tạo file)
  ─────────────────────────────────
  [Bảng kết quả per-placeholder]
  ─────────────────────────────────
  ✅ 3 sẵn sàng  ⚠️ 1 warning  ❌ 1 thiếu data
```

---

### F3 – Retry cho file lỗi ✅ (ver1.10)

#### Vấn đề
Nếu 1 trong 10 file bị lỗi (file Word bị lock, template bị corrupt, placeholder sai tên...), phải chạy lại toàn bộ batch — mất thời gian, xóa lại `3. Files/`, overwrite các file đã OK.

#### Mục tiêu
Chạy lại chỉ các file bị lỗi mà không ảnh hưởng đến các file đã thành công.

#### Yêu cầu

**F3-01:** Sau khi batch kết thúc, nếu có ít nhất 1 file ❌ → hiển thị nút "🔄 Chạy lại file lỗi".

**F3-02:** Nút này chỉ hiện khi có lỗi; ẩn khi tất cả file thành công.

**F3-03:** Khi nhấn "Chạy lại file lỗi":
- Lấy danh sách template tương ứng với các dòng ❌ trong log
- **Không** xóa `3. Files/` (giữ nguyên các file đã OK)
- **Không** copy lại file đã OK từ Templates
- Chỉ copy và xử lý lại các template bị lỗi
- Context (Option + Gói thầu) giữ nguyên từ lần chạy trước

**F3-04:** Sau retry, cập nhật log: dòng lỗi cũ được thay bằng kết quả mới (✅ hoặc ❌ với lý do mới).

**F3-05:** Có thể retry nhiều lần liên tiếp.

**F3-06:** Lưu trạng thái retry trong `gr.State` — reset về None khi người dùng nhấn "Chọn lại từ đầu" hoặc thay đổi Option/Gói thầu.

#### State cần thêm

```python
last_run_state = gr.State(None)
# last_run_state = {
#   "option_key": "Opt1: ...",
#   "package_label": "1. MS26-01 ...",
#   "all_results": [{"template": "...", "status": "❌", "error": "..."}],
#   "failed_templates": ["3. Yeu cau bao gia", "14. Hop dong"]
# }
```

---

### F4 – Export log ra file

#### Vấn đề
Log hiện tại chỉ hiển thị trên UI Gradio. Khi đóng app hoặc chạy lại → log mất hoàn toàn. Không có cách nào tra cứu "lần trước tôi đã xử lý những file gì, có lỗi gì không".

#### Mục tiêu
Tự động ghi log mỗi lần chạy ra file text, lưu tại thư mục dự án.

#### Yêu cầu

**F4-01:** Tạo thư mục `logs/` trong `{ProjectPath}/` nếu chưa có.

**F4-02:** Mỗi lần chạy (kể cả dry-run và retry) tạo 1 file log mới:
```
logs/
├── 2026-07-29_143022_MS26-01_Opt1.log
├── 2026-07-29_150011_MS26-02_Opt1.log   
└── 2026-07-29_161533_MS26-01_Opt1_retry.log   ← suffix "_retry" nếu là retry
```
Pattern tên file: `{YYYY-MM-DD}_{HHmmss}_{GoiThau_ID}_{Option}.log`

**F4-03:** Nội dung file log:
```
=====================================
KisorDoc – Run Log
=====================================
Thời gian     : 2026-07-29 14:30:22
Option        : Opt1 – Các giấy tờ đến bước Hợp đồng
Gói thầu      : MS26-01 – XLNT Bệnh viện Sản Nhi
Chế độ        : Chạy thật / Dry-run / Retry
Tổng file     : 5
=====================================

[14:30:22] ✅ 0. Danh muc.A-MS26-01.docx
[14:30:24] ✅ 3. Yeu cau bao gia-MS26-01.docx
[14:30:25] ❌ 5. QD phe duyet-MS26-01.docx
           Lỗi: PermissionError – File đang được mở bởi ứng dụng khác
[14:30:27] ⚠️ 14. Hop dong-MS26-01.docx
           Warning: Placeholder {{CDT_DiaChi}} không có data

=====================================
Kết quả: 3 thành công / 1 lỗi / 1 warning
Thời gian chạy: 6.2 giây
=====================================
```

**F4-04:** Log được ghi **incremental** trong khi chạy (không đợi đến cuối) — nếu app crash giữa chừng vẫn có log.

**F4-05:** Tự động xóa log cũ hơn 30 ngày khi app khởi động (giữ tối đa 100 file log).

**F4-06:** Thêm nút "📋 Mở thư mục log" trong Tab 3 (cạnh nút "📂 Mở thư mục output").

**F4-07:** Log encoding: UTF-8 với BOM (`utf-8-sig`) để mở đúng trong Notepad Windows.

#### Triển khai

```python
import logging
from pathlib import Path

def setup_run_logger(config, goi_thau_id, option, mode="run") -> tuple[logging.Logger, Path]:
    log_dir = config.project_path / "logs"
    log_dir.mkdir(exist_ok=True)
    
    ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    suffix = f"_{mode}" if mode != "run" else ""
    log_file = log_dir / f"{ts}_{goi_thau_id}_{option}{suffix}.log"
    
    logger = logging.getLogger(f"kisordoc_{ts}")
    handler = logging.FileHandler(log_file, encoding="utf-8-sig")
    handler.setFormatter(logging.Formatter("[%(asctime)s] %(message)s", "%H:%M:%S"))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    
    return logger, log_file
```

---

### F5 – Version pin cho Config/Tables sheet

#### Vấn đề
Người dùng có thể vô tình sửa sheet `Config` hoặc `Tables` trong Excel giữa các lần chạy (thêm/xóa/đổi tên cột, thay đổi mapping). Không có cách nào biết config đang dùng là version nào, hay đã bị thay đổi gì so với lần chạy trước.

#### Mục tiêu
Snapshot config mỗi lần chạy để có thể đối chiếu khi có vấn đề.

#### Yêu cầu

**F5-01:** Khi bắt đầu mỗi lần chạy (sau validation, trước khi xử lý file), snapshot nội dung 2 sheet quan trọng:
- Sheet `Config` → JSON object `{key: col_name}`
- Sheet `Tables` → list of dicts `[{GoiThau_ID, Word, Name, Sheet, Range, Hide}]`

**F5-02:** Lưu snapshot cùng file log (thêm vào cuối file `.log`):
```
=====================================
CONFIG SNAPSHOT
=====================================
{"KHLCNT_TTr": "KHLCNT_TTr", "HoTen_NguoiLap": "HoTen_NguoiLap", ...}

TABLES SNAPSHOT (filter GoiThau_ID = MS26-01)
[{"Word": "Danh muc", "Name": "DanhMuc", "Sheet": "S.DanhMuc", "Range": "A1:F", "Hide": ""}]
=====================================
```

**F5-03:** Khi app khởi động, so sánh config hiện tại với snapshot của lần chạy gần nhất cho cùng gói thầu:
- Nếu có thay đổi → hiển thị banner `⚠️ Config đã thay đổi so với lần chạy trước` kèm diff đơn giản
- Nếu không thay đổi → không hiển thị gì

**F5-04:** Diff hiển thị dạng đơn giản:
```
⚠️ Config thay đổi so với lần chạy trước (2026-07-28 15:30):
  + Thêm mới: CDT_DiaChiDayDu → Dia_Chi_CDT_Full
  - Xóa: CDT_DiaChi
  ~ Đổi: HoTen_NguoiLap: "Ho_Ten" → "Ho_Ten_NLP"
```

**F5-05:** Diff chỉ hiển thị 1 lần khi app load, không lặp lại trong session.

**F5-06:** Nút "Bỏ qua cảnh báo" để dismiss banner.

#### Lưu ý
Tính năng này hoàn toàn informational — không block việc chạy, chỉ cảnh báo.

---

### F6 – Xử lý file đang mở (File Locked) ✅ (ver1.10)

#### Vấn đề
Khi KisorDoc ghi file output vào `3. Files/` mà file đó đang mở trong Microsoft Word (hoặc đang được process khác giữ), Python raise `PermissionError`. Hiện tại lỗi này bị catch chung cùng tất cả exception khác → người dùng thấy thông báo lỗi khó hiểu và không biết cách xử lý.

#### Mục tiêu
Phát hiện riêng lỗi file locked, thông báo rõ ràng, và cung cấp cơ chế retry tự động.

#### Yêu cầu

**F6-01:** Phân biệt `PermissionError` khỏi các exception khác trong `run_batch()`:

```python
except PermissionError as e:
    if "being used by another process" in str(e) or e.errno == 13:
        # File locked
        results.append(f"🔒 {tpl_name}: File đang mở – vui lòng đóng lại và thử lại")
    else:
        results.append(f"❌ {tpl_name}: Lỗi quyền truy cập – {e}")
```

**F6-02:** Khi phát hiện file locked → tự động retry tối đa **3 lần**, mỗi lần chờ **2 giây**:
```
Lần 1 → PermissionError → chờ 2s → thử lại
Lần 2 → PermissionError → chờ 2s → thử lại  
Lần 3 → PermissionError → dừng, báo lỗi 🔒
```

**F6-03:** Trong thời gian chờ retry, hiển thị trạng thái:
```
🔒 14. Hop dong: File đang bị khóa – thử lại lần 1/3 sau 2 giây...
```

**F6-04:** Thông báo lỗi cuối cùng (sau 3 lần thất bại) phải dễ hiểu:
```
🔒 14. Hop dong-MS26-01.docx: Không thể ghi file vì đang được mở trong Word.
   → Vui lòng đóng file này và nhấn "🔄 Chạy lại file lỗi"
```

**F6-05:** File locked được đánh dấu là `🔒` (khác `❌`) trong log để người dùng biết nguyên nhân cụ thể và cách xử lý.

**F6-06:** Nếu tất cả file lỗi đều là `🔒` (không có `❌`), nút "Chạy lại file lỗi" đổi label thành "🔄 Chạy lại (đã đóng file chưa?)".

**F6-07:** Cũng áp dụng logic retry cho bước **đọc** file template (khi `copy_templates_to_output` fail do template đang mở trong Word).

#### Hàm tiện ích

```python
import time, errno

def write_with_retry(func, max_retries=3, delay=2.0, yield_fn=None):
    """
    Wrapper retry cho bất kỳ thao tác file nào.
    yield_fn: callable(msg) để update UI trong khi chờ.
    """
    for attempt in range(1, max_retries + 1):
        try:
            return func()
        except PermissionError as e:
            if attempt == max_retries:
                raise
            msg = f"🔒 File bị khóa – thử lại lần {attempt}/{max_retries} sau {delay:.0f}s..."
            if yield_fn:
                yield_fn(msg)
            time.sleep(delay)
```

---

### Thứ tự implement

```
F1 (Validation)     ← Làm trước, dễ, chặn nhiều lỗi ngay
    ↓
F6 (File locked)    ← Làm sớm, hay gặp trong thực tế
    ↓
F4 (Export log)     ← Nền tảng cho F5, làm trước F5
    ↓
F3 (Retry)          ← Phụ thuộc vào log state từ F4
    ↓
F2 (Dry-run)        ← Độc lập, làm sau khi pipeline ổn
    ↓
F5 (Version pin)    ← Nice-to-have, làm cuối
```

---

## 11. Các cập nhật lớn trong Phiên bản 3.0 (2026-07-31)

### 11.1 Dịch chuyển định dạng Jinja2 & Modifier Mới
*   Nâng cấp toàn bộ các placeholder biến dạng cũ `<<TenBien>>` và `<<TenBien.Modifier>>` trong các template Word (`Opt1` $\rightarrow$ `Opt5`) sang dạng đóng mở ngoặc kép **Jinja2** `{{ }}`, kết hợp sử dụng bộ lọc `|` chuẩn:
    *   `<<TenBien.Date.Long>>` $\rightarrow$ `{{ TenBien_Date|date_long }}`
    *   `<<TenBien.Date>>` $\rightarrow$ `{{ TenBien_Date|date }}`
    *   `<<TenBien.Upper>>` $\rightarrow$ `{{ TenBien|upper }}`
    *   `<<TenBien.Number>>` $\rightarrow$ `{{ TenBien|number }}`
    *   `<<TenBien.Num2Text>>` $\rightarrow$ `{{ TenBien|num2text }}`
*   Ánh xạ thông minh: Toàn bộ các hậu tố dấu chấm `.Date` và `.Date.Long` trong file cấu hình Excel tự động được ánh xạ thành **`_Date`** để tránh xung đột ghi đè giữa Số quyết định (Ví dụ: `KH_QD`) và Ngày quyết định (`KH_QD_Date`).

### 11.2 Phân vùng Config theo Option (`config_range`)
*   Bổ sung cột `Config` (hoặc `config_range`) trong sheet `Options` của Excel (Định dạng: `StartRow-EndRow`, ví dụ `2-97`). 
*   Cho phép giới hạn phạm vi nạp cấu hình mapping cho từng Option riêng biệt, tránh xung đột placeholder chéo giữa các quy trình khác nhau.

### 11.3 Copy bảng biểu từ file Excel động (`File` column)
*   Sheet `Tables` hỗ trợ thêm cột `File` để chỉ định cụ thể file Excel nguồn chứa bảng dữ liệu (Ví dụ: `S.Oto.xlsx`).
*   Nếu cột `File` không có hoặc để trống (chuẩn cũ) $\rightarrow$ Tự động fallback về file danh mục mặc định của gói thầu.
*   **Tự động nhân bản bảng:** Nếu trong Word có $N$ thẻ placeholder giống nhau nhưng Excel chỉ khai báo ít hơn $N$ dòng cấu hình, hệ thống sẽ tự động sao chép bảng dữ liệu cuối cùng khả dụng để điền vào tất cả các vị trí thẻ còn lại, không để sót thẻ thô.

### 11.4 Liên kết bảng thông minh (Join Sheets)
*   Cột `Sheet` trong sheet `Options` của Excel hỗ trợ liên kết các bảng dữ liệu:
    *   **Join 2 bảng (Sử dụng ký hiệu rút gọn):** `[Bảng 1] [Ký hiệu] [Bảng 2] @ [Khóa]`.
        *   Các ký hiệu: `<*` (Left Join), `*>` (Right Join), `*` (Inner Join), `<*>` (Full Outer Join).
        *   Ví dụ: `GoiThau <* TCGTTD @ GoiThau_ID`
    *   **Join 3 bảng trở lên (Sử dụng SQL trực tiếp):** Nếu biểu thức bắt đầu bằng `SELECT`, hệ thống sẽ cho chạy trực tiếp câu lệnh truy vấn SQL này trên DuckDB.
*   **Cảnh báo trùng tên cột (Column Collision Warning):** Khi thực hiện join, nếu phát hiện 2 bảng cùng có các cột trùng tên nhau, hệ thống sẽ in cảnh báo chi tiết màu vàng lên Gradio UI/Log để người dùng nhận diện và điều chỉnh lại trong Excel, tránh ghi đè dữ liệu sai lệch.

### 11.5 Tối ưu hóa khởi động hệ thống
*   Nạp DuckDB thông qua thư viện trung gian giúp giảm thời gian khởi chạy ban đầu xuống dưới 1 giây.

---

## 12. Các cập nhật trong Phiên bản 3.1 (2026-08-03)

### 12.1 Đổi tên Core Package sang `kisorlib`
*   Đổi tên toàn bộ thư mục thư viện cốt lõi từ `kisordoc/` thành **`kisorlib/`**.
*   Đồng bộ toàn bộ các import tham chiếu tới `kisordoc` sang `kisorlib` trong tất cả file của dự án (`app.py`, `api.py`, `runner.py`, `tests/test_engine.py`).
*   Cách đặt tên mới giúp tăng tính chuyên nghiệp, định vị rõ ràng đây là thư viện dùng chung cho toàn bộ dự án.

### 12.2 Cải tiến hiển thị khoảng trắng ngày tháng trống
*   Khi sử dụng các modifier định dạng ngày tháng như `.Date.Long` (hoặc `|date_long`), nếu chuỗi ngày tháng chứa dấu `/` nhưng bị bỏ trống (chỉ có khoảng trắng) thì hệ thống sẽ định dạng với **3 khoảng trắng** cố định cho phần trống.
    *   Ví dụ ngày trống: `"   /07/2026"` $\rightarrow$ `"ngày   tháng 07 năm 2026"`
    *   Ví dụ cả ngày và tháng trống: `"  /   /2026"` $\rightarrow$ `"ngày   tháng   năm 2026"`
*   Cách hiển thị này tạo sự đồng đều thẩm mỹ trên văn bản và chừa khoảng trống vừa vặn để viết tay hoặc đóng dấu sau này.

*   Tự động bỏ qua các file Excel nguồn chứa bảng biểu bắt đầu bằng chữ `S.` lúc khởi động app. Các file này chỉ được mở đọc khi thực hiện tiến trình chèn bảng biểu vào Word, giúp ứng dụng khởi động tức thì và tiết kiệm 90% dung lượng RAM.
*   Hệ thống tự động thực hiện **gộp (concatenate)** dữ liệu từ nhiều file Excel khi phát hiện trùng tên sheet (Ví dụ: Sheet `Tables` có ở cả `Tables.xlsx` và `DanhMuc-MSSC.xlsx`).

---

## 13. Các cập nhật trong Phiên bản 4.0, 4.0.1, 4.0.2 & 4.0.3 (2026-08-05)

### 13.1 Tái cấu trúc mã nguồn app.py (Refactoring app.py)
*   Phân rã file `app.py` cồng kềnh dài ~1780 dòng thành các module nhỏ, đơn nhiệm và được đóng gói trong thư mục cốt lõi `kisorlib/`:
    *   `kisorlib/utils.py`: Đóng gói các hàm tiện ích thuần túy (pure functions) như xử lý chuỗi, định dạng, phân tích cú pháp join rút gọn.
    *   `kisorlib/service.py`: Chứa lớp `KisorService` chịu trách nhiệm xử lý logic nghiệp vụ, giao tiếp DataSet (DuckDB). Loại bỏ trạng thái toàn cục `config` và `ds` qua Dependency Injection.
    *   `kisorlib/batch.py`: Đóng gói logic sinh tài liệu hàng loạt `run_batch` và `run_retry_batch`. Module này hoàn toàn độc lập với Gradio UI bằng cách sử dụng callback `progress_cb: Callable` thay vì `gr.Progress`.
    *   `app.py`: Rút gọn xuống chỉ còn ~360 dòng, đóng vai trò là entry point cấu hình layout và kết nối event handler của giao diện Gradio UI.

### 13.2 Loại bỏ mã nguồn trùng lặp trong engine.py
*   Loại bỏ hoàn toàn các hàm clone trùng lặp trước đó (`_clean_config_key`, `_get_option_config_from_ds`) trong `kisorlib/engine.py`.
*   Chuyển sang import và tái sử dụng trực tiếp cấu hình/tiện ích từ `kisorlib.utils` và `kisorlib.service`.

### 13.3 Khắc phục lỗi hiển thị Preview trong chế độ lặp Repeat
*   Khắc phục lỗi composite `key_id` (khi cấu hình `key_id` chứa ký tự ghép `|`) trong hàm `run_preview`, đảm bảo phân tích chính xác khóa bảng chính (`left_key`) để tìm đúng thông tin bảng liên kết hiển thị lên Preview UI.

### 13.4 Bộ kiểm thử tự động (Unit Tests)
*   Bổ sung thư mục kiểm thử `tests/` chứa các bộ test case tự động:
    *   [tests/test_utils.py](tests/test_utils.py): Kiểm tra các hàm thuần túy như phân tích giá tiền, dọn dẹp key config, biểu thức join rút gọn, giải quyết khoảng dòng,...
    *   [tests/test_filters.py](tests/test_filters.py): Kiểm tra bộ lọc định dạng số `filter_number`.
*   Giúp chạy kiểm thử hồi quy tức thì và bảo vệ tính đúng đắn của logic tính toán khi phát triển ứng dụng.

### 13.5 Bản vá 4.0.1 (Bảo mật SQL & Bổ sung Unit Tests)
*   **Bảo mật SQL (Parameter Binding):** Tách biệt dữ liệu truyền vào khỏi câu lệnh SQL trong `kisorlib/service.py` bằng DuckDB parameter binding `?` để ngăn chặn hoàn toàn SQL Injection.
*   **Dọn dẹp hardcode còn lại trong api.py:** Cập nhật `api.py` sử dụng thuộc tính đường dẫn động `cfg.data_path` và tên bảng động `cfg.DataSheet`, đồng thời sửa lỗi khởi tạo `DataSet` truyền tham số `excel_files` lỗi thời.
*   **Kiểm thử dịch vụ [tests/test_service.py](tests/test_service.py):** Bổ sung unit tests cho `KisorService`, kiểm tra Repeat mode, đăng ký bảng thành viên tạm thời và preview composite key.

### 13.6 Bản vá 4.0.2 (Động hóa lựa chọn lặp & Sửa lỗi Radio State của Gradio)
*   **Tổng quát hóa bộ chọn lặp:** Đổi mặc định nhóm lặp trong [ui_labels.json](ui_labels.json) từ gán cứng `"Tổ chuyên gia" / "Tổ thẩm định"` thành `"Nhóm lặp 1" / "Nhóm lặp 2"`.
*   **Sửa lỗi đồng bộ Radio State của Gradio:** Khắc phục lỗi `gradio.exceptions.Error` khi khởi tạo/reset app về quy trình trống bằng cách khôi phục lại các lựa chọn mặc định và giá trị hợp lệ cho `group_radio`.
*   **Cập nhật tài liệu lỗi [known-issues.md](known-issues.md):** Bổ sung kết quả xử lý và phân tích rủi ro bảo mật tiềm ẩn SQL Identifier Injection.

### 13.7 Bản vá 4.0.3 (Whitelist định danh SQL - SQL Identifier Whitelist)
*   **Whitelist định danh SQL:** Tích hợp hàm `validate_sql_identifier` với regex whitelist `^[A-Za-z0-9_\s\-\.\#\u00C0-\u1EF9]+$` để lọc sạch toàn bộ tên bảng và cột động đọc từ Excel trước khi đưa vào SQL, triệt tiêu hoàn toàn nguy cơ SQL Identifier Injection.
*   **Mở rộng Unit Tests:** Thêm các test case trong [test_utils.py](tests/test_utils.py) để kiểm chứng bộ lọc an toàn với các trường hợp inject ký tự độc hại.
*   **Cập nhật Tài liệu Lỗi:** Đánh dấu lỗi SQL Identifier Injection đã được khắc phục hoàn toàn trong [known-issues.md](known-issues.md).

---

## 14. Các cập nhật trong Phiên bản 4.1.x & 5.0.0 (2026-08-05 đến 2026-08-07)

### 14.1 Phiên bản 4.1.0: Xử lý triệt để Dual-pipeline
*   **Module Core Sync `kisorlib/generator.py`:** Tích hợp logic sinh tài liệu đồng bộ làm Single Source of Truth cho cả UI (`batch.py`) và API (`engine.py`).
*   **Tách module `kisorlib/sql_join.py`:** Di chuyển các tiện ích xử lý SQL Join biểu thức từ `utils.py` sang module mới.
*   **Unit Tests Hồi Quy:** Bổ sung [tests/test_generator.py](tests/test_generator.py) với 25 test cases kiểm chứng chi tiết hoạt động của core `generator`.

### 14.2 Bản vá 4.1.1 & 4.1.2 & 4.1.3: Cải tiến an toàn & Sửa lỗi
*   **Sửa lỗi Windows copy:** Khắc phục lỗi `SameFileError` khi core engine cố gắng sao chép một tệp tin đè lên chính nó.
*   **Nạp template Repeat:** Chỉ truy vấn bảng gói thầu (`left_sheet`) khi lấy danh sách template của Quy trình Lặp, tránh lỗi `BinderException`.
*   **Fallback đuôi mở rộng Excel:** Tự động điền đuôi `.xlsx` nếu tên file trong cột `File` của bảng cấu hình bị thiếu.
*   **Đồng bộ hóa thống kê log:** Tăng cả `ok_count` lẫn `warning_count` khi gặp file cảnh báo nhưng sinh thành công.
*   **Dynamic AppConfig:** Tự động load `FileMaxRetries` và `FileRetryDelay` từ cấu hình cấu trúc thay vì gán cứng.

### 14.3 Phiên bản 5.0.0: Tích hợp Bộ kiểm thử Toàn diện & Search UI
*   **Thêm Thanh tìm kiếm và nút Dừng trên Gradio UI:** 
    - Thêm chức năng tìm kiếm trực tiếp các template / đối tượng lặp.
    - Hỗ trợ nút Stop để dừng tiến trình tạo tài liệu hàng loạt ngay lập tức.
*   **Bộ 147 Unit Tests Mới:** Tích hợp bộ kiểm thử đầy đủ nâng tổng số test case lên 198, bao gồm:
    - [tests/test_sql_join.py](tests/test_sql_join.py) (+76 tests)
    - [tests/test_batch.py](tests/test_batch.py) (+48 tests)
    - [tests/test_merger.py](tests/test_merger.py) (+23 tests)
*   **Bảo mật SQL UI:** Chuyển đổi gọi query lấy tên nhóm lặp trong `app.py` sang dạng Parameterized Query `?`.
*   **Đồng bộ hóa bảng lặp tạm thời:** Áp dụng `_safe_table_name` cho các truy vấn bảng lặp `_Goc` trong `kisorlib/service.py` để tương thích với tên sheet tiếng Việt có dấu/khoảng trắng.

---

## 15. Các cập nhật trong Phiên bản 5.1.x & 5.2.x (2026-08-10 đến 2026-08-12)

### 15.1 Phiên bản 5.1.0: Core Migrator Mới
*   **Thư viện Core Migrator (`kisorlib/migrator.py`)**: Gộp toàn bộ logic của các script cũ vào một thư viện duy nhất để hỗ trợ API đồng bộ: `migrate_xml`, `migrate_file`, `migrate_folder`.
*   **Unit Tests đầy đủ**: Bổ sung bộ unit test toàn diện cho migrator tại `tests/test_migrator.py` đảm bảo các quy tắc chuyển đổi `<<Biến>>` $\rightarrow$ `{{Biến}}` chính xác 100%.

### 15.2 Phiên bản 5.2.0, 5.2.1 & 5.2.2: WHERE SQL & Single Instance & Tự động ẩn cột phụ & Chuyển tab tự động
*   **Tự động chuyển tab sau khi chạy (v5.2.2)**: Cấu hình giao diện tự động chuyển từ tab "1. Chọn & Chạy" sang tab "2. Log & Kết quả" ngay sau khi hoàn tất tiến trình sinh file hoặc chạy lại file lỗi.
*   **Hỗ trợ WHERE trong cú pháp SQL Join**: Bổ sung cơ chế phân tích cú pháp join rút gọn để tách và chuyển đổi mệnh đề `WHERE` bổ sung (ví dụ: `GoiThau * TCGTTD @ GoiThau_ID WHERE GoiThau.GoiThau_HTDT == 'DTRR'`).
*   **Khắc phục lọc quy trình Repeat**: Đồng nhất sử dụng câu lệnh SQL đầy đủ thay vì chỉ trích xuất `left_sheet` cho việc lấy danh sách gói thầu chính của quy trình Repeat, giúp áp dụng chính xác các điều kiện `WHERE` trên giao diện UI.
*   **Mở thư mục dạng Single Instance**: Cải tiến cơ chế mở thư mục output và log sang dạng Single Instance qua PowerShell COM Object (`Shell.Application`), tự động khôi phục và đưa cửa sổ Explorer hiện tại lên foreground thay vì mở thêm nhiều cửa sổ mới. Sửa lỗi đường dẫn chứa khoảng trắng mở nhầm thư mục *Documents* bằng lệnh `Invoke-Item`.
*   **Tự động ẩn cột phụ khi copy bảng**: Tích hợp thuật toán tự động lọc bỏ các cột nháp, cột phụ (bắt đầu bằng ký tự `_`, khớp với các từ khóa `phụ`/`helper`/`temp`/`nháp`/`draft` hoặc chứa nhãn dạng `[phụ]`/`(phụ)`...) khi trích xuất copy bảng từ Excel sang Word.
*   **Đăng ký bộ lọc Jinja2 tương thích**: Đăng ký và map đầy đủ các custom filters (`filter_upper` / `upper`, `filter_num2text` / `num2text`) trong `merger.py` để tránh lỗi `UndefinedError` trên template.
