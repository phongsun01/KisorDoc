**KisorDoc**

**Product Requirements Document**

Chức năng: Migrate Text to Jinja2 Placeholder

| **Phiên bản**  | 1.3                        |
| -------------- | -------------------------- |
| **Trạng thái** | Approved                   |
| **Ngày tạo**   | 31/7/2026                  |
| **Hệ thống**   | KisorDoc - Word Automation |
| **Module**     | Template Migration         |

# 1\. Tổng quan

Chức năng Migrate Text to Placeholder là bước chuẩn bị bắt buộc trong pipeline KisorDoc, cho phép chuyển đổi các file .docx "mẫu thô" (chứa giá trị cụ thể như tên người, ngày tháng, số tiền) thành file template hợp lệ của Jinja2 (chứa placeholder dạng {{ TenBien }}), sẵn sàng để hệ thống render hàng loạt.

## 1.1 Bối cảnh & vấn đề

Khi triển khai KisorDoc trên tài liệu thực tế, đội dự án nhận được file .docx gốc đã được điền sẵn dữ liệu mẫu. Việc thay thế thủ công từng giá trị trong hàng chục file mất nhiều giờ và dễ sót. Ngoài ra, Word nội bộ xé nhỏ text thành nhiều XML run khác nhau, khiến find-replace thông thường không đáng tin cậy.

## 1.2 Mục tiêu

- Tự động hóa việc chuyển đổi file .docx thô → template Jinja2.
- Nguồn mapping lấy từ file Excel hiện có - không cần tạo thêm file cấu hình riêng.
- Cho phép preview toàn bộ thay đổi trước khi commit (dry-run).
- Đảm bảo an toàn: backup tự động, không mất dữ liệu gốc.

## 1.3 Phạm vi

**Trong phạm vi (In scope):**

- File .docx trong thư mục chỉ định (duyệt đệ quy), hỗ trợ lọc qua `--include` / `--exclude`.
- Xử lý XML trên `document.xml`, `header*.xml`, `footer*.xml` (Must), `footnotes.xml`, `endnotes.xml` (Should), và các thẻ text đặc biệt như textbox, SDT (Could - best-effort, không đánh fail cả file nếu không parse được).
- Mapping từ Excel: header row = tên biến, row mẫu = giá trị cụ thể cần thay (khớp chính xác tuyệt đối sau khi `strip`).
- Xuất report HTML + Excel khi chạy dry-run (và tùy chọn khi migrate thật).
- Backup tự động với hậu tố timestamp trước khi ghi đè.

**Ngoài phạm vi (Out of scope):**

- File .doc legacy (cần convert sang .docx trước).
- Jinja2 Filters (ví dụ: `{{ NgayKy|date }}` hay `|number`). Phase 1 chỉ thay thế `{{ TenBien }}` thuần túy; việc map filter sẽ do bước khác hoặc làm thủ công.
- Placeholder kiểu khác ngoài {{ TenBien }} (Jinja2 block/comment tag).
- Xử lý text trong ảnh, header/footer có watermark ảnh.
- Giao diện GUI - chức năng này là CLI/script.

## 1.4 Tích hợp KisorDoc (Lưu ý quan trọng)

- **Nguồn sự thật từ bảng Config:** Bảng Config của KisorDoc có cấu trúc `Key` = Tên placeholder, `Value` = Tên cột Excel.
- **Luồng map chuẩn xác:**
  - Header Excel → Tên cột.
  - Row mẫu → Giá trị thực tế tại cột đó.
  - Tìm trong bảng Config dòng có `Value` khớp với Tên cột (sau khi normalize: `strip()` khoảng trắng và chuyển về `lower` để so sánh case-insensitive) → Lấy `Key` → Placeholder `{{ Key }}`.
  - Mapping replace: Giá trị thực tế → `{{ Key }}`.
  - Key lấy từ Config phải được chuẩn hóa qua hàm `kisorlib.utils.clean_config_key` để đảm bảo template khớp chuẩn mail merge của hệ thống.
- **Cơ chế Fallback:** Nếu không tìm thấy cột trong Config, fallback tạo placeholder từ Header của Excel theo quy tắc slug bỏ dấu.

# 2\. User Stories

| **#** | **Vai trò**    | **Tôi muốn...**                                        | **Để...**                                       | **Độ ưu tiên** |
| ----- | -------------- | ------------------------------------------------------ | ----------------------------------------------- | -------------- |
| US-01 | Dev triển khai | Chạy script với --dry-run để xem trước tất cả thay đổi | Kiểm tra mapping trước khi sửa file thật        | Must           |
| US-02 | Dev triển khai | Script tự đọc mapping từ Excel (header + row mẫu)      | Không phải tạo file config riêng                | Must           |
| US-03 | Dev triển khai | Chỉ định row mẫu khi chạy (--row N)                    | Linh hoạt dùng bất kỳ row nào làm mẫu           | Must           |
| US-04 | Dev triển khai | Nhận report HTML mở trên browser ngay                  | Review nhanh, highlight rõ trước/sau            | Must           |
| US-05 | PM / QA        | Xuất report dạng Excel để filter/sort                  | Kiểm tra từng biến, phát hiện mapping sai       | Must           |
| US-06 | Dev triển khai | File gốc được backup .bak.docx tự động                 | Rollback nếu kết quả sai                        | Must           |
| US-07 | Dev triển khai | Script xử lý đúng dù Word xé run XML                   | Không sót placeholder do lỗi XML nội bộ         | Must           |
| US-08 | Dev triển khai | Placeholder dài được ưu tiên thay trước                | Tránh sai khi tên ngắn là substring của tên dài | Must           |

# 3\. Yêu cầu chức năng

## 3.1 Input

| **Tham số**                | **Bắt buộc** | **Mô tả**                                                                    |
| -------------------------- | ------------ | ---------------------------------------------------------------------------- |
| \--excel &lt;path&gt;      | Có           | File Excel chứa dữ liệu. Row 1 = header, row N = giá trị mẫu.                |
| \--row &lt;N&gt;           | Có           | Index dòng data (1-based, bỏ qua header). VD: `--row 1` = dòng 2 trong Excel.|
| \--docx-dir &lt;path&gt;   | Có           | Thư mục gốc chứa file .docx. Script duyệt đệ quy toàn bộ thư mục con.        |
| \--sheet &lt;name&gt;      | Không        | Tên sheet Excel chứa dữ liệu mẫu. Mặc định: sheet đầu tiên.                  |
| \--config-sheet &lt;name&gt;| Không      | Tên sheet Config chuẩn của KisorDoc. Mặc định nếu không truyền sẽ tìm sheet tên `Config`. |
| \--case-insensitive        | Không        | Bật tìm kiếm không phân biệt hoa thường khi match text trong .docx.           |
| \--min-length &lt;N&gt;    | Không        | Bỏ qua các giá trị mẫu ngắn hơn N ký tự. Mặc định: `3`.                      |
| \--dense-threshold &lt;N&gt;| Không      | Ngưỡng cảnh báo khi một giá trị xuất hiện quá nhiều lần trong 1 file. Mặc định: `15`.|
| \--include / --exclude     | Không        | Glob pattern lọc file .docx (ví dụ: `--exclude "*nhap*.docx"`).               |
| \--max-files &lt;N&gt;     | Không        | Giới hạn số lượng file xử lý (để test an toàn).                              |
| \--verbose                 | Không        | In log chi tiết (hữu ích để debug thuật toán XML run).                       |
| \--dry-run                 | Không        | Nếu có flag này: không sửa file, chỉ xuất report.                            |
| \--report-dir &lt;path&gt; | Không        | Thư mục lưu file report. Hỗ trợ xuất báo cáo kể cả khi chạy migrate thật.    |

## 3.2 Xử lý mapping

- **Đọc cấu hình tên biến (Column Mapping Config):**
  - Đọc bảng mapping từ sheet cấu hình (nếu có): ưu tiên lấy theo định nghĩa Key-Value hiện tại của KisorDoc.
  - **Cơ chế Fallback:** Nếu một cột không có định nghĩa trong bảng mapping, tự động sinh slug bằng thư viện chuẩn (ví dụ `unicodedata.normalize('NFKD', text).encode('ASCII', 'ignore')`). Format: giữ nguyên hoa/thường, khoảng trắng thay bằng gạch dưới `_` (VD: "Họ và Tên" -> "Ho_va_Ten").
- **Tạo cặp mapping (Xử lý Edge Cases):**
  - Đọc row mẫu chỉ định (dựa trên `--row`, bỏ qua header). Bỏ qua giá trị rỗng/null, hoặc có độ dài nhỏ hơn `--min-length`.
  - Thực hiện `strip()` khoảng trắng ở 2 đầu của giá trị mẫu.
  - **Collision Rule:** Nếu 2 cột khác nhau có cùng giá trị mẫu (ví dụ `"1000"` hoặc `"Nguyễn Văn A"`), hệ thống sẽ map theo cột đầu tiên tìm thấy và **in log cảnh báo (warning)**.
  - **Dense Match Warning:** Nếu một giá trị mẫu được match lớn hơn `--dense-threshold N` lần trong một file, hệ thống sẽ in cảnh báo (soft warning) trong report nhưng không chặn luồng chạy.
  - Build mapping hoàn chỉnh: `{ "Nguyễn Văn A" → "{{ Ho_va_Ten }}", ... }`
- **Sắp xếp mapping (Sort by length):**
  - Sắp xếp mapping theo độ dài giá trị mẫu giảm dần - đảm bảo các cụm từ dài được thay thế trước (ví dụ: "Nguyễn Văn An" thay trước "Nguyễn Văn A").

## 3.3 Xử lý XML docx

- Mở file .docx dưới dạng ZIP archive.
- Duyệt đệ quy qua các file XML cấu trúc của văn bản nằm trong `word/` (`document.xml`, `header*.xml`, `footer*.xml`, `footnotes.xml`, `endnotes.xml`).
- Với mỗi paragraph `<w:p>` (bao gồm cả trong bảng biểu `<w:tc>`, textbox, SDT):
  - Reconstruct full text bằng cách ghép nối tất cả `<w:t>` từ các run `<w:r>`. *(Gợi ý Implementation: Nên giữ map ánh xạ vị trí từ chuỗi full-text thuần túy về khoảng span của thẻ XML `(run_index, offset)` để dễ dàng thay thế về sau).*
  - Kiểm tra so khớp (hỗ trợ case-insensitive nếu được bật): nếu không có giá trị mẫu nào xuất hiện → bỏ qua.
  - **Tính Idempotent:** BỎ QUA token đang nằm bên trong một placeholder Jinja2 đã có sẵn (ví dụ: `{{ ... }}`). Chú ý: Chỉ bỏ qua chính xác token đó, KHÔNG bỏ qua toàn bộ đoạn văn (các text thô trong cùng đoạn vẫn được thay thế bình thường). Việc check idempotent trên dấu `{{` và `}}` luôn case-sensitive.
  - **Thay thế an toàn (Safe Run Replacement) - Span nhiều Run:**
    - Xác định chính xác khoảng run (span) chứa từ khóa cần thay thế.
    - Ghi nội dung mới (placeholder) vào run bắt đầu của span. Không cắt đôi placeholder giữa các run.
    - Xóa text tại các run ở giữa (nằm trọn trong span).
    - Tại run kết thúc: Trim (cắt bỏ) phần text bị trùng, giữ nguyên phần text còn lại.
      *(Ví dụ: Từ khóa "Nguyễn Văn A". Run 1: "Nguyễn Văn" → Sửa thành `{{ HoTen }}`. Run 2: " A và..." → Trim phần " A", giữ lại " và...").*
- Ghi đè lại vào file ZIP.

## 3.4 Chế độ Dry-run

- Không ghi bất kỳ thay đổi nào vào file .docx.
- Tính toán preview text dựa trên full-text ghép nối của paragraph (không mô phỏng thao tác cắt XML run). Do đó preview có thể lệch chút đỉnh so với XML thật tại biên các run. Khi migrate thật sẽ dùng safe-run. (Dùng `--verbose` hoặc so file nếu nghi ngờ).
- Ghi nhớ toàn bộ kết quả: file, paragraph, text gốc, text mới, danh sách match.
- Xuất 2 file report (xem mục 3.5 và 3.6).

## 3.5 HTML Report (dry-run)

- Header: thông tin Excel, row mẫu, timestamp.
- Stats bar: tổng file, số file có thay đổi, tổng số đoạn văn sẽ đổi.
- Bảng 4 cột: File | Đoạn văn gốc | Sau khi thay | Biến được thay.
- Highlight: text cũ gạch ngang màu đỏ, placeholder mới màu xanh.
- **Disclaimer (Ghi chú nhỏ):** Lưu ý báo cáo này là bản xem trước trên text ghép nối. Khi migrate thật, biên các thẻ XML có thể gây lệch đôi chút (dùng `--verbose` để debug).
- File không có thay đổi: hiển thị dòng ghi chú "- Không có thay đổi -".
- Tên file: dryrun_YYYYMMDD_HHMMSS.html.

## 3.6 Excel Report (dry-run)

- 5 cột: File | Đoạn văn gốc | Sau khi thay | Biến được thay | Số thay đổi.
- Cột "Đoạn văn gốc": nền đỏ nhạt (#FFE8E8).
- Cột "Sau khi thay": nền xanh nhạt (#E8FFE8).
- Freeze pane tại A2, auto filter toàn bảng.
- Tên file: dryrun_YYYYMMDD_HHMMSS.xlsx.

## 3.7 Backup & Safety

- Trước khi ghi đè, tạo file backup với hậu tố timestamp: `.YYYYMMDD_HHMMSS.bak.docx`. Đảm bảo các lần chạy sau không ghi đè mất backup cũ.
- Bỏ qua toàn bộ các file `*.bak.docx` trong quá trình duyệt xử lý.
- Ghi vào file .tmp trước, chỉ rename → file gốc sau khi thành công.
- Nếu có exception: xóa .tmp, giữ nguyên file gốc (script thoát với exit code != 0 để CI cảnh báo).

# 4\. Yêu cầu phi chức năng

| **Tiêu chí**    | **Yêu cầu**                                                                               |
| --------------- | ----------------------------------------------------------------------------------------- |
| Hiệu năng       | Mục tiêu: 20 file .docx (mỗi file ~50 trang) xử lý xong trong < 60 giây.                  |
| Độ chính xác    | Không sót hoặc thay sai do XML run bị xé. Đảm bảo 100% paragraph có match đều được xử lý. |
| An toàn dữ liệu | File gốc luôn được backup trước khi sửa. Không mất dữ liệu khi script lỗi giữa chừng.     |
| Tương thích     | Python 3.10+, Windows/macOS/Linux. Thư viện: openpyxl, lxml (đều có trên pip).            |
| Encoding        | Mọi file XML đọc/ghi đều dùng UTF-8. Hỗ trợ tiếng Việt có dấu đầy đủ.                     |

# 5\. Luồng xử lý chính

## 5.1 Dry-run flow

python migrate_text_to_placeholders.py \\

\--excel "data.xlsx" --row 1 \\

\--docx-dir "2. Templates/" \\

\--dry-run --report-dir "reports/"

- Script đọc Excel → build mapping (N biến).
- Duyệt toàn bộ .docx trong docx-dir (đệ quy).
- Với mỗi file: mở ZIP → scan XML → tính preview → ghi nhớ kết quả.
- Xuất dryrun_&lt;ts&gt;.html và dryrun_&lt;ts&gt;.xlsx vào report-dir.
- In tóm tắt ra terminal.

## 5.2 Migrate thật flow

python migrate_text_to_placeholders.py \
  --excel "data.xlsx" --row 1 \
  --docx-dir "2. Templates/" --report-dir "reports/"

- Tương tự dry-run nhưng thực sự ghi file.
- Mỗi file: copy → .YYYYMMDD_HHMMSS.bak.docx, ghi nội dung mới → .tmp, rename .tmp → file gốc.
- In summary log ra terminal. Nếu có truyền `--report-dir`, vẫn xuất báo cáo HTML/Excel như dry-run để phục vụ audit.

# 6\. Acceptance Criteria

| **#** | **Tiêu chí**            | **Pass khi**                                                                        |
| ----- | ----------------------- | ----------------------------------------------------------------------------------- |
| AC-01 | Mapping đúng từ Excel   | N biến từ header row được đọc đầy đủ, không bị mất do null hay whitespace.          |
| AC-02 | Thay đúng khi run bị xé | Paragraph có text bị xé thành nhiều &lt;w:r&gt; vẫn được thay đúng placeholder. *(Lưu ý: Chỉ test được bằng migrate thật hoặc qua report của cờ `--verbose`, không test được trên dry-run report).* |
| AC-03 | Sort by length          | "Nguyễn Văn An" được thay trước "Nguyễn Văn A" - không tạo ra kết quả lồng nhau.    |
| AC-04 | Dry-run không sửa file  | Sau --dry-run, checksum SHA256 của file .docx gốc không thay đổi.                   |
| AC-05 | HTML report đúng        | Mỗi paragraph có match hiển thị đúng cột, text cũ highlight đỏ, mới highlight xanh. |
| AC-06 | Excel report đúng       | Filter theo cột File trả về đúng số dòng = số paragraph thay đổi của file đó.       |
| AC-07 | Backup tồn tại          | Sau migrate thật, file `*.YYYYMMDD_HHMMSS.bak.docx` tồn tại và mở được, nội dung = file gốc. |
| AC-08 | Lỗi không mất file      | Nếu inject exception giữa chừng, file gốc nguyên vẹn, .tmp bị xóa, script exit code != 0.           |
| AC-09 | Tiếng Việt đúng         | Giá trị mẫu chứa tiếng Việt có dấu được tìm và thay thế thành công, không bị lỗi font. |
| AC-10 | Trùng giá trị (Collision)| Nếu 2 cột Excel có cùng giá trị, script cảnh báo và áp dụng mapping của cột đầu tiên.|
| AC-11 | Tính Idempotent         | Nếu văn bản có sẵn `{{ X }}`, script chỉ bỏ qua token đó, KHÔNG bỏ qua toàn bộ đoạn văn. Các text thô trong cùng đoạn vẫn được thay thế bình thường. |
| AC-12 | Lọc File Glob           | Script nhận diện đúng `--include` và `--exclude`, luôn tự động bỏ qua `*.bak.docx`. |
| AC-13 | Min-length filter       | `--min-length 3` bỏ qua các giá trị mẫu ngắn hơn 3 ký tự (VD: "1", "Ha").           |
| AC-14 | Timestamp Backup        | File backup luôn được gắn timestamp (VD: `file.20260806_120000.bak.docx`).          |
| AC-15 | Case-insensitive Config | Lookup Config cột "Mã gói" hay "mã gói" đều map đúng Key theo rule `--case-insensitive` (strip + lower). |
| AC-16 | Dense Match Warning     | Nếu 1 giá trị lặp lại > `--dense-threshold`, log warning xuất hiện trong HTML/Excel report. |

# 7\. Cấu trúc file & Output

| **Script chính**     | migrate_text_to_placeholders.py         |
| -------------------- | --------------------------------------- |
| **Input: Excel**     | data.xlsx (header row 1, data từ row 2) |
| **Input: Templates** | 2\. Templates/\*.docx (đệ quy)          |
| **Backup**           | 2\. Templates/\*.YYYYMMDD_HHMMSS.bak.docx (tạo tự động) |
| **Report HTML**      | reports/dryrun_YYYYMMDD_HHMMSS.html     |
| **Report Excel**     | reports/dryrun_YYYYMMDD_HHMMSS.xlsx     |
| **Dependencies**     | Python 3.10+ \| openpyxl \| lxml        |

# 8\. Rủi ro & Giả định

## 8.1 Rủi ro

| **Rủi ro**                                                                                          | **Xác suất** | **Ảnh hưởng** | **Giảm thiểu**                                                |
| --------------------------------------------------------------------------------------------------- | ------------ | ------------- | ------------------------------------------------------------- |
| Giá trị mẫu xuất hiện ở chỗ không mong muốn (VD: số "2024" là năm sinh nhưng cũng có trong địa chỉ) | Trung bình   | Cao           | Luôn chạy dry-run trước, review report kỹ.                    |
| Word xé run theo cách không thể reconstruct đúng (VD: run bị cắt giữa ký tự đặc biệt)               | Thấp         | Trung bình    | Dùng merge_runs nếu gặp; test trên file thật trước khi batch. |
| File .docx bị khóa hoặc corrupt                                                                     | Thấp         | Thấp          | Script bắt exception, bỏ qua file lỗi, in cảnh báo.           |

## 8.2 Giả định

- File Excel luôn có row 1 là header (tên biến hợp lệ cho Jinja2).
- Giá trị mẫu trong Excel khớp chính xác với text trong docx (bao gồm khoảng trắng, dấu câu).
- Môi trường chạy có Python 3.10+ và quyền đọc/ghi thư mục template.

# 9\. Definition of Done (DoD)

- Đáp ứng toàn bộ Acceptance Criteria (AC-01 đến AC-16).
- Script chạy thành công trên các mẫu hợp đồng tiêu biểu của hệ thống mà không phá vỡ cấu trúc file gốc.
- QA review và sign-off bản report Dry-run trước khi cấp quyền chạy migrate thật trên dữ liệu Production.
- Review Code Pass và merge vào nhánh chính.

# 10\. Ghi chú mở rộng

- **Liên kết script cũ:** Trong KisorDoc đã có `migrate_templates.py` (chuyển `<< >>` sang `{{ }}`) và `migrate_modifiers.py` (chuyển `.Date` sang `|date`). Script này thuần túy giải quyết bài toán: text thô → `{{ placeholder }}`. 
- **Tái sử dụng I/O:** Script mới tái sử dụng hoàn toàn luồng đọc/ghi an toàn của các script trên (mở `.docx` qua `zipfile`, thao tác trên RAM, ghi sang `.tmp`, đổi tên đè file gốc). Tuy nhiên, thuật toán thao tác XML lõi phải dùng `lxml` thay vì Regex thuần nhằm đảm bảo không phá vỡ cấu trúc thẻ khi xử lý văn bản phân mảnh qua nhiều `<w:r>`. Pattern Regex tìm `{{...}}` cũ được tái sử dụng để làm Idempotency Scan.
- **Thứ tự đề xuất (Pipeline):** Nhờ cơ chế Idempotency, việc chạy script này trước hay sau `migrate_templates` đều an toàn. Tuy nhiên, thứ tự tối ưu là: `migrate_text_to_placeholders` (với raw data) → `migrate_templates` / `migrate_modifiers` (nếu cần đổi format) → `runner dry-run generate`.
- **ASCII Decision Flow (Config & Fallback):**
  ```text
  [Header Excel] --> (Tên Cột)
                       |
                       v
            [Tìm trong Config Sheet?]
              /                  \
          (Có)                  (Không)
           |                       |
           v                       v
      [Lấy Key]             [Sinh Slug (unicodedata)]
           |                       |
           +-----------> (Chuẩn hóa clean_config_key)
                                   |
                                   v
                             {{ Key }}
  ```

# 11\. Lịch sử thay đổi

| **Phiên bản** | **Ngày**  | **Tác giả** | **Mô tả**                                               |
| ------------- | --------- | ----------- | ------------------------------------------------------- |
| 1.0           | 31/7/2026 | -           | Khởi tạo PRD cho chức năng Migrate Text to Placeholder. |
| 1.1           | 06/8/2026 | KisorDoc AI | Bổ sung quy tắc giải quyết trùng lặp, Idempotent, timestamp backup, làm rõ nguồn mapping Config. |
| 1.2           | 06/8/2026 | KisorDoc AI | Sửa rule Config Key/Value, làm rõ phạm vi XML (textbox/SDT best effort), cảnh báo match-dense, đồng bộ AC-07. Trạng thái Approved. |
| 1.3           | 06/8/2026 | KisorDoc AI | Cập nhật Safe Run Replacement example, Fallback slug algorithm, Report disclaimer, thêm Flowchart và AC-15/AC-16. |