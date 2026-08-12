# Changelog

## [5.2.2] - 2026-08-12

### Added
- **Tự động chuyển tab sau khi chạy**: Thiết lập giao diện tự động chuyển từ tab "1. Chọn & Chạy" sang tab "2. Log & Kết quả" ngay sau khi tiến trình chạy chính hoặc chạy lại file lỗi hoàn tất.

## [5.2.1] - 2026-08-12

### Added
- **Tự động nhận diện và ẩn cột phụ khi copy bảng**: Bổ sung cơ chế tự động ẩn các cột nháp, cột phụ hoặc cột trung gian của Excel (bắt đầu bằng `_`, trùng khớp với các từ khóa `phụ`/`helper`/`temp`/`nháp`/`draft` hoặc chứa các nhãn dạng `[phụ]`/`(phụ)`...) khi thực hiện trích xuất dữ liệu copy sang Word.
- **Sửa lỗi mở thư mục chứa khoảng trắng**: Chuyển đổi lệnh mở thư mục trong nền Windows sang `Invoke-Item` của PowerShell, khắc phục hoàn toàn sự cố phân tách đối số do khoảng trắng gây ra việc mở nhầm thư mục *Documents*.

## [5.2.0] - 2026-08-11

### Added
- **Hỗ trợ WHERE trong cú pháp SQL Join**: Bổ sung hỗ trợ tùy biến thêm điều kiện lọc dữ liệu `WHERE` vào trực tiếp cú pháp join rút gọn (ví dụ: `GoiThau * TCGTTD @ GoiThau_ID WHERE GoiThau.GoiThau_HTDT == 'DTRR'`).
- **Mở thư mục dạng Single Instance**: Cải tiến chức năng mở thư mục output và log sang dạng Single Instance thông qua PowerShell COM Object (Shell.Application), tự động focus vào cửa sổ đang mở thay vì sinh nhiều cửa sổ mới chồng chéo trên Windows.
- **Bảo trì và sửa lỗi cốt lõi (10 Bug Fixes)**:
  - Khắc phục lỗi lọc gói thầu quy trình Repeat khi cấu hình có chứa điều kiện WHERE.
  - Đăng ký và chuẩn hóa đầy đủ các custom filters (`filter_upper`, `upper`, `filter_num2text`, `num2text`) để tránh crash lỗi `UndefinedError` trên template.
  - Tích hợp fallback lấy khóa liên kết theo nhiều cột tùy chọn (`GoiThau_ID`, `MS_GoiThau`, `ID`) thay vì hardcode.
  - Khắc phục lỗi nạp tham số `option` của route `get_packages` trên FastAPI.
  - Suppress cảnh báo `UserWarning: Data Validation extension is not supported` phiền toái từ thư viện `openpyxl`.

## [5.1.0] - 2026-08-10

### Added
- **Core Migrator Mới (kisorlib/migrator.py)**: Gộp logic của cả 3 script cũ (migrate_templates.py, migrate_modifiers.py, migrate_table_braces.py) thành 1 library duy nhất.
- Hỗ trợ API: migrate_xml, migrate_file, migrate_folder để migrate giao diện/template linh hoạt, hỗ trợ dry_run.
- Bổ sung Unit Tests đầy đủ cho migrator (	ests/test_migrator.py).
- Sửa lỗi regex _RAW_LT_GT_RE bị thiếu dấu . gây lỗi khi replace chuỗi rỗng.

## [5.0.0] - 2026-08-07

### Added
- **Bá»™ 147 Unit Tests Má»›i (Há»“i Quy & Äá»™ Phá»§):**
  - `tests/test_sql_join.py` (+76 tests): Kiá»ƒm nghiá»‡m toÃ n diá»‡n logic xá»­ lÃ½ SQL Join vÃ  resolve truy váº¥n sheet.
  - `tests/test_batch.py` (+48 tests): Bao phá»§ class `IncrementalRunLogger`, cÃ¡c helpers, logic validation vÃ  xá»­ lÃ½ log sá»± kiá»‡n.
  - `tests/test_merger.py` (+23 tests): Kiá»ƒm chá»©ng hoáº¡t Ä‘á»™ng cá»§a cÆ¡ cháº¿ merge tÃ i liá»‡u `mail_merge` vÃ  `mail_merge_safe`.
- **Kháº¯c Phá»¥c Lá»—i Há»‡ Thá»‘ng (Báº£n VÃ¡ Há»“i Quy):**
  - Sá»­a lá»—i parameterized query trong luá»“ng chá»n UI group láº·p cá»§a `app.py`.
  - Tá»‘i Æ°u náº¡p `KisorService` trong `kisorlib/engine.py` (trÃ¡nh query 2 láº§n).
  - TÃ­ch há»£p hÃ m `_safe_table_name` cho cÃ¡c truy váº¥n báº£ng láº·p `_Goc` trong `kisorlib/service.py` Ä‘á»ƒ tÆ°Æ¡ng thÃ­ch vá»›i sheet cÃ³ khoáº£ng tráº¯ng/kÃ½ tá»± Ä‘áº·c biá»‡t.
  - Kháº¯c phá»¥c lá»—i copy template giá»‘ng nhau `src == dst` trÃªn Windows gÃ¢y `SameFileError`.
  - Kháº¯c phá»¥c lá»—i xá»­ lÃ½ warning thá»‘ng kÃª log file trong `kisorlib/batch.py`.

## [4.2.0] - 2026-08-06

### Added
- **TÃ­nh nÄƒng Migrate Text to Jinja2 Placeholder:** Tá»± Ä‘á»™ng chuyá»ƒn Ä‘á»•i cÃ¡c tá»« khÃ³a/giÃ¡ trá»‹ máº«u trong Word (.docx) sang tháº» placeholder Jinja2 sá»­ dá»¥ng sÆ¡ Ä‘á»“ mapping cá»§a Excel.
  - Há»— trá»£ giáº£i quyáº¿t trÃ¹ng láº·p giÃ¡ trá»‹ máº«u (Collision Warning) vÃ  cáº£nh bÃ¡o táº§n suáº¥t xuáº¥t hiá»‡n quÃ¡ cao (Dense Match Warning).
  - Thuáº­t toÃ¡n **Safe Run Replacement** báº±ng `lxml` xá»­ lÃ½ an toÃ n text bá»‹ Word xÃ© nhá» ra nhiá»u run, báº£o toÃ n Ä‘á»‹nh dáº¡ng in nghiÃªng, in Ä‘áº­m xung quanh.
  - TÃ­nh idempotent ngÄƒn ngá»«a viá»‡c wrap Ä‘Ã¨ placeholder hai láº§n.
  - Há»— trá»£ cháº¿ Ä‘á»™ dry-run xuáº¥t bÃ¡o cÃ¡o trá»±c quan dáº¡ng HTML vÃ  Excel.
- **Unit Tests Há»“i Quy:** Bá»• sung file test má»›i [tests/test_migrate_text_to_placeholders.py](tests/test_migrate_text_to_placeholders.py) kiá»ƒm thá»­ 100% cÃ¡c logic nghiá»‡p vá»¥ lÃµi (to_slug, safe_replace, idempotent, load_mapping).

## [4.1.3] - 2026-08-06

### Fixed
- **LOG-01 (batch.py):** Sá»­a lá»—i Ä‘áº¿m thá»‘ng kÃª, tÄƒng `ok_count` khi gáº·p cáº£nh bÃ¡o nhÆ°ng sinh file thÃ nh cÃ´ng.
- **CFG-01 (batch.py & engine.py):** Chuyá»ƒn Ä‘á»•i retry delay vÃ  max retries sá»­ dá»¥ng Ä‘á»™ng tá»« cáº¥u hÃ¬nh `AppConfig`.
- **BAT-01 (batch.py):** NgÄƒn cháº·n viá»‡c append file bá»‹ lá»—i sao chÃ©p vÃ o danh sÃ¡ch render tiáº¿p theo.
- **ENG-01 (engine.py):** Tá»‘i Æ°u hÃ³a khá»Ÿi táº¡o `KisorService` má»™t láº§n duy nháº¥t.
- **GEN-01 (generator.py):** Kháº¯c phá»¥c lá»—i `SameFileError` trÃªn Windows khi sao chÃ©p cÃ¹ng má»™t Ä‘Æ°á»ng dáº«n.
- **SVC-01 (service.py):** Äá»“ng bá»™ hÃ³a tÃªn báº£ng táº¡m thÃ nh viÃªn cá»§a Repeat mode (`_Goc`) qua hÃ m `_safe_table_name` Ä‘á»ƒ trÃ¡nh lá»—i vá»›i tÃªn Sheet cÃ³ khoáº£ng tráº¯ng hoáº·c kÃ½ tá»± Ä‘áº·c biá»‡t.

## [4.1.2] - 2026-08-05

### Added
- **ÄuÃ´i file máº·c Ä‘á»‹nh cho cá»™t File khi copy báº£ng:** Há»— trá»£ tá»± Ä‘á»™ng Ä‘iá»n Ä‘uÃ´i tá»‡p tin `.xlsx` náº¿u tÃªn file khai bÃ¡o trong cá»™t `File` cá»§a báº£ng cáº¥u hÃ¬nh bá»‹ thiáº¿u pháº§n má»Ÿ rá»™ng.

## [4.1.1] - 2026-08-05

### Fixed
- **TrÃ¡nh lá»—i tá»± sao chÃ©p tá»‡p tin trÃªn Windows:** Kháº¯c phá»¥c lá»—i `PermissionError` (Sharing violation) khi core engine cá»‘ gáº¯ng sao chÃ©p má»™t tá»‡p tin Ä‘Ã¨ lÃªn chÃ­nh nÃ³ táº¡i thÆ° má»¥c Ä‘áº§u ra trong `generate_one`, `generate_one_repeat` vÃ  `copy_templates_to_output`.
- **Kháº¯c phá»¥c lá»—i náº¡p template cho Quy trÃ¬nh Láº·p (Repeat Type):** Chá»‰nh sá»­a hÃ m `get_workflow_templates` Ä‘á»ƒ chá»‰ truy váº¥n báº£ng gÃ³i tháº§u (`left_sheet`) khi láº¥y danh sÃ¡ch template cá»§a Quy trÃ¬nh Láº·p. Viá»‡c nÃ y loáº¡i bá» hoÃ n toÃ n lá»—i cÆ¡ sá»Ÿ dá»¯ liá»‡u `BinderException` do báº£ng thÃ nh viÃªn (`TCGTTD`) chÆ°a Ä‘Æ°á»£c táº¡o cá»™t khÃ³a phá»¥ Ä‘á»™ng (`GoiThau_ID`) á»Ÿ thá»i Ä‘iá»ƒm táº£i cáº¥u hÃ¬nh.

## [4.1.0] - 2026-08-05

### Added
- **Module Core Sync `kisorlib/generator.py`:** TÃ­ch há»£p logic sinh tÃ i liá»‡u Ä‘á»“ng bá»™ lÃ m Single Source of Truth cho cáº£ UI (`batch.py`) vÃ  API (`engine.py`). Giáº£i quyáº¿t triá»‡t Ä‘á»ƒ lá»—i dual-pipeline.
- **TÃ¡ch module `kisorlib/sql_join.py`:** Di chuyá»ƒn toÃ n bá»™ cÃ¡c hÃ m tiá»‡n Ã­ch vÃ  xá»­ lÃ½ SQL Join biá»ƒu thá»©c (`parse_join_expression`, `resolve_sheet_query`, `validate_sql_identifier`,...) tá»« `utils.py` sang module má»›i.
- **Unit Tests Há»“i Quy:** Bá»• sung file test má»›i [tests/test_generator.py](tests/test_generator.py) vá»›i 25 test cases kiá»ƒm chá»©ng chi tiáº¿t hoáº¡t Ä‘á»™ng cá»§a core `generator` (sinh 1 file, nhiá»u file, repeat mode, lock/retry).

### Changed
- **TÃ¡i cáº¥u trÃºc `kisorlib/batch.py`:** Loáº¡i bá» hoÃ n toÃ n logic sinh file, merge docx, chá»‰ giá»¯ láº¡i vai trÃ² async/yield progress wrapper vÃ  ghi log Gradio qua `IncrementalRunLogger`.
- **TÃ¡i cáº¥u trÃºc `kisorlib/engine.py`:** Chuyá»ƒn Ä‘á»•i gá»i trá»±c tiáº¿p xuá»‘ng core `generator.py` thay vÃ¬ xá»­ lÃ½ merge ná»™i bá»™.
- **Cáº­p nháº­t `kisorlib/utils.py`:** Loáº¡i bá» cÃ¡c hÃ m SQL trÃ¹ng láº·p, chá»‰ giá»¯ láº¡i Ä‘á»‹nh dáº¡ng chuá»—i vÃ  AST cÆ¡ báº£n, Ä‘á»“ng thá»i re-export tá»« `sql_join.py` Ä‘á»ƒ tÆ°Æ¡ng thÃ­ch ngÆ°á»£c.
- **Cáº­p nháº­t TÃ i liá»‡u Lá»—i:** ÄÃ¡nh dáº¥u dual-pipeline Ä‘Ã£ Ä‘Æ°á»£c xá»­ lÃ½ hoÃ n toÃ n trong [known-issues.md](docs/known-issues.md).

## [4.0.3] - 2026-08-05

### Added
- **Báº£o máº­t Äá»‹nh danh SQL (SQL Identifier Whitelist):** Bá»• sung hÃ m `validate_sql_identifier` vá»›i regex whitelist `^[A-Za-z0-9_\s\-\.\#\u00C0-\u1EF9]+$` Ä‘á»ƒ lá»c sáº¡ch toÃ n bá»™ tÃªn báº£ng vÃ  tÃªn cá»™t Ä‘á»™ng Ä‘á»c tá»« Excel trÆ°á»›c khi thá»±c hiá»‡n SQL. Triá»‡t tiÃªu hoÃ n toÃ n nguy cÆ¡ SQL Identifier Injection.
- **Má»Ÿ rá»™ng Unit Tests:** Bá»• sung cÃ¡c test case trong [test_utils.py](tests/test_utils.py#L118) kiá»ƒm chá»©ng hoáº¡t Ä‘á»™ng cá»§a hÃ m whitelist vá»›i cÃ¡c chuá»—i Ä‘á»™c háº¡i.

### Changed
- **Cáº­p nháº­t TÃ i liá»‡u Lá»—i:** ÄÃ¡nh dáº¥u lá»—i SQL Identifier Injection Ä‘Ã£ Ä‘Æ°á»£c kháº¯c phá»¥c hoÃ n toÃ n trong [known-issues.md](docs/known-issues.md#L65).

## [4.0.2] - 2026-08-05

### Changed
- **TÃ i liá»‡u hÃ³a cÃ¡c váº¥n Ä‘á» kiáº¿n trÃºc (known-issues.md):** Cáº­p nháº­t tÃ¬nh hÃ¬nh sá»­a lá»—i kiáº¿n trÃºc, dá»n dáº¹p hardcode, báº£o máº­t SQL Parameter Binding vÃ  ghi nháº­n rá»§i ro SQL Identifier Injection tiá»m áº©n.
- **Tá»•ng quÃ¡t hÃ³a giao diá»‡n:** Cáº­p nháº­t máº·c Ä‘á»‹nh nhÃ³m láº·p trong [ui_labels.json](ui_labels.json) tá»« gÃ¡n cá»©ng `"Tá»• chuyÃªn gia" / "Tá»• tháº©m Ä‘á»‹nh"` thÃ nh `"NhÃ³m láº·p 1" / "NhÃ³m láº·p 2"`.

### Fixed
- **Sá»­a lá»—i Ä‘á»“ng bá»™ Radio State cá»§a Gradio:** Kháº¯c phá»¥c lá»—i `gradio.exceptions.Error` khi khá»Ÿi táº¡o/reset app vá» quy trÃ¬nh trá»‘ng báº±ng cÃ¡ch khÃ´i phá»¥c láº¡i cÃ¡c lá»±a chá»n máº·c Ä‘á»‹nh vÃ  giÃ¡ trá»‹ há»£p lá»‡ cho `group_radio`.

## [4.0.1] - 2026-08-05

### Changed
- **NÃ¢ng cáº¥p báº£o máº­t SQL (Parameter Binding):** Chuyá»ƒn Ä‘á»•i toÃ n bá»™ cÃ¡c cÃ¢u lá»‡nh SQL ghÃ©p chuá»—i giÃ¡ trá»‹ trong `kisorlib/service.py` sang parameter binding `?` cá»§a DuckDB Ä‘á»ƒ ngÄƒn cháº·n SQL Injection vÃ  lá»—i cÃº phÃ¡p dá»¯ liá»‡u chá»©a nhÃ¡y Ä‘Æ¡n.
- **Dá»n dáº¹p hardcode cÃ²n láº¡i trong api.py:** Cáº­p nháº­t `api.py` sá»­ dá»¥ng thuá»™c tÃ­nh Ä‘Æ°á»ng dáº«n Ä‘á»™ng `cfg.data_path` vÃ  tÃªn báº£ng Ä‘á»™ng `cfg.DataSheet`, Ä‘á»“ng thá»i sá»­a lá»—i khá»Ÿi táº¡o `DataSet` truyá»n tham sá»‘ `excel_files` khÃ´ng cÃ²n Ä‘Æ°á»£c há»— trá»£.

### Added
- **Má»Ÿ rá»™ng Unit Tests:** Bá»• sung file test chuyÃªn biá»‡t [tests/test_service.py](tests/test_service.py) Ä‘á»ƒ bao phá»§ `KisorService`, Repeat mode, Ä‘Äƒng kÃ½ báº£ng táº¡m thÃ nh viÃªn vÃ  preview composite key.

## [4.0.0] - 2026-08-05

### Changed
- **TÃ¡i cáº¥u trÃºc mÃ£ nguá»“n app.py toÃ n diá»‡n:** PhÃ¢n rÃ£ file `app.py` khá»•ng lá»“ dÃ i ~1780 dÃ²ng thÃ nh cÃ¡c module nhá», Ä‘Æ¡n nhiá»‡m bÃªn trong thÆ° má»¥c `kisorlib/` Ä‘á»ƒ nÃ¢ng cao cháº¥t lÆ°á»£ng code vÃ  dá»… báº£o trÃ¬:
  - [kisorlib/utils.py](kisorlib/utils.py): CÃ¡c hÃ m tiá»‡n Ã­ch thuáº§n tÃºy (pure functions).
  - [kisorlib/service.py](kisorlib/service.py): Lá»›p nghiá»‡p vá»¥ `KisorService` loáº¡i bá» global state `config`/`ds` thÃ´ng qua Dependency Injection.
  - [kisorlib/batch.py](kisorlib/batch.py): HÃ m sinh tÃ i liá»‡u hÃ ng loáº¡t `run_batch` vÃ  `run_retry_batch` Ä‘á»™c láº­p hoÃ n toÃ n vá»›i thÆ° viá»‡n Gradio.
  - [app.py](app.py): RÃºt gá»n xuá»‘ng chá»‰ cÃ²n ~360 dÃ²ng Ä‘Ã³ng vai trÃ² káº¿t ná»‘i UI Gradio.
- **Dá»n dáº¹p mÃ£ trÃ¹ng láº·p trong engine.py:** XÃ³a bá» cÃ¡c hÃ m clone trÃ¹ng láº·p trÆ°á»›c Ä‘Ã³ (`_clean_config_key`, `_get_option_config_from_ds`) trong `kisorlib/engine.py` vÃ  tÃ¡i sá»­ dá»¥ng trá»±c tiáº¿p tá»« `utils.py` vÃ  `service.py`.

### Added
- **Unit Tests tá»± Ä‘á»™ng:** Bá»• sung thÆ° má»¥c kiá»ƒm thá»­ `tests/` vá»›i cÃ¡c bá»™ test tá»± Ä‘á»™ng [tests/test_utils.py](tests/test_utils.py) vÃ  [tests/test_filters.py](tests/test_filters.py) giÃºp cháº¡y há»“i quy nhanh chÃ³ng.

### Fixed
- **Sá»­a lá»—i Preview á»Ÿ cháº¿ Ä‘á»™ Repeat:** Kháº¯c phá»¥c triá»‡t Ä‘á»ƒ lá»—i phÃ¢n tÃ­ch sai ID gÃ³i tháº§u dáº¡ng composite (`key_id` chá»©a `|`) trong chá»©c nÄƒng Preview (`run_preview`).

## [3.2.3] - 2026-08-04

### Fixed
- **TÆ°Æ¡ng thÃ­ch hoÃ n toÃ n parser cho Join Expression:** Viáº¿t láº¡i hÃ m `_parse_repeat_sheet_config` sá»­ dá»¥ng cÃ¹ng báº£ng Ã¡nh xáº¡ toÃ¡n tá»­ `_OP_MAP` vá»›i `parse_join_expression`. Fix triá»‡t Ä‘á»ƒ lá»—i phÃ¢n tÃ­ch sai tÃªn sheet trÃ¡i/pháº£i (`left_sheet`, `right_sheet`) Ä‘á»‘i vá»›i cÃ¡c cÃº phÃ¡p join phá»©c táº¡p nhÆ° `<*>`, `<*`, `*>`, `*`.
- **An toÃ n hÃ³a contract write_with_retry:** Chuáº©n hÃ³a kiá»ƒu tráº£ vá» cá»§a hÃ m `write_with_retry` luÃ´n tráº£ vá» tuple `(bool, str)`. NgÄƒn cháº·n hoÃ n toÃ n lá»—i runtime `TypeError: cannot unpack non-iterable` khi hÃ m con bÃªn trong khÃ´ng tráº£ vá» tuple (vÃ­ dá»¥: `do_copy` tráº£ vá» `None`).
- **Kháº¯c phá»¥c stale data trong cháº¿ Ä‘á»™ Repeat:** ThÃªm guard clause kiá»ƒm tra káº¿t quáº£ tráº£ vá» cá»§a hÃ m `register_temporary_tcgttd` trÆ°á»›c khi thá»±c hiá»‡n cÃ¢u lá»‡nh SQL chÃ­nh. Bá» qua vÃ  ghi log `SKIP` cho thÃ nh viÃªn bá»‹ lá»—i thay vÃ¬ truy váº¥n Ä‘Ã¨ lÃªn dá»¯ liá»‡u cÅ© cá»§a thÃ nh viÃªn trÆ°á»›c Ä‘Ã³.
- **Äá»“ng bá»™ split KeyId:** Cáº­p nháº­t cÃ¡c hÃ m thay Ä‘á»•i dá»¯ liá»‡u trÃªn giao diá»‡n (`on_package_change`, `on_group_change`) Ä‘á»ƒ phÃ¢n tÃ¡ch KeyId thÃ´ng qua `_parse_repeat_key_id` trÆ°á»›c khi truy xuáº¥t giÃ¡ trá»‹ tá»« dict, sá»­a lá»—i láº¥y rá»—ng `goi_thau_id` khi cáº¥u hÃ¬nh KeyId dáº¡ng ghÃ©p.
- **Loáº¡i bá» triá»‡t Ä‘á»ƒ hardcode cÃ²n láº¡i:** Chuyá»ƒn Ä‘á»•i toÃ n bá»™ cÃ¡c fallback gÃ¡n cá»©ng tÃªn sheet `"GoiThau"`, tÃªn cá»™t `"ID"`, dáº¡ng show `"{TT}"` sang sá»­ dá»¥ng giÃ¡ trá»‹ cáº¥u hÃ¬nh tÆ°Æ¡ng á»©ng trong `AppConfig` (`config.DataSheet`, `config.DefaultKeyId`, `config.DefaultShow`). Giao diá»‡n chá»n nhÃ³m láº·p cÅ©ng Ä‘Æ°á»£c Ä‘á»™ng hÃ³a khÃ´ng cÃ²n cá»©ng `"Tá»• chuyÃªn gia" / "Tá»• tháº©m Ä‘á»‹nh"`.

## [3.2.2] - 2026-08-04

### Added
- **Äá»™ng hÃ³a hoÃ n toÃ n cáº¥u hÃ¬nh `KeyId` ghÃ©p `|` cho cháº¿ Ä‘á»™ Repeat:** Há»— trá»£ cáº¥u hÃ¬nh `KeyId` dáº¡ng ghÃ©p báº±ng dáº¥u `|` (VÃ­ dá»¥: `GoiThau_ID | CCCD`), tá»± Ä‘á»™ng phÃ¢n tÃ¡ch thÃ nh `left_key` (khÃ³a báº£ng chÃ­nh) vÃ  `right_key` (khÃ³a báº£ng con) Ä‘á»ƒ JOIN query chÃ­nh xÃ¡c tuyá»‡t Ä‘á»‘i, trÃ¡nh trÃ¹ng láº·p há» tÃªn thÃ nh viÃªn khi xá»­ lÃ½ láº·p.
- **TrÃ­ch xuáº¥t tÃªn cá»™t Há» tÃªn Ä‘á»™ng tá»« cá»™t `Show`:** Thay vÃ¬ gÃ¡n cá»©ng `"Há» vÃ  tÃªn"`, chÆ°Æ¡ng trÃ¬nh tá»± Ä‘á»™ng trÃ­ch xuáº¥t tÃªn cá»™t Ä‘á»‹nh danh thÃ nh viÃªn tá»« pháº§n bÃªn pháº£i cá»§a cá»™t `Show` (sau dáº¥u `|`) Ä‘á»ƒ Ä‘Æ°a giÃ¡ trá»‹ chuáº©n vÃ o file template Word.
- **NhÃ¢n báº£n báº£ng DuckDB trÃ¡nh ghi Ä‘Ã¨ dá»¯ liá»‡u gá»‘c:** Tá»± Ä‘á»™ng táº¡o báº£n sao dá»± phÃ²ng `_Goc` cho toÃ n bá»™ cÃ¡c báº£ng trong DuckDB khi khá»Ÿi Ä‘á»™ng, giÃºp cÃ¡c quy trÃ¬nh xá»­ lÃ½ láº·p song song hoáº·c tuáº§n láº·p khÃ´ng bá»‹ ghi Ä‘Ã¨ hay máº¥t dá»¯ liá»‡u gá»‘c cá»§a sheet Excel.

### Changed
- **Default Show Format an toÃ n:** Chuyá»ƒn giÃ¡ trá»‹ máº·c Ä‘á»‹nh cá»§a `show` trong `get_option_config` khi khÃ´ng cÃ³ cáº¥u hÃ¬nh thÃ nh `"{TT}"` thay vÃ¬ hardcode cá»™t cá»§a gÃ³i tháº§u cá»¥ thá»ƒ.
- **Äá»™ng hÃ³a `left_sheet` cho Repeat:** Cáº­p nháº­t `get_packages` Ä‘á»ƒ truy xuáº¥t báº£ng chÃ­nh thÃ´ng qua `left_sheet` láº¥y tá»« cá»™t `Sheet` trong Options thay vÃ¬ cá»©ng nháº¯c `"GoiThau"`.
- **Äá»™ng hÃ³a `DANH_MUC_FILE` tá»« `.env`:** Di chuyá»ƒn cáº¥u hÃ¬nh tÃªn file Danh Má»¥c dá»± Ã¡n thÃ nh biáº¿n mÃ´i trÆ°á»ng `DANH_MUC_FILE` Ä‘á»ƒ thuáº­n tiá»‡n tÃ¹y biáº¿n.
- **VÃ´ hiá»‡u hÃ³a tá»± Ä‘á»™ng báº­t trÃ¬nh duyá»‡t web:** Táº¯t cÆ¡ cháº¿ tá»± Ä‘á»™ng gá»i `webbrowser.open` táº¡i `runner.py` vÃ  `app.py` khi khá»Ÿi Ä‘á»™ng/náº¡p láº¡i code Ä‘á»ƒ trÃ¡nh má»Ÿ tab rÃ¡c trÃªn trÃ¬nh duyá»‡t cá»§a ngÆ°á»i dÃ¹ng.

## [3.2.1] - 2026-08-03

### Changed
- **TÃ¡i cáº¥u trÃºc thÆ° viá»‡n dÃ¹ng chung (Patch-v8):** Di chuyá»ƒn cÃ¡c class/hÃ m helper `NestedVal` vÃ  `make_nested_dict` tá»« `app.py` vÃ o module dÃ¹ng chung `kisorlib/app_helpers.py` Ä‘á»ƒ cáº£ `app.py` vÃ  `engine.py` cÃ¹ng chia sáº», giáº£m thiá»ƒu trÃ¹ng láº·p mÃ£ vÃ  tÄƒng Ä‘á»™ á»•n Ä‘á»‹nh cá»§a há»‡ thá»‘ng.
- Cáº­p nháº­t vÃ  tá»‘i Æ°u hÃ³a an toÃ n relative imports bÃªn trong core library `kisorlib`.

## [3.2.0] - 2026-08-03

### Changed
- **Äá»•i tÃªn Core Package (`kisorlib`):** Äá»•i tÃªn toÃ n bá»™ thÆ° má»¥c thÆ° viá»‡n dÃ¹ng chung tá»« `kisordoc/` thÃ nh `kisorlib/` Ä‘á»ƒ tÄƒng tÃ­nh chuyÃªn nghiá»‡p, Ä‘á»“ng thá»i cáº­p nháº­t toÃ n bá»™ import trong `app.py`, `api.py`, `runner.py`, vÃ  `tests/test_engine.py`.
- **Äá»‹nh dáº¡ng Date khoáº£ng tráº¯ng chá»«a trá»‘ng:** Cá»‘ Ä‘á»‹nh **3 khoáº£ng tráº¯ng** cho ngÃ y vÃ  thÃ¡ng khi bá»‹ trá»‘ng dá»¯ liá»‡u trÃªn Excel (VÃ­ dá»¥: `"ngÃ y   thÃ¡ng 07 nÄƒm 2026"`, `"ngÃ y   thÃ¡ng   nÄƒm 2026"`).

## [3.1.0] - 2026-08-03

### Added
- **Xá»­ lÃ½ ngÃ y thÃ¡ng trá»‘ng má»™t pháº§n (Chá»«a khoáº£ng trá»‘ng ghi tay):** Bá»• sung logic xá»­ lÃ½ cho filter ngÃ y thÃ¡ng (nhÆ° `.Date.Long`), tá»± Ä‘á»™ng Ä‘á»‹nh dáº¡ng cÃ¡c chuá»—i ngÃ y thÃ¡ng chá»©a dáº¥u gáº¡ch chÃ©o `/` nhÆ°ng bá»‹ khuyáº¿t thÃ´ng tin ngÃ y hoáº·c thÃ¡ng (VD: `"   /07/2026"`, `"  /   /2026"`) thÃ nh `"ngÃ y   thÃ¡ng 07 nÄƒm 2026"` vÃ  `"ngÃ y   thÃ¡ng   nÄƒm 2026"`, máº·c Ä‘á»‹nh sá»­ dá»¥ng 3 khoáº£ng tráº¯ng cho pháº§n bá»‹ trá»‘ng Ä‘á»ƒ ghi tay sau.
- **Cáº£i tiáº¿n giao diá»‡n:** Cáº¥u hÃ¬nh cá»™t template bÃªn pháº£i luÃ´n hiá»ƒn thá»‹ (`visible=True`) ngay tá»« Ä‘áº§u Ä‘á»ƒ trÃ¡nh co giÃ£n layout, Ä‘á»“ng thá»i giá»¯ nguyÃªn logic cáº­p nháº­t danh sÃ¡ch Ä‘á»™ng khi ngÆ°á»i dÃ¹ng chá»n Quy trÃ¬nh & GÃ³i tháº§u.

### Fixed
- Kháº¯c phá»¥c lá»—i `Binder Error` khi cháº¡y quy trÃ¬nh láº·p `Repeat` (`Opt6`) do cÃ¢u lá»‡nh SQL Join bá»‹ rá»—ng lÃºc chÆ°a náº¡p danh sÃ¡ch thÃ nh viÃªn táº¡m thá»i.
- Kháº¯c phá»¥c lá»—i thiáº¿u thÆ° viá»‡n `os` khi nháº¥n nÃºt má»Ÿ thÆ° má»¥c log/output trÃªn giao diá»‡n.

## [3.0.0] - 2026-08-03

### Added
- **Core Library Packaging (`kisordoc/`):** TÃ¡i cáº¥u trÃºc Ä‘Ã³ng gÃ³i toÃ n bá»™ logic xá»­ lÃ½ nghiá»‡p vá»¥ (`config.py`, `dataset.py`, `table_copier.py`, `merger.py`, `file_utils.py`, `filters.py`) vÃ o package `kisordoc/`.
- **Core Engine API (`kisordoc/engine.py`):** XÃ¢y dá»±ng Ä‘iá»ƒm vÃ o duy nháº¥t (Public API) cho cÃ¡c tÃ¡c vá»¥ mail-merge vÃ  dry-run sá»­ dá»¥ng Pydantic models (`GenerateRequest`/`GenerateResult`) vÃ  cÆ¡ cháº¿ callback tiáº¿n trÃ¬nh `on_progress`.
- **TÃ­ch há»£p FastAPI Backend (`api.py`):** Cung cáº¥p cÃ¡c RESTful API endpoints `/generate`, `/templates`, `/packages` tá»± Ä‘á»™ng sinh tÃ i liá»‡u Swagger.
- **Khá»Ÿi cháº¡y song song (`runner.py`):** Há»— trá»£ khá»Ÿi cháº¡y Ä‘á»“ng thá»i Gradio UI (`app.py` á»Ÿ cá»•ng 7864) vÃ  FastAPI API (`api.py` á»Ÿ cá»•ng 8000) thÃ´ng qua thread an toÃ n chá»‰ báº±ng má»™t lá»‡nh duy nháº¥t.
- **TÃ¡ch biá»‡t cáº¥u hÃ¬nh nhÃ£n giao diá»‡n (`ui_labels.json`):** Chuyá»ƒn toÃ n bá»™ chuá»—i kÃ½ tá»± hiá»ƒn thá»‹ trÃªn Gradio UI ra file cáº¥u hÃ¬nh JSON Ä‘á»™c láº­p giÃºp thay Ä‘á»•i nhÃ£n Ä‘á»™ng khÃ´ng cáº§n sá»­a code.
- **Cáº£i tiáº¿n giao diá»‡n chá»n template:** Tá»± Ä‘á»™ng áº©n cá»™t chá»n file template khi chÆ°a chá»n gÃ³i tháº§u vÃ  chá»‰ hiá»ƒn thá»‹ sau khi Ä‘Ã£ náº¡p dá»¯ liá»‡u thÃ nh cÃ´ng.

### Fixed
- Kháº¯c phá»¥c lá»—i `âŒ KhÃ´ng tÃ¬m tháº¥y dÃ²ng dá»¯ liá»‡u tÆ°Æ¡ng á»©ng` khi cháº¡y batch hoáº·c dry-run Ä‘á»‘i vá»›i cÃ¡c Option Ä‘áº·c thÃ¹ (nhÆ° Mua sáº¯m nhá» Opt1 láº¥y tá»« báº£ng `MuaSamNho`).
- Kháº¯c phá»¥c lá»—i crash do component Textbox cá»§a phiÃªn báº£n Gradio má»›i khÃ´ng há»— trá»£ tham sá»‘ `show_copy_button`.
- Sá»­a lá»—i nÃºt Má»Ÿ thÆ° má»¥c log/output trÃªn Windows console báº±ng cÃ¡ch chuyá»ƒn sang `subprocess.Popen`.

## [2.2.2] - 2026-08-03

### Added
- **Xá»­ lÃ½ File Locked nÃ¢ng cao (F6):** Báº¯t lá»—i `PermissionError` (khi Word chiáº¿m dá»¥ng file), tá»± Ä‘á»™ng retry tá»‘i Ä‘a `FILE_MAX_RETRIES` láº§n sau má»—i `FILE_RETRY_DELAY` giÃ¢y. ÄÃ¡nh dáº¥u tráº¡ng thÃ¡i file bá»‹ khÃ³a dáº¡ng `ðŸ”’` thay vÃ¬ `âŒ`.
- **Validation trÆ°á»›c khi cháº¡y (F1):** Cháº·n sá»›m viá»‡c cháº¡y batch vÃ  thÃ´ng bÃ¡o lá»—i rÃµ rÃ ng náº¿u thiáº¿u Option, GÃ³i tháº§u, Template hoáº·c dá»¯ liá»‡u chÆ°a sáºµn sÃ ng.
- **Dry-run / Preview Mode (F2):** Bá»• sung nÃºt "ðŸ” Kiá»ƒm tra" trÃªn UI, trÃ­ch xuáº¥t placeholder an toÃ n tá»« ZIP Docx, Ã¡p dá»¥ng custom filters vÃ  hiá»ƒn thá»‹ báº£ng káº¿t quáº£ Preview (`gr.Dataframe`) trá»±c quan.
- **LÆ°u tráº¡ng thÃ¡i & Cháº¡y láº¡i file lá»—i (F3):** Tá»± Ä‘á»™ng phÃ¡t hiá»‡n lá»—i vÃ  hiá»ƒn thá»‹ nÃºt "Cháº¡y láº¡i file lá»—i" Ä‘á»ƒ chá»‰ merge láº¡i cÃ¡c file bá»‹ lá»—i mÃ  khÃ´ng cáº§n xÃ³a/xá»­ lÃ½ láº¡i cÃ¡c file thÃ nh cÃ´ng.
- **Export log ra file text (F4):** Ghi log incremental khi cháº¡y, Ä‘á»‹nh dáº¡ng `utf-8-sig` chuáº©n Notepad Windows, tá»± Ä‘á»™ng dá»n dáº¹p log cÅ© >30 ngÃ y (giá»›i háº¡n 100 file log).
- **Lá»™ trÃ¬nh Refactor codebase:** Thá»‘ng nháº¥t káº¿ hoáº¡ch Ä‘Ã³ng gÃ³i thÆ° viá»‡n `kisordoc/`, FastAPI endpoints, vÃ  script khá»Ÿi cháº¡y song song `runner.py`.

## [2.2.0] - 2026-07-31

### Added
- **Quy trÃ¬nh Láº·p (Repeat Type Options):** Há»— trá»£ cháº¡y hÃ ng loáº¡t nhiá»u dÃ²ng dá»¯ liá»‡u cho 1 file template thÃ´ng qua bá»™ nháº­n diá»‡n `Type` = `Repeat` trong sheet `Options` (VÃ­ dá»¥: xuáº¥t cam káº¿t cho tá»«ng thÃ nh viÃªn cá»§a Tá»• chuyÃªn gia/Tá»• tháº©m Ä‘á»‹nh).
- **Bá»™ chá»n nhÃ³m & thÃ nh viÃªn Ä‘á»™ng:** TÃ­ch há»£p radio chá»n nhÃ³m ("Tá»• chuyÃªn gia" / "Tá»• tháº©m Ä‘á»‹nh") vÃ  load danh sÃ¡ch thÃ nh viÃªn Ä‘á»™ng tá»« sheet `S.TCGTTD` cá»§a gÃ³i tháº§u lÃªn checkbox Ä‘á»ƒ ngÆ°á»i dÃ¹ng chá»n ngÆ°á»i cáº§n xuáº¥t tÃ i liá»‡u.
- **LiÃªn káº¿t Ä‘á»™ng báº±ng Há» tÃªn:** Tá»± Ä‘á»™ng káº¿t ná»‘i dá»¯ liá»‡u chi tiáº¿t cá»§a thÃ nh viÃªn trong sheet `S.TCGTTD` cá»§a gÃ³i tháº§u vá»›i báº£ng dá»¯ liá»‡u dÃ¹ng chung `TCGTTD` báº±ng so khá»›p Há» tÃªn, sau Ä‘Ã³ gÃ¡n `GoiThau_ID` Ä‘á»™ng Ä‘á»ƒ thá»±c hiá»‡n phÃ©p Join cá»§a há»‡ thá»‘ng.

### Changed
- **Sá»­a giÃ¡ trá»‹ khÃ³a máº·c Ä‘á»‹nh:** Thay Ä‘á»•i giÃ¡ trá»‹ fallback máº·c Ä‘á»‹nh cá»§a `key_id` tá»« `"GoiThau_ID"` thÃ nh `"ID"` giÃºp há»‡ thá»‘ng linh hoáº¡t hÆ¡n khi cáº¥u hÃ¬nh.

## [2.1.0] - 2026-07-31

### Added
- **LiÃªn káº¿t báº£ng (Join Sheets):** Há»— trá»£ liÃªn káº¿t 2 báº£ng báº±ng kÃ½ hiá»‡u rÃºt gá»n (`<*`, `*>`, `*`, `<*>`) vÃ  trÃªn 3 báº£ng báº±ng cÃ¢u lá»‡nh SQL trá»±c tiáº¿p (`SELECT ...`).
- **Gá»™p sheet trÃ¹ng tÃªn tá»± Ä‘á»™ng:** Tá»± Ä‘á»™ng gá»™p dÃ²ng dá»¯ liá»‡u tá»« cÃ¡c file Excel khÃ¡c nhau khi phÃ¡t hiá»‡n cÃ³ sheet trÃ¹ng tÃªn (VÃ­ dá»¥: `Tables` trong `Tables.xlsx` vÃ  `DanhMuc-MSSC.xlsx`).
- **Nguá»“n file Excel Ä‘á»™ng cho Tables:** ThÃªm cá»™t `File` trong sheet `Tables` Ä‘á»ƒ tÃ¹y biáº¿n file nguá»“n copy báº£ng (VÃ­ dá»¥: `S.Oto.xlsx`).
- **NhÃ¢n báº£n báº£ng biá»ƒu tá»± Ä‘á»™ng:** Tá»± Ä‘á»™ng copy láº·p láº¡i báº£ng dá»¯ liá»‡u cuá»‘i cÃ¹ng khi sá»‘ placeholder trong Word nhiá»u hÆ¡n dÃ²ng cáº¥u hÃ¬nh Excel.
- **Cáº£nh bÃ¡o trÃ¹ng tÃªn cá»™t (Column Collision Warning):** Hiá»ƒn thá»‹ cáº£nh bÃ¡o trá»±c quan trÃªn log UI khi cÃ¡c cá»™t bá»‹ trÃ¹ng tÃªn trong quÃ¡ trÃ¬nh Join.

### Fixed
- Kháº¯c phá»¥c lá»—i so khá»›p tÃªn file chá»©a sá»‘ tháº­p phÃ¢n (VÃ­ dá»¥: `9.1 BC tham dinh ...`) khi chÃ¨n báº£ng biá»ƒu.
- Sá»­a lá»—i gá»™p Ã´ (merge cell) khi remap chá»‰ sá»‘ dÃ²ng master row bá»‹ sai tá»a Ä‘á»™ trong `table_copier.py`.
- Kháº¯c phá»¥c lá»—i crash validation cá»§a Gradio (`choices=[]`) khi chuyá»ƒn Ä‘á»•i quy trÃ¬nh hoáº·c gÃ³i tháº§u trÃªn UI.
- Tá»‘i Æ°u hÃ³a bá» qua cÃ¡c file `S.*` lÃºc khá»Ÿi Ä‘á»™ng giÃºp tiáº¿t kiá»‡m 90% dung lÆ°á»£ng RAM vÃ  tÄƒng tá»‘c app.

## [2.0.1] - 2026-07-31

### Fixed
- Kháº¯c phá»¥c hoÃ n toÃ n lá»—i cáº£nh bÃ¡o thiáº¿u dá»¯ liá»‡u (`Warning: Placeholder ... khÃ´ng cÃ³ data`) báº±ng cÃ¡ch tá»± Ä‘á»™ng Ã¡nh xáº¡ Ä‘uÃ´i `.Date` thÃ nh háº­u tá»‘ `_Date` trong cáº£ logic náº¡p context (`clean_config_key`) cá»§a `main.py` vÃ  script `migrate_modifiers.py`.
- Sá»­a lá»—i xung Ä‘á»™t (collision) ghi Ä‘Ã¨ giá»¯a Sá»‘ quyáº¿t Ä‘á»‹nh (vÃ­ dá»¥: `KHLCNT_QD`) vÃ  NgÃ y quyáº¿t Ä‘á»‹nh (vÃ­ dá»¥: `KHLCNT_QD_Date`).

## [2.0.0] - 2026-07-30

### Added
- **Cáº¥u hÃ¬nh Option Ä‘á»™ng (Dynamic Options):** Há»— trá»£ khai bÃ¡o cÃ¡c cá»™t `Sheet`, `Show`, `KeyId` trong sheet `Options` Ä‘á»ƒ tá»± Ä‘á»™ng hÃ³a Ä‘á»‹nh dáº¡ng nhÃ£n hiá»ƒn thá»‹ vÃ  tÃªn sheet nguá»“n dá»¯ liá»‡u chÃ­nh.
- **Lá»c Ä‘iá»u kiá»‡n Ä‘á»™ng (Dynamic Conditions):** Bá»• sung cá»™t `Condition` trong sheet `Workflow` há»— trá»£ cÃº phÃ¡p ngoáº·c nhá»n `{TÃªn cá»™t/TÃªn biáº¿n}` vÃ  tá»± Ä‘á»™ng parse chuá»—i sá»‘ Excel (vÃ­ dá»¥: `150.000.000` -> `150000000`) khi so sÃ¡nh logic trong Python `eval`.
- **PhÃ¢n vÃ¹ng Config theo Option (Config Range):** Bá»• sung cá»™t `Config` trong sheet `Options` cho phÃ©p tÃ¡ch biá»‡t cÃ¡c vÃ¹ng Ã¡nh xáº¡ trong sheet Config (vÃ­ dá»¥: `2-97`, `99-253`) trÃ¡nh xung Ä‘á»™t dá»¯ liá»‡u.
- **Táº­p lá»‡nh di chuyá»ƒn nÃ¢ng cao:** ThÃªm script `migrate_modifiers.py` há»— trá»£ nÃ¢ng cáº¥p Ä‘á»“ng bá»™ toÃ n bá»™ modifier (date, day, month, year, number, chu/text) sang cÃº phÃ¡p bá»™ lá»c Jinja2 `|` chuáº©n xÃ¡c.

### Changed
- Cáº£i tiáº¿n hiá»‡u nÄƒng náº¡p file Excel trong DuckDB dataset báº±ng cÆ¡ cháº¿ cache `query_rows` theo range dÃ²ng, giÃºp chá»‰ má»Ÿ file 1 láº§n khi nhiá»u Option dÃ¹ng chung phÃ¢n vÃ¹ng.
- Tá»± Ä‘á»™ng di chuyá»ƒn toÃ n bá»™ file `.bak.docx` sang thÆ° má»¥c con `bak/` tÆ°Æ¡ng á»©ng cá»§a tá»«ng thÆ° má»¥c quy trÃ¬nh.

## [1.9.0] - 2026-07-29

### Added
- Cáº¥u hÃ¬nh há»‡ sá»‘ chuyá»ƒn Ä‘á»•i Ä‘á»™ rá»™ng cá»™t tá»« Excel sang Word báº±ng biáº¿n mÃ´i trÆ°á»ng EXCEL_TO_WORD_WIDTH_FACTOR (máº·c Ä‘á»‹nh = 90) trong file .env vÃ  AppConfig.
- TÃ­ch há»£p 6 Ä‘á» xuáº¥t tÃ­nh nÄƒng nÃ¢ng cao (F1 - F6) vÃ o tÃ i liá»‡u PRD chÃ­nh thá»©c [PRD-WordBatchProcessor.md](docs/PRD-WordBatchProcessor.md).
- ThÃªm **Phase 8** vÃ  má»¥c **CÃ¡c lá»—i cáº§n sá»­a sau** vÃ o tÃ i liá»‡u [CHECKLIST-KisorDoc-AI.md](docs/CHECKLIST-KisorDoc-AI.md).

### Changed
- Cáº­p nháº­t hÃ m un_batch trong main.py Ä‘á»ƒ Ä‘á»‹nh dáº¡ng log káº¿t quáº£ xuá»‘ng dÃ²ng (join báº±ng \n) giÃºp hiá»ƒn thá»‹ rÃµ rÃ ng trÃªn giao diá»‡n Gradio Textbox thay vÃ¬ hiá»ƒn thá»‹ dáº¡ng danh sÃ¡ch thÃ´.
- Há»£p nháº¥t tÃ i liá»‡u PRD báº±ng cÃ¡ch gá»™p file PRD-Enhancements.md vÃ o PRD-WordBatchProcessor.md vÃ  xÃ³a file PRD-Enhancements.md cÅ©.

### Fixed
- Sá»­a lá»—i lá»‡ch tháº» XML (Opening and ending tag mismatch: body line 2 and tc) gÃ¢y há»ng file Word báº±ng cÃ¡ch cáº­p nháº­t regex di chuyá»ƒn template an toÃ n hÆ¡n, khÃ´ng quÃ©t xuyÃªn qua cÃ¡c tag cáº¥u trÃºc XML quan trá»ng (p, 	c, 	r, 	bl).
- KhÃ´i phá»¥c vÃ  di chuyá»ƒn thÃ nh cÃ´ng toÃ n bá»™ 36 file template gá»‘c sang cáº¥u trÃºc template {{}} má»›i an toÃ n.
- Sá»­a lá»—i tÃ¬m kiáº¿m sheet Excel khi tÃªn sheet chá»©a khoáº£ng tráº¯ng ngoÃ i mong muá»‘n (vÃ­ dá»¥: ' S.DoDa' thay vÃ¬ 'S.DoDa') báº±ng cÃ¡ch so khá»›p strip whitespace.
- Sá»­a lá»—i crash khi Ä‘á»‹nh dáº¡ng ngÃ y do dá»¯ liá»‡u chá»©a giÃ¡ trá»‹ NaT (Not a Time) trong Pandas.
- Cáº£i thiá»‡n nÃºt má»Ÿ thÆ° má»¥c output sá»­ dá»¥ng subprocess.Popen Ä‘Ã¡ng tin cáº­y hÆ¡n trÃªn Windows tá»« Gradio background thread.

