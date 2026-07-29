# Changelog – KisorDoc-AI

## ver1.10 (2026-07-29)

### F3 – Retry cho file lỗi ✅
- Hiển thị nút "🔄 Chạy lại file lỗi" khi có file ❌ trong kết quả
- Nút ẩn khi tất cả file thành công
- Retry chỉ xử lý lại các template bị lỗi, giữ nguyên các file đã OK
- Lưu trạng thái retry trong `gr.State`, reset khi thay đổi Option/Gói thầu
- Hỗ trợ retry nhiều lần liên tiếp

### F6 – Xử lý file đang mở (File Locked) ✅
- Phân biệt `PermissionError` khỏi các exception khác
- Tự động retry 3 lần, mỗi lần chờ 2 giây với thông báo trạng thái
- File locked đánh dấu `🔒` (khác `❌`) trong log
- Nếu tất cả lỗi đều là `🔒`, nút retry đổi thành "🔄 Chạy lại (đã đóng file chưa?)"
- Áp dụng retry cho cả bước copy template

### Bug fixes
- Fix: `on_retry_click` trả 2 thay vì 3 values → Gradio error
- Fix: Retry `copied` vs `template_names` lệch nhau khi file bị skip
- Fix: `do_merge` closure late binding → thêm default arg
- Fix: Tab 2 nằm ngoài `gr.Tabs()` → không chuyển tab được khi rerun
- Fix: `run_retry_batch` yield `gr.update()` vào `gr.State` → yield raw value
- Fix: Parse tpl_name từ log string bị sai nếu tên có dấu `:` → track trực tiếp trong loop
- Fix: `get_retry_label` + `enable_retry` double-update → gộp vào 1 bước
- Fix: Layout `package_radio` nằm ngoài `Column` → đưa vào đúng vị trí
- Fix: `import shutil` trong function → move lên top-level
- Fix: `retry_btn variant="warning"` không hợp lệ → đổi thành `"stop"`
- Fix: `status_text` không reset khi rerun → thêm vào outputs
- Fix: Nút Chạy/Retry không disable khi đang xử lý → thêm toggle

### Config
- Thêm `FileRetryDelay` và `FileMaxRetries` vào `.env` (có thể config)
- `write_with_retry()` wrapper cho retry logic

---

## ver1.9
- Force kill WINWORD.EXE trước khi clearing folder
- `show_progress="full"` cho progress bar
- Excel*143 column width conversion

## ver1.8
- File access retry logic
- Proportional column width scaling

## ver1.7
- Lenient partial file name matching cho tables
- Column width conversion improvements
- XML error handling

## ver1.6
- Fixed table copy order (tables before mail_merge)
- Added progress bar
- Progress placeholder names in context

## ver1.5
- Auto-load templates on package selection
- Table copy error handling
- NaT datetime handling fix

## ver1.4
- Merged Tab 1&2 into 2-column layout
- Fixed open folder button
- Fixed rerun Radio reset

## ver1.3
- 12 UI improvements (validation, package preview, better logging, textbox log, etc.)

## ver1.2
- Applied patches: filters.py, dataset.py, table_copier.py, main.py
- Added `mail_merge_safe()` to merger.py

## ver1.1
- `.env-example`, `config.py` updated
- README updated

## ver1.0
- Initial release
