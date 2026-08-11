Cách deploy vào repo:

kisorlib/
  text_migrator.py   ← core logic (load_mapping, migrate_*, format_summary)
  text_reporter.py   ← HTML + Excel report (tách riêng để không kéo openpyxl styles vào core)
  migrator.py        ← giữ nguyên (không đụng)

migrate_text_to_placeholders.py  ← CLI thin wrapper (để ngang root, cạnh migrate_templates.py)

So sánh với migrator.py — interface đồng nhất:

	migrator.py	text_migrator.py
File-level	migrate_file(path, ...)	text_migrate_file(path, mapping, ...)
Folder-level	migrate_folder(folder, ...)	text_migrate_folder(folder, mapping, ...)
Summary	format_summary(results)	text_format_summary(results)
Progress cb	on_progress: OnProgress	on_progress: OnProgress (cùng dict schema)
Result type	FileResult	TextFileResult (thêm warnings)

Khi cần add tab "4. Migrate Text" vào app.py, chỉ cần gọi load_mapping() + text_migrate_folder() + text_format_summary() — cùng pattern với tab migrate hiện tại.