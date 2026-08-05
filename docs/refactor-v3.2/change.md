Những thay đổi
File mới

kisorlib/generator.py — Core sync, single source of truth:

build_context(selected_pkg, config_rows, extra) — xây Jinja2 context phẳng, pure/sync
write_with_retry(func, max_retries, delay, on_retry) — retry file lock, chuyển từ batch
generate_one(...) — sinh 1 file: copy → merge → copy_tables → rename
generate_many(...) — lặp template, gom FileResult[], gọi on_progress callback
generate_one_repeat(...) — sinh 1 file cho 1 thành viên Repeat
OnProgress = Callable[[dict], None] — event schema chuẩn cho UI/API

kisorlib/sql_join.py — SQL join helpers tách ra (Phase B):

parse_join_expression, resolve_sheet_query, _parse_repeat_sheet_config, validate_sql_identifier

tests/test_generator.py — 25 unit test mới

File refactor

kisorlib/batch.py — Chỉ còn: async wrapper + IncrementalRunLogger + adapter on_progress cho Gradio. Không còn DocxTemplate, mail_merge_safe, hay retry logic riêng.

kisorlib/engine.py — Chỉ còn: Pydantic schema + _build_context_for_request + gọi generate_many. Không còn _process_one hay merge nội bộ.

kisorlib/utils.py — Re-export SQL helpers từ sql_join.py để backward-compat; xóa định nghĩa trùng.

Apply patch
bash
cd KisorDoc
git apply refactor-v3.2.patch
pytest tests/ -q