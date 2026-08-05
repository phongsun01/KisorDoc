# Kế hoạch Tái cấu trúc KisorDoc — Refactor v3.2

**Trạng thái:** Draft / Ready to implement  
**Cập nhật:** 05/08/2026  
**Phạm vi phiên bản mục tiêu:** 4.1.x (sau 4.0.3)  
**Liên quan:** `docs/known-issues.md`, `CHANGELOG.md`, `kisorlib/batch.py`, `kisorlib/engine.py`

---

## 1. Bối cảnh

### 1.1 Đã hoàn thành (v4.0.0 – v4.0.3)

- Tách `app.py` God-file: logic batch chuyển sang `kisorlib/batch.py`, nghiệp vụ vào `kisorlib/service.py`.
- `app.py` chỉ còn kết nối Gradio UI (~360–460 dòng).
- Bỏ global `service`; khởi tạo qua `init() -> KisorService` trong `create_ui()`.
- SQL parameter binding (`?`) + whitelist identifier (`validate_sql_identifier`).
- Unit test: `test_utils`, `test_filters`, `test_service`.
- README / `.env-example` / hardcode `api.py` đã được dọn.

### 1.2 Nợ kiến trúc còn lại

| Vấn đề | Hệ quả |
|--------|--------|
| **Dual-pipeline** | UI dùng `batch.run_batch`; API dùng `engine.generate_documents` — logic sinh file trùng / dễ lệch |
| **`utils.py` đa trách nhiệm** | Join SQL, format, AST, identifier… chung một file |
| Feature parity khó đảm bảo | Sửa merge/filter/retry phải nhớ cả hai nhánh |

Refactor v3.1 chỉ giải quyết các nợ trên; **không** mở rộng clean architecture hay thêm DI framework.

---

## 2. Định nghĩa hoàn thành (Definition of Done)

> **Xong = dual-pipeline hết.**

Sau Refactor v3.1:

1. Mọi thay đổi logic sinh file (build context, filter, merge, copy bảng, lock/retry, rename) **chỉ** sửa trong `kisorlib/generator.py`.
2. `kisorlib/batch.py` (UI) và `kisorlib/engine.py` (API) **không** còn block render/merge riêng; chỉ còn validate, map tham số, progress/log adapter.
3. `pytest` xanh; smoke UI + API (generate thường + Repeat) cùng một pipeline.
4. `docs/known-issues.md` đánh dấu dual-pipeline **resolved**; CHANGELOG ghi rõ phiên bản.

---

## 3. Kiến trúc mục tiêu

```
                  ┌──────────────┐
                  │    app.py    │  Gradio UI
                  └──────┬───────┘
                         │
                  ┌──────▼───────┐
                  │   batch.py   │  Async wrapper + Progress + IncrementalRunLogger
                  └──────┬───────┘
                         │
┌──────────────┐  ┌──────▼───────┐
│    api.py    ├──►  engine.py   │  Pydantic I/O + job/log adapter
└──────────────┘  └──────┬───────┘
                         │
                  ┌──────▼───────┐
                  │ generator.py │  Core sync: sinh văn bản (single source of truth)
                  └──────┬───────┘
                         │
          ┌──────────────┼──────────────┐
          ▼              ▼              ▼
     service.py     merger /       file_utils /
     (context,      table_copier   write_with_retry
      repeat reg)
```

**Nguyên tắc:**

- Core **sync**, không import Gradio.
- UI/API chỉ là adapter.
- `KisorService` map dữ liệu (label → id, option config, register bảng tạm); không chứa logic ghi file docx.

---

## 4. Bước A — Gộp pipeline (ưu tiên cao)

### 4.1 Phân ranh trách nhiệm (core sync)

| Thành phần | Trách nhiệm | Ghi chú |
|------------|------------|---------|
| `build_context(...)` | Jinja context + nested dict + gắn filters | Pure/sync; input từ service/ds |
| `generate_one(template_path, context, tables_cfg, output_path, ...)` | Merge 1 file + copy bảng + rename | **Không** biết Repeat |
| `generate_many(...)` | Lặp template, gọi `generate_one`, gom `FileResult`, gọi `on_progress` | Sync |
| Repeat orchestration | Register bảng tạm + lặp theo thành viên → gọi `generate_many` / `generate_one` | Tầng cao (`generator` facade hoặc `service` + generator); **không** nhét vào `generate_one` |

### 4.2 Progress callback

`generator` không phụ thuộc `IncrementalRunLogger` hay `gr.Progress`.

```python
OnProgress = Callable[[dict], None] | None

# event gợi ý:
# {
#   "level": "info" | "warning" | "error" | "success",
#   "message": str,
#   "step": int | None,
#   "total": int | None,
#   "template": str | None,
# }
```

| Adapter | Cách dùng callback |
|---------|-------------------|
| `batch.py` | Cập nhật `gr.Progress` + ghi `IncrementalRunLogger` |
| `engine.py` | Append log job / logging chuẩn |

### 4.3 Lock / retry (chọn A1)

- `write_with_retry` nằm trong đường ghi của core (trong `generator` hoặc gọi `file_utils` từ generator).
- Mọi output (UI và API) đều đi qua retry khi file bị Word khóa.

### 4.4 Wrapper mỏng

| Module | Việc được làm | Việc không được làm |
|--------|----------------|---------------------|
| `generator.py` | Sync core sinh file | Import Gradio; tự `load_config()` |
| `batch.py` | Async/`yield` log UI; map UI state → tham số core | Tự merge/copy docx |
| `engine.py` | Validate Pydantic; map `GenerateRequest` → core; trả `GenerateResult` | Tự merge/copy docx |

### 4.5 Mapping request

```
GenerateRequest (API)  ──map──►  args generate_many(...)
UI (option, package_label, templates, group)
                       ──map──►  cùng args generate_many(...)
```

- `package_label` → `package_id` / row chi tiết: **`KisorService`**.
- `generator` chỉ nhận id/context/đường dẫn đã resolve.

### 4.6 Nguyên tắc testability

- `generator` nhận `KisorService` **hoặc** `(config, ds)` qua tham số.
- Không gọi `load_config()` bên trong core.
- Unit test mock service/ds in-memory (pattern `test_service.py`).

---

## 5. Bước B — Tách nhẹ `utils`

### 5.1 `kisorlib/sql_join.py` (mới)

Chuyển:

- `_OP_MAP`, `_JOIN_RE`
- `parse_join_expression`, `resolve_sheet_query`
- `_parse_repeat_sheet_config`
- `validate_sql_identifier`

### 5.2 `kisorlib/utils.py` (giữ)

- `_str`, `safe_format`, `clean_config_key`
- `_parse_price`, `_parse_row_range`
- `_safe_eval_condition`
- `_parse_repeat_key_id`

### 5.3 Import

- `service` / `batch` / `engine` / `generator` import SQL từ `sql_join` trực tiếp.
- **Không** re-export từ `utils` sang `sql_join` (tránh hai đường import).

Làm **sau** khi pipeline đã gộp ổn định (phase B trong roadmap), trừ khi conflict import buộc làm sớm hơn trên branch riêng.

---

## 6. Bước C — Dependency injection

- Giữ khởi tạo thủ công: `KisorService(config, ds)` trong `app` / `engine`.
- **Không** thêm DI container / framework.
- Core chỉ nhận dependency qua tham số (xem §4.6).

---

## 7. Roadmap thực hiện

| Phase | Công việc | Rủi ro | Done when |
|-------|-----------|--------|-----------|
| **A0** | Thêm `generator.generate_one` (copy logic merge+copy từ `batch`); chưa xóa code cũ | Thấp | Unit test sinh 1 file qua generator |
| **A1** | `generate_many` + `on_progress`; `batch.run_batch` gọi core | Thấp | UI batch chạy qua `generate_many` |
| **A2** | `engine.generate_documents` gọi core; xóa merge trùng trong engine | Trung bình | API generate/dry-run khớp hành vi UI (cùng context cơ bản) |
| **A3** | Repeat orchestration tầng cao; UI + API cùng path | Trung bình–cao | Repeat smoke test UI và API |
| **Parity** | Checklist §8; mới được xóa dead code trong batch/engine | — | Không regress tính năng đã liệt kê |
| **B** | Tách `sql_join.py` + sửa import | Thấp | `pytest` xanh |
| **Docs** | `known-issues` + CHANGELOG | Thấp | Dual-pipeline đánh dấu resolved |

**Khuyến nghị:** một người làm tuần tự A0→A3→Parity rồi B; tránh song song A và B trên cùng branch.

---

## 8. Feature parity checklist (bắt buộc trước khi xóa code cũ)

Đối chiếu UI (batch cũ) với core mới:

- [ ] Mail merge Jinja2 + toàn bộ filters (`date`, `date_long`, `number`, `num2text`, …)
- [ ] Nested context (dấu chấm)
- [ ] Copy bảng Excel→Word (merge cell, style, file nguồn động)
- [ ] Nhân bản bảng khi thiếu dòng config
- [ ] `write_with_retry` / trạng thái file khóa
- [ ] Cảnh báo missing placeholder
- [ ] Clear output / copy template / rename output
- [ ] Retry chỉ file lỗi (không xóa folder thành công)
- [ ] Incremental log (UI) tương đương về thông tin sự kiện
- [ ] Config range theo Option
- [ ] Condition trên Workflow (`check_condition` / AST)
- [ ] Repeat: chọn nhóm, thành viên, register tạm, không stale data
- [ ] Dry-run / preview (nếu API/UI còn hỗ trợ)
- [ ] API: `GenerateResult` đủ success/fail/paths/log tối thiểu

---

## 9. Rủi ro và giảm thiểu

| # | Rủi ro | Giảm thiểu |
|---|--------|------------|
| 1 | Lệch behavior / mất tính năng nhỏ | Checklist §8; giữ code cũ đến khi parity đạt |
| 2 | DuckDB session / bảng tạm Repeat xung đột | `generator` không quản session; `service` register/cleanup quanh lần gọi generate |
| 3 | Core bị biến thành async | Core luôn sync; chỉ `batch` async/`yield` |
| 4 | API worker gặp file lock | `write_with_retry` trên mọi đường ghi (A1) |
| 5 | API/UI map khác nhau → khác context | Cùng hàm map qua `KisorService`; test so sánh context key set |

---

## 10. Phạm vi không làm (Non-goals)

- Không viết lại Gradio UI.
- Không thêm DI framework, message queue, hay microservice.
- Không đổi format template Word / schema Excel nghiệp vụ.
- Không tách `utils` thành nhiều file ngoài `sql_join.py` trong đợt này.
- Không bắt buộc API hỗ trợ full Repeat UI parity trong A2 (A3 mới là mốc Repeat).

---

## 11. Tiêu chí chấp nhận (Acceptance)

1. Không còn logic merge/copy/retry song song trong `batch.py` và `engine.py`.
2. Sửa một bug merge trong `generator.py` là đủ cho cả UI và API.
3. `pytest` (utils, filters, service, và test generator tối thiểu) xanh.
4. Smoke: 1 option thường + 1 option Repeat trên UI; 1 lần `POST /generate` (và dry-run nếu có) qua API.
5. Tài liệu: DoD §2 thỏa; `known-issues.md` cập nhật; CHANGELOG có mục Refactor v3.1 / dual-pipeline.

---

## 12. Tài liệu liên quan cần cập nhật khi xong

| File | Nội dung cập nhật |
|------|-------------------|
| `docs/known-issues.md` | Đánh dấu dual-pipeline resolved; ghi module `generator.py` |
| `CHANGELOG.md` | Mục Added/Changed theo phase merge |
| `README.md` | (Tuỳ chọn) một dòng về core `generator` trong cấu trúc mã |
| `docs/PRD-*.md` | Chỉ khi PRD còn mô tả pipeline cũ |

---

## 13. Tóm tắt quyết định thiết kế

| Chủ đề | Quyết định |
|--------|------------|
| Single core | `kisorlib/generator.py` |
| Sync vs async | Core sync; async chỉ ở `batch` |
| Progress | `on_progress: Callable[[dict], None] \| None` |
| File lock | A1 — retry trong đường ghi core |
| Repeat | Orchestration ngoài `generate_one` |
| Map label→id | `KisorService` |
| SQL join helpers | `sql_join.py` (phase B) |
| DI | Thủ công, không framework |