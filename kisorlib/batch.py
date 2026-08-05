"""
kisorlib/batch.py — Refactor v3.2
──────────────────────────────────
Async wrapper + Progress + IncrementalRunLogger cho Gradio UI.

KHÔNG còn logic merge/copy/retry nội bộ:
tất cả đi qua generator.generate_many / generator.generate_one_repeat.
"""

import re
import time
from pathlib import Path
from datetime import datetime
from typing import Callable, Optional

from .file_utils import clear_output_folder, copy_templates_to_output
from .utils import (
    _str,
    clean_config_key,
    safe_format,
    resolve_sheet_query,
    _parse_repeat_sheet_config,
    _parse_repeat_key_id,
)
from .app_helpers import make_nested_dict
from .service import KisorService
from .generator import (
    build_context,
    generate_many,
    generate_one_repeat,
    FileResult,
)


# ──────────────────────────────────────────────
# IncrementalRunLogger (giữ nguyên)
# ──────────────────────────────────────────────

class IncrementalRunLogger:
    def __init__(self, config, goi_thau_id, option, mode="run"):
        self.config         = config
        self.goi_thau_id    = goi_thau_id
        self.option         = option
        self.mode           = mode
        self.log_dir        = Path(config.ProjectPath) / "logs"
        self.log_dir.mkdir(exist_ok=True)
        ts                  = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        suffix              = f"_{mode}" if mode != "run" else ""
        self.filepath       = self.log_dir / f"{ts}_{goi_thau_id}_{option}{suffix}.log"
        self.start_time     = time.time()
        self.ok_count       = 0
        self.error_count    = 0
        self.warning_count  = 0
        self._fh            = open(self.filepath, "w", encoding="utf-8-sig")

    def write_header(self, option_desc, package_desc, total_files):
        time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        mode_str = "Chạy thật" if self.mode == "run" else ("Retry" if self.mode == "retry" else "Dry-run")
        self._fh.write(
            f"=====================================\n"
            f"KisorDoc – Run Log\n"
            f"=====================================\n"
            f"Thời gian     : {time_str}\n"
            f"Option        : {option_desc}\n"
            f"Gói thầu      : {package_desc}\n"
            f"Chế độ        : {mode_str}\n"
            f"Tổng file     : {total_files}\n"
            f"=====================================\n\n"
        )
        self._fh.flush()

    def log_event(self, emoji, name, extra_info=None):
        time_str = datetime.now().strftime("%H:%M:%S")
        line = f"[{time_str}] {emoji} {name}\n"
        if extra_info:
            line += f"           {extra_info}\n"
        self._fh.write(line)
        self._fh.flush()

    def write_footer(self):
        elapsed = time.time() - self.start_time
        self._fh.write(
            f"\n=====================================\n"
            f"Kết quả: {self.ok_count} thành công / "
            f"{self.error_count} lỗi / {self.warning_count} warning\n"
            f"Thời gian chạy: {elapsed:.1f} giây\n"
            f"=====================================\n"
        )
        self._fh.close()

    # ── record helpers ──

    def record_result(self, results: list, file_result: FileResult, display_name: str):
        """Ghi kết quả từ FileResult vào logger + results list."""
        name = file_result.template_name
        if file_result.success:
            if file_result.warnings:
                warn_msg = "; ".join(file_result.warnings)
                results.append(f"⚠️ {name} → {display_name}\n   → {warn_msg}")
                self.warning_count += 1
                self.log_event("⚠️", f"{name} → {display_name}", warn_msg)
            else:
                results.append(f"✅ {name} → {display_name}")
                self.ok_count += 1
                self.log_event("✅", f"{name} → {display_name}")
        else:
            err_msg = file_result.error or "Lỗi không xác định"
            is_locked = "đang mở" in err_msg or "bị khóa" in err_msg or "PermissionError" in err_msg
            if is_locked:
                results.append(f"🔒 {name}: {err_msg}")
                self.log_event("🔒", name, f"Lỗi: PermissionError – {err_msg}")
            else:
                results.append(f"❌ {name}: {err_msg}")
                self.log_event("❌", name, f"Lỗi: {err_msg}")
            self.error_count += 1

    # Backward-compat (dùng bởi code cũ còn sót)
    def record_ok(self, results, name, display_name):
        results.append(f"✅ {name} → {display_name}")
        self.ok_count += 1
        self.log_event("✅", f"{name} → {display_name}")

    def record_ok_with_warning(self, results, name, display_name, warn_msg):
        results.append(f"⚠️ {name} → {display_name}\n   → {warn_msg}")
        self.warning_count += 1
        self.log_event("⚠️", f"{name} → {display_name}", warn_msg)

    def record_locked(self, results, name, has_locked, failed_templates, detail=None):
        has_locked[0] = True
        failed_templates.append(name)
        msg = detail or "Lỗi ghi file (đang mở)"
        results.append(f"🔒 {name}: {msg}")
        self.error_count += 1
        self.log_event("🔒", name, f"Lỗi: PermissionError – {msg}")

    def record_error(self, results, name, err, has_other_error, failed_templates):
        has_other_error[0] = True
        failed_templates.append(name)
        msg = str(err)
        results.append(f"❌ {name}: {msg}")
        self.error_count += 1
        self.log_event("❌", name, f"Lỗi: {msg}")


# ──────────────────────────────────────────────
# Helpers nội bộ
# ──────────────────────────────────────────────

def _is_locked_error(err_msg: str) -> bool:
    return (
        "đang mở" in err_msg
        or "bị khóa" in err_msg
        or "PermissionError" in err_msg
        or "being used by another process" in err_msg
    )


def _find_danh_muc_file(config) -> Optional[Path]:
    xlsx_files = sorted(config.data_path.glob("*.xlsx"))
    return next(
        (f for f in xlsx_files if config.DanhMucFile.lower() in f.stem.lower()),
        xlsx_files[0] if xlsx_files else None,
    )


def _make_progress_cb(progress_cb, i, total, label):
    """Tạo on_progress adapter cho generator."""
    def _on_prog(event: dict):
        if progress_cb and event.get("level") in ("info", "warning", "error", "success"):
            progress_cb((i + 1) / total, event.get("message", label))
    return _on_prog


# ──────────────────────────────────────────────
# run_batch (Repeat branch)
# ──────────────────────────────────────────────

async def _run_batch_repeat(
    service: KisorService,
    option_key: str,
    package_label: str,
    selected_templates: list[str],
    group_name: str,
    goi_thau_id: str,
    left_key: str,
    right_key: str,
    progress_cb,
    retry_state: dict | None,
):
    config = service.config
    ds     = service.ds
    opt    = option_key.split(":")[0].strip() if ":" in option_key else option_key.strip()

    opt_config = service.get_option_config(option_key)
    key_id     = opt_config.get("key_id", "ID")
    show_format_full = opt_config.get("show", "")

    # Tìm template file
    templates = service.get_workflow_templates(option_key, package_label)
    matched_tpl = next((t for t in templates if _str(t.get("Name")) == group_name), None)
    if not matched_tpl:
        yield "", f"❌ Không tìm thấy template cho '{group_name}'", None
        return

    fname_raw = _str(matched_tpl.get("File", ""))
    fname     = fname_raw if fname_raw.endswith(".docx") else fname_raw + ".docx"

    if not retry_state:
        clear_output_folder(config)

    try:
        tables_rows = ds.query("SELECT * FROM Tables")
    except Exception:
        tables_rows = []

    config_rows   = service.get_config_for_option(option_key)
    danh_muc_file = _find_danh_muc_file(config)

    results          = []
    failed_templates = []
    has_locked       = [False]
    has_other_error  = [False]
    logger           = IncrementalRunLogger(config, goi_thau_id, opt, "retry" if retry_state else "run")
    logger.write_header(option_key, package_label, len(selected_templates))

    src_dir = config.template_path / opt
    src = src_dir / fname
    if not src.exists():
        possible = list(src_dir.glob(f"{fname}*"))
        if possible:
            src = possible[0]

    for i, member_name in enumerate(selected_templates):
        if progress_cb:
            progress_cb((i + 1) / len(selected_templates), f"Đang xử lý thành viên: {member_name}")

        try:
            ok = service.register_temporary_tcgttd(goi_thau_id, [member_name], group_name, key_id, option_key)
            if not ok:
                msg = f"⚠️ Bỏ qua thành viên '{member_name}': không tìm thấy trong bảng dữ liệu"
                results.append(msg)
                continue

            sql = resolve_sheet_query(opt_config.get("sheet", config.DataSheet))
            joined_rows = ds.query(sql)
            if not joined_rows:
                raise RuntimeError(f"Không thể kết nối thông tin cho thành viên {member_name}")
            member_pkg_row = joined_rows[0]

            # member col
            member_col = right_key
            if "|" in show_format_full:
                right_format = show_format_full.split("|", 1)[1]
                matches = re.findall(r"\{([^}]+)\}", right_format)
                if matches:
                    member_col = matches[0].strip()
            member_val     = _str(member_pkg_row.get(member_col, member_name))
            clean_member_k = clean_config_key(member_col)

            extra = {
                right_key:      _str(member_pkg_row.get(right_key, "")),
                clean_member_k: member_val,
                "HoTen":        member_val,
                "Ho_va_ten":    member_val,
            }
            ctx = build_context(member_pkg_row, config_rows, extra=extra)
            nested_ctx = make_nested_dict(ctx)
            nested_ctx["now"] = datetime.now()

            on_progress_cb = _make_progress_cb(progress_cb, i, len(selected_templates), member_name)

            file_result = generate_one_repeat(
                template_path  = src,
                template_name  = fname,
                member_name    = member_name,
                nested_context = nested_ctx,
                cfg            = config,
                goi_thau_id    = goi_thau_id,
                tables_rows    = tables_rows,
                danh_muc_file  = danh_muc_file,
                key_id         = left_key,
                max_retries    = 3,
                retry_delay    = 2.0,
                on_progress    = on_progress_cb,
            )

            display_name = Path(file_result.output_path).name if file_result.output_path else member_name
            logger.record_result(results, file_result, display_name)
            if not file_result.success:
                err_msg = file_result.error or ""
                if _is_locked_error(err_msg):
                    has_locked[0] = True
                else:
                    has_other_error[0] = True
                failed_templates.append(member_name)

        except Exception as e:
            has_other_error[0] = True
            failed_templates.append(member_name)
            results.append(f"❌ {member_name}: {e}")
            logger.log_event("❌", member_name, str(e))
            logger.error_count += 1

        yield "\n".join(results), f"Đang xử lý {i + 1}/{len(selected_templates)}...", None

    elapsed = time.time() - logger.start_time
    summary = f"✅ {logger.ok_count}  ❌ {logger.error_count}  /  {len(selected_templates)} người  ({elapsed:.1f}s)"
    logger.write_footer()

    retry_state_data = None
    if failed_templates:
        retry_state_data = {
            "option_key":        option_key,
            "package_label":     package_label,
            "failed_templates":  failed_templates,
            "all_locked":        has_locked[0] and not has_other_error[0],
            "group_name":        group_name,
        }
    yield "\n".join(results), summary, retry_state_data


# ──────────────────────────────────────────────
# run_batch (normal branch)
# ──────────────────────────────────────────────

async def _run_batch_normal(
    service: KisorService,
    option_key: str,
    package_label: str,
    selected_templates: list[str],
    goi_thau_id: str,
    left_key: str,
    selected_pkg: dict,
    progress_cb,
    retry_state: dict | None,
):
    config = service.config
    ds     = service.ds
    opt    = option_key.split(":")[0].strip() if ":" in option_key else option_key.strip()
    opt_config = service.get_option_config(option_key)
    key_id     = opt_config.get("key_id", "ID")

    try:
        tables_rows = ds.query("SELECT * FROM Tables")
    except Exception:
        tables_rows = []

    config_rows = service.get_config_for_option(option_key)
    ctx         = build_context(selected_pkg, config_rows)
    nested_ctx  = make_nested_dict(ctx)
    nested_ctx["now"] = datetime.now()

    danh_muc_file = _find_danh_muc_file(config)

    table_placeholder_names = {
        _str(t.get("Name", "")).strip("{} ")
        for t in tables_rows
        if _str(t.get(key_id)) == goi_thau_id
    }

    if progress_cb:
        progress_cb(0, "Bắt đầu xử lý...")
    yield "", "Bắt đầu...", None

    mode   = "retry" if retry_state else "run"
    logger = IncrementalRunLogger(config, goi_thau_id, opt, mode)

    # Chuẩn bị danh sách template file
    if retry_state:
        failed_names = retry_state.get("failed_templates", [])
        template_filenames, template_names = [], []
        for r in service.get_workflow_templates(option_key, package_label):
            if r.get("Name", "") in failed_names:
                fn = _str(r.get("File", ""))
                fn = fn if fn.endswith(".docx") else fn + ".docx"
                template_filenames.append(fn)
                template_names.append(r.get("Name", ""))
        from .file_utils import copy_templates_to_output as _cto
        src_dir = config.template_path / opt
        copied  = []
        names_actual = []
        for fn, tn in zip(template_filenames, template_names):
            src = src_dir / fn
            if not src.exists():
                possible = list(src_dir.glob(f"{fn}*"))
                src = possible[0] if possible else None
            if src:
                import shutil as _sh
                dst = config.output_path / src.name
                try:
                    _sh.copy2(str(src), str(dst))
                    copied.append(dst)
                    names_actual.append(tn)
                except Exception:
                    copied.append(dst)
                    names_actual.append(tn)
        template_names = names_actual
    else:
        clear_output_folder(config)
        template_filenames, template_names = [], []
        for r in service.get_workflow_templates(option_key, package_label):
            if r.get("Name", "") in selected_templates:
                fn = _str(r.get("File", ""))
                fn = fn if fn.endswith(".docx") else fn + ".docx"
                template_filenames.append(fn)
                template_names.append(r.get("Name", ""))
        copied = copy_templates_to_output(config, opt, template_filenames)

    total = len(copied)
    logger.write_header(option_key, package_label, total)

    results          = []
    failed_templates = []
    has_locked       = [False]
    has_other_error  = [False]

    # Adapter on_progress cho generator → emit event → update Gradio progress + log
    # Generator sẽ gọi on_progress per event; ta cần per-file counters
    file_results = generate_many(
        template_paths          = list(zip(copied, template_names)),
        nested_context          = nested_ctx,
        cfg                     = config,
        goi_thau_id             = goi_thau_id,
        tables_rows             = tables_rows,
        danh_muc_file           = danh_muc_file,
        key_id                  = left_key,
        table_placeholder_names = table_placeholder_names,
        dry_run                 = False,
        max_retries             = 3,
        retry_delay             = 2.0,
        on_progress             = None,  # chúng ta yield từng file bên dưới
    )

    for i, file_result in enumerate(file_results):
        if progress_cb:
            progress_cb((i + 1) / total, f"Đang xử lý {i + 1}/{total}...")

        display_name = Path(file_result.output_path).name if file_result.output_path else file_result.template_name
        logger.record_result(results, file_result, display_name)

        if not file_result.success:
            err_msg = file_result.error or ""
            if _is_locked_error(err_msg):
                has_locked[0] = True
            else:
                has_other_error[0] = True
            failed_templates.append(file_result.template_name)

        yield "\n".join(results), f"Đang xử lý {i + 1}/{total}...", None

    elapsed = time.time() - logger.start_time
    summary = (
        f"✅ {logger.ok_count}  ⚠️ {logger.warning_count}  "
        f"❌ {logger.error_count}  /  {total} file  ({elapsed:.1f}s)"
    )
    logger.write_footer()

    retry_state_data = None
    if failed_templates:
        retry_state_data = {
            "option_key":       option_key,
            "package_label":    package_label,
            "failed_templates": failed_templates,
            "all_locked":       has_locked[0] and not has_other_error[0],
        }

    yield "\n".join(results), summary, retry_state_data


# ──────────────────────────────────────────────
# run_batch — public entry point
# ──────────────────────────────────────────────

async def run_batch(
    service: KisorService,
    option_key: str,
    package_label: str,
    selected_templates: list[str],
    group_name: str = "",
    progress_cb: Optional[Callable[[float, str], None]] = None,
    retry_state: dict | None = None,
):
    config = service.config
    ds     = service.ds

    if not option_key or not option_key.strip():
        yield "", "⚠️ Vui lòng chọn quy trình", None
        return
    if not package_label or not package_label.strip():
        yield "", "⚠️ Vui lòng chọn gói thầu", None
        return
    if not selected_templates:
        yield "", "⚠️ Vui lòng chọn ít nhất 1 template hoặc thành viên", None
        return

    opt_config   = service.get_option_config(option_key)
    sheet        = opt_config.get("sheet", config.DataSheet)
    key_id       = opt_config.get("key_id", "ID")
    show_format  = opt_config.get("show", "")
    if "|" in show_format:
        show_format = show_format.split("|")[0].strip()

    left_key, right_key = _parse_repeat_key_id(key_id)

    # Query dòng package
    if opt_config.get("type") == "Repeat":
        ls, _, _ = _parse_repeat_sheet_config(opt_config)
        temp_sql = f'SELECT * FROM "{ls}"' if ls else resolve_sheet_query(sheet)
    else:
        temp_sql = resolve_sheet_query(sheet)

    temp_rows = ds.query(temp_sql)

    selected_pkg = None
    for r in temp_rows:
        if safe_format(show_format, r) == package_label:
            selected_pkg = r
            break
    if not selected_pkg:
        yield "", "❌ Không tìm thấy dòng dữ liệu tương ứng", None
        return

    goi_thau_id = _str(selected_pkg.get(left_key))

    if opt_config.get("type") == "Repeat":
        async for item in _run_batch_repeat(
            service, option_key, package_label, selected_templates,
            group_name, goi_thau_id, left_key, right_key, progress_cb, retry_state,
        ):
            yield item
    else:
        async for item in _run_batch_normal(
            service, option_key, package_label, selected_templates,
            goi_thau_id, left_key, selected_pkg, progress_cb, retry_state,
        ):
            yield item


# ──────────────────────────────────────────────
# run_retry_batch
# ──────────────────────────────────────────────

async def run_retry_batch(
    service: KisorService,
    retry_state: dict,
    progress_cb: Optional[Callable[[float, str], None]] = None,
):
    if not retry_state or not retry_state.get("failed_templates"):
        yield "⚠️ Không có file lỗi nào để chạy lại", "⚠️ Không có file lỗi", None
        return
    option_key        = retry_state["option_key"]
    package_label     = retry_state["package_label"]
    failed_templates  = retry_state["failed_templates"]
    group_name        = retry_state.get("group_name", "")
    async for log, status, new_state in run_batch(
        service            = service,
        option_key         = option_key,
        package_label      = package_label,
        selected_templates = failed_templates,
        group_name         = group_name,
        progress_cb        = progress_cb,
        retry_state        = retry_state,
    ):
        yield log, status, new_state
