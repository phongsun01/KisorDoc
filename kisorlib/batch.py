import time
import shutil
import re
import errno
import pandas as pd
import jinja2
from pathlib import Path
from datetime import datetime
from typing import Callable, Optional
from docxtpl import DocxTemplate

from .merger import mail_merge_safe
from .table_copier import copy_tables_for_file
from .file_utils import clear_output_folder, copy_templates_to_output, rename_output
from .app_helpers import make_nested_dict
from .utils import (
    _str,
    clean_config_key,
    safe_format,
    resolve_sheet_query,
    _parse_repeat_sheet_config,
    _parse_repeat_key_id,
)
from .service import KisorService
from .filters import (
    filter_date, filter_date_long, filter_number, filter_num2text,
    filter_day, filter_month, filter_year, filter_add_days,
    filter_add_months, filter_date_diff, filter_quarter,
    filter_weekday, filter_date_text
)


def write_with_retry(func, max_retries=3, delay=2.0, yield_fn=None) -> tuple[bool, str]:
    """
    Thực thi func() có retry khi gặp PermissionError (file đang bị khóa).
    Luôn trả về (bool, str) để caller dùng pattern: ok, err = write_with_retry(...)
    """
    for attempt in range(1, max_retries + 1):
        try:
            result = func()
            if isinstance(result, tuple) and len(result) == 2:
                return bool(result[0]), str(result[1])
            return True, ""
        except PermissionError as e:
            if "being used by another process" in str(e) or e.errno == errno.EACCES:
                if attempt == max_retries:
                    raise
                msg = f"🔒 File bị khóa – thử lại lần {attempt}/{max_retries} sau {delay:.0f}s..."
                if yield_fn:
                    yield_fn(msg)
                time.sleep(delay)
            else:
                raise


class IncrementalRunLogger:
    def __init__(self, config, goi_thau_id, option, mode="run"):
        self.config = config
        self.goi_thau_id = goi_thau_id
        self.option = option
        self.mode = mode
        
        self.log_dir = Path(config.ProjectPath) / "logs"
        self.log_dir.mkdir(exist_ok=True)
        
        ts = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        suffix = f"_{mode}" if mode != "run" else ""
        self.filepath = self.log_dir / f"{ts}_{goi_thau_id}_{option}{suffix}.log"
        self.start_time = time.time()
        self.ok_count = 0
        self.error_count = 0
        self.warning_count = 0
        self._fh = open(self.filepath, "w", encoding="utf-8-sig")

    def write_header(self, option_desc, package_desc, total_files):
        time_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        mode_str = "Chạy thật" if self.mode == "run" else ("Retry" if self.mode == "retry" else "Dry-run")
        header = f"""=====================================
KisorDoc – Run Log
=====================================
Thời gian     : {time_str}
Option        : {option_desc}
Gói thầu      : {package_desc}
Chế độ        : {mode_str}
Tổng file     : {total_files}
=====================================

"""
        self._fh.write(header)
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
        footer = f"""
=====================================
Kết quả: {self.ok_count} thành công / {self.error_count} lỗi / {self.warning_count} warning
Thời gian chạy: {elapsed:.1f} giây
=====================================
"""
        self._fh.write(footer)
        self._fh.close()

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


async def run_batch(service: KisorService, option_key: str, package_label: str, selected_templates: list[str],
                    group_name: str = "", progress_cb: Optional[Callable[[float, str], None]] = None,
                    retry_state: dict | None = None):

    config = service.config
    ds = service.ds

    if not option_key or not option_key.strip():
        yield "", "⚠️ Vui lòng chọn quy trình", None
        return
    if not package_label or not package_label.strip():
        yield "", "⚠️ Vui lòng chọn gói thầu", None
        return
    if not selected_templates or len(selected_templates) == 0:
        yield "", "⚠️ Vui lòng chọn ít nhất 1 template hoặc thành viên", None
        return

    opt = option_key.split(":")[0].strip() if ":" in option_key else option_key.strip()

    opt_config = service.get_option_config(option_key)
    sheet = opt_config.get("sheet", config.DataSheet)
    key_id = opt_config.get("key_id", "ID")
    show_format = opt_config.get("show", "")
    show_format_full = show_format
    if "|" in show_format:
        show_format = show_format.split("|")[0].strip()

    # Query initial package to get goi_thau_id
    if opt_config.get("type") == "Repeat":
        ls, _, _ = _parse_repeat_sheet_config(opt_config)
        temp_sql = f'SELECT * FROM "{ls}"' if ls else resolve_sheet_query(sheet)
    else:
        temp_sql = resolve_sheet_query(sheet)
    temp_rows = ds.query(temp_sql)
    selected_pkg_initial = None
    for r in temp_rows:
        if safe_format(show_format, r) == package_label:
            selected_pkg_initial = r
            break
    if not selected_pkg_initial:
        yield "", "❌ Không tìm thấy dòng dữ liệu tương ứng", None
        return
    left_key, right_key = _parse_repeat_key_id(key_id)
    goi_thau_id = _str(selected_pkg_initial.get(left_key))

    sql = resolve_sheet_query(sheet)

    if opt_config.get("type") == "Repeat":
        templates = service.get_workflow_templates(option_key, package_label, sheet_rows=temp_rows)
        matched_tpl = next((t for t in templates if _str(t.get("Name")) == group_name), None)
        if not matched_tpl:
            yield "", f"❌ Không tìm thấy template cho '{group_name}'", None
            return

        fname_raw = _str(matched_tpl.get("File", ""))
        fname = fname_raw if fname_raw.endswith(".docx") else fname_raw + ".docx"

        if not retry_state:
            clear_output_folder(config)

        try:
            tables_rows = ds.query("SELECT * FROM Tables")
        except Exception:
            tables_rows = []

        config_rows = service.get_config_for_option(option_key)

        xlsx_files = sorted(config.data_path.glob("*.xlsx"))
        danh_muc_file = next(
            (f for f in xlsx_files if config.DanhMucFile.lower() in f.stem.lower()),
            xlsx_files[0] if xlsx_files else None
        )

        results = []
        failed_templates = []
        has_locked = [False]
        has_other_error = [False]
        
        logger = IncrementalRunLogger(config, goi_thau_id, opt, "retry" if retry_state else "run")
        logger.write_header(option_key, package_label, len(selected_templates))

        for i, member_name in enumerate(selected_templates):
            if progress_cb:
                progress_cb((i + 1) / len(selected_templates), f"Đang xử lý thành viên: {member_name}")

            try:
                ok = service.register_temporary_tcgttd(goi_thau_id, [member_name], group_name, key_id, option_key)
                if not ok:
                    msg = f"⚠️ Bỏ qua thành viên '{member_name}': không tìm thấy trong bảng dữ liệu"
                    logger.write_item(member_name, fname, "SKIP", msg)
                    results.append((member_name, "SKIP"))
                    continue

                joined_rows = ds.query(sql)
                if not joined_rows:
                    raise RuntimeError(f"Không thể kết nối thông tin cho thành viên {member_name}")
                member_pkg_row = joined_rows[0]

                context = {}
                for r in config_rows:
                    key = _str(r.get("Key"))
                    col = _str(r.get("Value"))
                    if not key or not col:
                        continue
                    clean_key = clean_config_key(key)
                    raw_value = member_pkg_row.get(col, "")
                    try:
                        is_na = pd.isna(raw_value)
                    except (TypeError, ValueError):
                        is_na = False
                    if is_na:
                        raw_value = ""
                    elif isinstance(raw_value, datetime):
                        raw_value = raw_value.strftime("%d/%m/%Y")
                    elif raw_value is None:
                        raw_value = ""
                    context[clean_key] = str(raw_value)

                context[right_key] = _str(member_pkg_row.get(right_key, ""))

                member_col = right_key
                if "|" in show_format_full:
                    right_format = show_format_full.split("|", 1)[1]
                    matches = re.findall(r"\{([^}]+)\}", right_format)
                    if matches:
                        member_col = matches[0].strip()
                member_val = _str(member_pkg_row.get(member_col, member_name))
                clean_member_key = clean_config_key(member_col)
                context[clean_member_key] = member_val
                context["HoTen"] = member_val
                context["Ho_va_ten"] = member_val

                nested_context = make_nested_dict(context)
                nested_context["now"] = datetime.now()

                src_dir = config.template_path / opt
                src = src_dir / fname
                if not src.exists():
                    possible = list(src_dir.glob(f"{fname}*"))
                    if possible:
                        src = possible[0]
                dst = config.output_path / src.name

                def do_copy(s=src, d=dst):
                    shutil.copy2(s, d)
                write_with_retry(do_copy, max_retries=3, delay=2.0)

                def do_merge():
                    return mail_merge_safe(dst, nested_context, dst)
                ok, err = write_with_retry(do_merge, max_retries=3, delay=2.0)
                if not ok:
                    raise RuntimeError(err)

                if danh_muc_file and danh_muc_file.exists():
                    try:
                        copy_tables_for_file(dst, config, goi_thau_id, tables_rows, danh_muc_file, left_key)
                    except Exception as table_err:
                        print(f"⚠️ Lỗi copy bảng: {table_err}")

                new_filename = f"{src.stem.replace('-Template', '')}-{goi_thau_id}-{member_name}.docx"
                new_path = config.output_path / new_filename
                if dst.exists():
                    dst.rename(new_path)

                logger.record_ok(results, member_name, new_filename)

            except PermissionError as e:
                logger.record_locked(results, member_name, has_locked, failed_templates)
            except Exception as e:
                logger.record_error(results, member_name, e, has_other_error, failed_templates)

            yield "\n".join(results), f"Đang xử lý {i + 1}/{len(selected_templates)}...", None

        elapsed = time.time() - logger.start_time
        summary = f"✅ {logger.ok_count}  ❌ {logger.error_count}  /  {len(selected_templates)} người  ({elapsed:.1f}s)"
        logger.write_footer()

        retry_state_data = None
        if failed_templates:
            retry_state_data = {
                "option_key": option_key,
                "package_label": package_label,
                "failed_templates": failed_templates,
                "all_locked": has_locked[0] and not has_other_error[0],
                "group_name": group_name
            }
        yield "\n".join(results), summary, retry_state_data
        return

    sql = resolve_sheet_query(sheet)
    goi_thau_rows = ds.query(sql)
    selected_pkg = None
    for r in goi_thau_rows:
        label = safe_format(show_format, r)
        if label == package_label:
            selected_pkg = r
            break

    if not selected_pkg:
        yield "", "❌ Không tìm thấy dòng dữ liệu tương ứng", None
        return

    left_key, right_key = _parse_repeat_key_id(key_id)
    goi_thau_id = _str(selected_pkg.get(left_key))
    config_rows = service.get_config_for_option(option_key)

    try:
        tables_rows = ds.query("SELECT * FROM Tables")
    except Exception:
        tables_rows = []

    context = {}
    for r in config_rows:
        key = _str(r.get("Key"))
        col = _str(r.get("Value"))
        if not key or not col:
            continue
        clean_key = clean_config_key(key)

        raw_value = selected_pkg.get(col, "")

        try:
            is_na = pd.isna(raw_value)
        except (TypeError, ValueError):
            is_na = False
        if is_na:
            raw_value = ""
        elif isinstance(raw_value, datetime):
            raw_value = raw_value.strftime("%d/%m/%Y")
        elif raw_value is None:
            raw_value = ""

        context[clean_key] = str(raw_value)

    nested_context = make_nested_dict(context)
    nested_context["now"] = datetime.now()

    xlsx_files = sorted(config.data_path.glob("*.xlsx"))
    danh_muc_file = next(
        (f for f in xlsx_files if config.DanhMucFile.lower() in f.stem.lower()),
        xlsx_files[0] if xlsx_files else None
    )

    if progress_cb:
        progress_cb(0, "Bắt đầu xử lý...")
    yield "", "Bắt đầu...", None

    mode = "retry" if retry_state else "run"
    logger = IncrementalRunLogger(config, goi_thau_id, opt, mode)

    if retry_state:
        failed_names = retry_state.get("failed_templates", [])
        template_filenames, template_names = [], []
        for r in service.get_workflow_templates(option_key, package_label):
            if r.get("Name", "") in failed_names:
                fname_raw = _str(r.get("File", ""))
                fname = fname_raw if fname_raw.endswith(".docx") else fname_raw + ".docx"
                template_filenames.append(fname)
                template_names.append(r.get("Name", ""))
        copied = []
        template_names_actual = []
        src_dir = config.template_path / opt
        for fname, tname in zip(template_filenames, template_names):
            src = src_dir / fname
            if not src.exists():
                possible = list(src_dir.glob(f"{fname}*"))
                if possible:
                    src = possible[0]
                else:
                    continue
            dst = config.output_path / src.name
            def do_copy(s=src, d=dst):
                shutil.copy2(s, d)
            try:
                write_with_retry(do_copy, max_retries=3, delay=2.0)
                copied.append(dst)
                template_names_actual.append(tname)
            except PermissionError:
                copied.append(dst)
                template_names_actual.append(tname)
        template_names = template_names_actual
    else:
        clear_output_folder(config)
        template_filenames, template_names = [], []
        for r in service.get_workflow_templates(option_key, package_label):
            if r.get("Name", "") in selected_templates:
                fname_raw = _str(r.get("File", ""))
                fname = fname_raw if fname_raw.endswith(".docx") else fname_raw + ".docx"
                template_filenames.append(fname)
                template_names.append(r.get("Name", ""))
        copied = copy_templates_to_output(config, opt, template_filenames)

    total = len(copied)
    logger.write_header(option_key, package_label, total)

    results = []
    used_names = set()
    has_locked = [False]
    has_other_error = [False]
    failed_templates = []

    table_placeholder_names = {
        _str(t.get("Name", "")).strip("{} ")
        for t in tables_rows
        if _str(t.get(key_id)) == goi_thau_id
    }

    for i, (src_path, tpl_name) in enumerate(zip(copied, template_names)):
        if progress_cb:
            progress_cb((i + 1) / total, f"Đang xử lý: {tpl_name}")
        try:
            missing_placeholders = []
            try:
                jenv = jinja2.Environment()
                jenv.filters["date"] = filter_date
                jenv.filters["date_long"] = filter_date_long
                jenv.filters["number"] = filter_number
                jenv.filters["num2text"] = filter_num2text
                jenv.filters["day"] = filter_day
                jenv.filters["month"] = filter_month
                jenv.filters["year"] = filter_year
                jenv.filters["add_days"] = filter_add_days
                jenv.filters["add_months"] = filter_add_months
                jenv.filters["date_diff"] = filter_date_diff
                jenv.filters["quarter"] = filter_quarter
                jenv.filters["weekday"] = filter_weekday
                jenv.filters["date_text"] = filter_date_text

                doc = DocxTemplate(str(src_path))
                undeclared = doc.get_undeclared_template_variables(jinja_env=jenv)
                for var in undeclared:
                    var_clean = var.strip()
                    if var_clean not in nested_context and var_clean not in table_placeholder_names:
                        missing_placeholders.append(var_clean)
            except Exception as e:
                print(f"⚠️  Lỗi quét placeholder cho {tpl_name}: {e}")

            def do_merge(s=src_path):
                return mail_merge_safe(s, nested_context, s)

            def on_locked_retry(msg):
                if progress_cb:
                    progress_cb((i + 1) / total, f"{tpl_name}: {msg}")

            ok, err = write_with_retry(do_merge, max_retries=3, delay=2.0, yield_fn=on_locked_retry)
            if not ok:
                raise RuntimeError(err)

            if danh_muc_file and danh_muc_file.exists():
                try:
                    copy_tables_for_file(src_path, config, goi_thau_id, tables_rows, danh_muc_file, left_key)
                except PermissionError as table_err:
                    if "being used by another process" in str(table_err) or getattr(table_err, 'errno', None) == errno.EACCES:
                        raise PermissionError(table_err) from table_err
                    raise
                except Exception as table_err:
                    print(f"⚠️  Lỗi copy bảng: {table_err}")

            new_path = rename_output(src_path, goi_thau_id, used_names)

            if missing_placeholders:
                warn_msg = f"Warning: Placeholder {', '.join('{{' + k + '}}' for k in missing_placeholders)} không có data"
                logger.record_ok_with_warning(results, tpl_name, new_path.name, warn_msg)
            else:
                logger.record_ok(results, tpl_name, new_path.name)

        except PermissionError as e:
            if "being used by another process" in str(e) or getattr(e, 'errno', None) == errno.EACCES:
                err_msg = "Không thể ghi file vì đang được mở trong Word.\n   → Vui lòng đóng file này và nhấn \"🔄 Chạy lại file lỗi\""
                logger.record_locked(results, tpl_name, has_locked, failed_templates, detail=err_msg)
            else:
                logger.record_error(results, tpl_name, e, has_other_error, failed_templates)

        except Exception as e:
            logger.record_error(results, tpl_name, e, has_other_error, failed_templates)

        yield "\n".join(results), f"Đang xử lý {i + 1}/{total}...", None

    elapsed = time.time() - logger.start_time
    summary = f"✅ {logger.ok_count}  ⚠️ {logger.warning_count}  ❌ {logger.error_count}  /  {total} file  ({elapsed:.1f}s)"
    
    logger.write_footer()

    retry_state_data = None
    if failed_templates:
        retry_state_data = {
            "option_key": option_key,
            "package_label": package_label,
            "failed_templates": failed_templates,
            "all_locked": has_locked[0] and not has_other_error[0],
        }

    yield "\n".join(results), summary, retry_state_data


async def run_retry_batch(service: KisorService, retry_state: dict, progress_cb: Optional[Callable[[float, str], None]] = None):
    if not retry_state or not retry_state.get("failed_templates"):
        yield "⚠️ Không có file lỗi nào để chạy lại", "⚠️ Không có file lỗi", None
        return
    option_key = retry_state["option_key"]
    package_label = retry_state["package_label"]
    failed_templates = retry_state["failed_templates"]
    group_name = retry_state.get("group_name", "")
    async for log, status, new_state in run_batch(
        service=service,
        option_key=option_key,
        package_label=package_label,
        selected_templates=failed_templates,
        group_name=group_name,
        progress_cb=progress_cb,
        retry_state=retry_state
    ):
        yield log, status, new_state
