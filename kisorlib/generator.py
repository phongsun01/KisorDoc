"""
kisorlib/generator.py
─────────────────────
Core sync pipeline — single source of truth cho toàn bộ việc sinh văn bản.

Theo Refactor v3.2:
  - generate_one(...)  : merge 1 file + copy bảng + rename
  - generate_many(...) : lặp template, gọi generate_one, gom FileResult, gọi on_progress
  - build_context(...) : Jinja2 context + nested dict; pure/sync

Không import Gradio. Không gọi load_config() bên trong core.
UI (batch.py) và API (engine.py) chỉ là adapter mỏng gọi vào đây.
"""

from __future__ import annotations

import errno
import shutil
import time
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Optional

import jinja2
import pandas as pd
from docxtpl import DocxTemplate

from .app_helpers import make_nested_dict
from .filters import (
    filter_add_days, filter_add_months, filter_date, filter_date_diff,
    filter_date_long, filter_date_text, filter_day, filter_month,
    filter_num2text, filter_number, filter_quarter, filter_weekday, filter_year,
)
from .file_utils import rename_output
from .merger import mail_merge_safe
from .table_copier import copy_tables_for_file
from .utils import clean_config_key, _str


# ──────────────────────────────────────────────
# Progress event
# ──────────────────────────────────────────────

OnProgress = Optional[Callable[[dict], None]]

# event schema:
# {
#   "level":    "info" | "warning" | "error" | "success",
#   "message":  str,
#   "step":     int | None,
#   "total":    int | None,
#   "template": str | None,
# }

def _emit(cb: OnProgress, level: str, message: str, **extra) -> None:
    if cb is None:
        return
    try:
        cb({"level": level, "message": message, **extra})
    except Exception:
        pass


# ──────────────────────────────────────────────
# FileResult (tái xuất để engine/batch dùng)
# ──────────────────────────────────────────────

@dataclass
class FileResult:
    template_name:    str
    output_path:      Optional[str]  = None
    success:          bool           = False
    error:            Optional[str]  = None
    warnings:         list[str]      = field(default_factory=list)
    dry_run_context:  Optional[dict] = None


# ──────────────────────────────────────────────
# Jinja2 environment factory
# ──────────────────────────────────────────────

def _make_jinja_env(undefined=jinja2.DebugUndefined) -> jinja2.Environment:
    env = jinja2.Environment(undefined=undefined)
    env.filters["date"]       = filter_date
    env.filters["date_long"]  = filter_date_long
    env.filters["number"]     = filter_number
    env.filters["num2text"]   = filter_num2text
    env.filters["day"]        = filter_day
    env.filters["month"]      = filter_month
    env.filters["year"]       = filter_year
    env.filters["add_days"]   = filter_add_days
    env.filters["add_months"] = filter_add_months
    env.filters["date_diff"]  = filter_date_diff
    env.filters["quarter"]    = filter_quarter
    env.filters["weekday"]    = filter_weekday
    env.filters["date_text"]  = filter_date_text
    return env


# ──────────────────────────────────────────────
# write_with_retry (moved here from batch.py)
# ──────────────────────────────────────────────

def write_with_retry(func, max_retries: int = 3, delay: float = 2.0,
                     on_retry: OnProgress = None) -> tuple[bool, str]:
    """
    Thực thi func() có retry khi gặp PermissionError (file đang bị khóa).
    Luôn trả về (bool, str).
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
                _emit(on_retry, "warning", msg, template=None)
                time.sleep(delay)
            else:
                raise


# ──────────────────────────────────────────────
# build_context
# ──────────────────────────────────────────────

def build_context(
    selected_pkg: dict,
    config_rows:  list[dict],
    extra:        dict | None = None,
) -> dict:
    """
    Xây dựng Jinja2 context phẳng từ dòng dữ liệu và mapping Config.

    Tham số:
        selected_pkg : dòng dữ liệu đã được query (dict)
        config_rows  : danh sách mapping {Key, Value} từ sheet Config
        extra        : các key bổ sung thêm vào context sau (VD: HoTen, right_key…)

    Trả về dict phẳng (chưa nested). Caller gọi make_nested_dict() khi cần.
    """
    ctx: dict = {}
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
        ctx[clean_key] = str(raw_value)
    if extra:
        ctx.update(extra)
    return ctx


# ──────────────────────────────────────────────
# detect_missing_placeholders
# ──────────────────────────────────────────────

def detect_missing_placeholders(
    template_path:        Path,
    nested_context:       dict,
    table_placeholder_names: set[str],
) -> list[str]:
    """
    Trả về danh sách placeholder trong template không có trong context
    và không phải là table placeholder.
    """
    try:
        jenv = _make_jinja_env(undefined=jinja2.Undefined)
        doc  = DocxTemplate(str(template_path))
        undeclared = doc.get_undeclared_template_variables(jinja_env=jenv)
        missing = []
        for var in undeclared:
            var_clean = var.strip()
            if var_clean not in nested_context and var_clean not in table_placeholder_names:
                missing.append(var_clean)
        return missing
    except Exception as e:
        print(f"⚠️  Lỗi quét placeholder cho {template_path.name}: {e}")
        return []


# ──────────────────────────────────────────────
# generate_one
# ──────────────────────────────────────────────

def generate_one(
    *,
    template_path:    Path,
    template_name:    str,
    output_path:      Path,
    nested_context:   dict,
    cfg,
    goi_thau_id:      str,
    tables_rows:      list[dict],
    danh_muc_file:    Optional[Path],
    key_id:           str,
    used_names:       set[str],
    table_placeholder_names: set[str] | None = None,
    dry_run:          bool          = False,
    max_retries:      int           = 3,
    retry_delay:      float         = 2.0,
    on_progress:      OnProgress    = None,
) -> FileResult:
    """
    Sinh 1 file: copy template → mail_merge → copy_tables → rename.

    Không biết về Repeat, không biết về Gradio.
    Caller (batch / engine) chịu trách nhiệm chuẩn bị nested_context.
    """
    result = FileResult(template_name=template_name)

    try:
        if dry_run:
            _emit(on_progress, "info", f"[DRY-RUN] {template_name}", template=template_name)
            result.success = True
            result.dry_run_context = nested_context
            return result

        # 1. Copy template → output (với retry)
        if template_path.resolve() != output_path.resolve():
            def do_copy():
                shutil.copy2(str(template_path), str(output_path))
            write_with_retry(do_copy, max_retries=max_retries, delay=retry_delay,
                             on_retry=on_progress)
            _emit(on_progress, "info", f"[1/3] Copy: {template_name}", template=template_name)

        # 2. Detect missing placeholders (trước merge để output warning)
        if table_placeholder_names is not None:
            missing = detect_missing_placeholders(output_path, nested_context, table_placeholder_names)
        else:
            missing = []

        # 3. Mail merge (DebugUndefined: table placeholders sống qua bước này)
        def do_merge():
            return mail_merge_safe(output_path, nested_context, output_path)

        ok, err = write_with_retry(do_merge, max_retries=max_retries, delay=retry_delay,
                                   on_retry=on_progress)
        if not ok:
            raise RuntimeError(err)
        _emit(on_progress, "info", f"[2/3] Merge: {template_name}", template=template_name)

        # 4. Copy bảng Excel → Word (PHẢI sau mail_merge để table placeholder còn trong file)
        if danh_muc_file and danh_muc_file.exists():
            try:
                copy_tables_for_file(
                    output_path,    # doc_path
                    cfg,            # config
                    goi_thau_id,    # goi_thau_id
                    tables_rows,    # tables_data
                    danh_muc_file,  # xlsx_path
                    key_id,         # key_id
                )
                _emit(on_progress, "info", f"[3/3] Bảng: {template_name}", template=template_name)
            except PermissionError:
                raise
            except Exception as te:
                warn = f"Lỗi copy bảng (bỏ qua): {te}"
                result.warnings.append(warn)
                _emit(on_progress, "warning", f"⚠️ {template_name}: {warn}", template=template_name)

        # 5. Rename
        new_path = rename_output(output_path, goi_thau_id, used_names)

        result.success     = True
        result.output_path = str(new_path)
        if missing:
            warn_msg = f"Placeholder {', '.join('{{' + k + '}}' for k in missing)} không có data"
            result.warnings.append(warn_msg)
            _emit(on_progress, "warning",
                  f"⚠️ {template_name} → {new_path.name}: {warn_msg}",
                  template=template_name, output_name=new_path.name)
        else:
            _emit(on_progress, "success",
                  f"✓ {template_name} → {new_path.name}",
                  template=template_name, output_name=new_path.name)

    except PermissionError as exc:
        result.error   = "File đang mở trong Word — đóng lại và chạy lại"
        result.success = False
        _emit(on_progress, "error",
              f"🔒 {template_name}: {result.error}",
              template=template_name, is_locked=True)

    except Exception as exc:
        result.error   = f"{type(exc).__name__}: {exc}"
        result.success = False
        _emit(on_progress, "error",
              f"✗ {template_name}: {result.error}",
              template=template_name, is_locked=False,
              traceback=traceback.format_exc())

    return result


# ──────────────────────────────────────────────
# generate_many
# ──────────────────────────────────────────────

def generate_many(
    *,
    template_paths:          list[tuple[Path, str]],   # [(src_path, template_name), ...]
    nested_context:          dict,
    cfg,
    goi_thau_id:             str,
    tables_rows:             list[dict],
    danh_muc_file:           Optional[Path],
    key_id:                  str,
    table_placeholder_names: set[str] | None = None,
    dry_run:                 bool             = False,
    max_retries:             int              = 3,
    retry_delay:             float            = 2.0,
    on_progress:             OnProgress       = None,
) -> list[FileResult]:
    """
    Lặp qua danh sách template, gọi generate_one cho từng cái.
    Sync — không yield, không async.

    Caller (batch.py async) sẽ yield sau mỗi FileResult nếu muốn
    update UI progressively.
    """
    total      = len(template_paths)
    results    : list[FileResult] = []
    used_names : set[str]         = set()

    for step, (src_path, tpl_name) in enumerate(template_paths, start=1):
        _emit(on_progress, "info",
              f"[{step}/{total}] {tpl_name}",
              step=step, total=total, template=tpl_name)

        # Output path = cùng thư mục với src (output_path đã được batch/engine copy vào)
        out_path = src_path  # src_path đã là file trong output folder sau copy_templates_to_output

        result = generate_one(
            template_path           = src_path,
            template_name           = tpl_name,
            output_path             = out_path,
            nested_context          = nested_context,
            cfg                     = cfg,
            goi_thau_id             = goi_thau_id,
            tables_rows             = tables_rows,
            danh_muc_file           = danh_muc_file,
            key_id                  = key_id,
            used_names              = used_names,
            table_placeholder_names = table_placeholder_names,
            dry_run                 = dry_run,
            max_retries             = max_retries,
            retry_delay             = retry_delay,
            on_progress             = on_progress,
        )
        results.append(result)

    return results


# ──────────────────────────────────────────────
# Repeat orchestration helper
# ──────────────────────────────────────────────

def generate_one_repeat(
    *,
    template_path:   Path,
    template_name:   str,
    member_name:     str,
    nested_context:  dict,
    cfg,
    goi_thau_id:     str,
    tables_rows:     list[dict],
    danh_muc_file:   Optional[Path],
    key_id:          str,
    max_retries:     int          = 3,
    retry_delay:     float        = 2.0,
    on_progress:     OnProgress   = None,
) -> FileResult:
    """
    Sinh 1 file cho 1 thành viên trong Repeat.
    Copy template → rename output path → merge → copy_tables → rename final.
    """
    result = FileResult(template_name=member_name)
    used_names: set[str] = set()

    try:
        dst = cfg.output_path / template_path.name

        if template_path.resolve() != dst.resolve():
            def do_copy():
                shutil.copy2(str(template_path), str(dst))
            write_with_retry(do_copy, max_retries=max_retries, delay=retry_delay,
                             on_retry=on_progress)

        def do_merge():
            return mail_merge_safe(dst, nested_context, dst)

        ok, err = write_with_retry(do_merge, max_retries=max_retries, delay=retry_delay,
                                   on_retry=on_progress)
        if not ok:
            raise RuntimeError(err)

        if danh_muc_file and danh_muc_file.exists():
            try:
                copy_tables_for_file(dst, cfg, goi_thau_id, tables_rows, danh_muc_file, key_id)
            except PermissionError:
                raise
            except Exception as te:
                warn = f"Lỗi copy bảng: {te}"
                result.warnings.append(warn)
                _emit(on_progress, "warning", f"⚠️ {member_name}: {warn}", template=member_name)

        # Rename: stem-goi_thau_id-member_name.docx
        stem = template_path.stem.replace("-Template", "")
        new_filename = f"{stem}-{goi_thau_id}-{member_name}.docx"
        new_path = cfg.output_path / new_filename
        if dst.exists():
            dst.rename(new_path)

        result.success     = True
        result.output_path = str(new_path)
        _emit(on_progress, "success",
              f"✓ {member_name} → {new_path.name}",
              template=member_name, output_name=new_path.name)

    except PermissionError:
        result.error   = "File đang mở trong Word — đóng lại và chạy lại"
        result.success = False
        _emit(on_progress, "error",
              f"🔒 {member_name}: {result.error}",
              template=member_name, is_locked=True)

    except Exception as exc:
        result.error   = f"{type(exc).__name__}: {exc}"
        result.success = False
        _emit(on_progress, "error",
              f"✗ {member_name}: {result.error}",
              template=member_name, is_locked=False,
              traceback=traceback.format_exc())

    return result
