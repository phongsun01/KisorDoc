"""
kisorlib/engine.py — Refactor v3.2
────────────────────────────────────
Pydantic I/O + job/log adapter cho REST API.

KHÔNG còn logic merge/copy/retry nội bộ:
tất cả đi qua generator.generate_many.
"""

from __future__ import annotations

import logging
import time
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, List, Optional

from pydantic import BaseModel, field_validator

logger = logging.getLogger(__name__)


# ──────────────────────────────────────────────
# 1. Schema
# ──────────────────────────────────────────────

class GenerateRequest(BaseModel):
    excel_path:       Optional[str] = None
    option:           str
    package_id:       str
    templates:        List[str]
    output_dir:       Optional[str] = None
    config_row_range: Optional[str] = None
    dry_run:          bool          = False

    @field_validator("option", "package_id")
    @classmethod
    def must_not_be_empty(cls, v: str) -> str:
        if not v or not v.strip():
            raise ValueError("Không được để trống")
        return v.strip()

    @field_validator("templates")
    @classmethod
    def templates_not_empty(cls, v: List[str]) -> List[str]:
        if not v:
            raise ValueError("Phải chọn ít nhất 1 template")
        return [t.strip() for t in v if t.strip()]


@dataclass
class GenerateResult:
    success:          bool
    files:            List = field(default_factory=list)   # List[FileResult]
    total:            int  = 0
    succeeded:        int  = 0
    failed:           int  = 0
    skipped:          int  = 0
    duration_seconds: float = 0.0
    error:            Optional[str] = None

    @property
    def output_paths(self) -> List[str]:
        return [f.output_path for f in self.files if f.success and f.output_path]


# ──────────────────────────────────────────────
# 2. Progress helpers
# ──────────────────────────────────────────────

LEVEL_INFO    = "info"
LEVEL_SUCCESS = "success"
LEVEL_WARNING = "warning"
LEVEL_ERROR   = "error"

ProgressCallback = Optional[Callable[[dict], None]]


def _emit(callback: ProgressCallback, level: str, message: str, **extra) -> None:
    event = {"level": level, "message": message, **extra}
    {LEVEL_ERROR: logger.error, LEVEL_WARNING: logger.warning}.get(
        level, logger.info)(message)
    if callback:
        try:
            callback(event)
        except Exception:
            pass


# ──────────────────────────────────────────────
# 3. Internal helpers
# ──────────────────────────────────────────────

def _resolve_paths(req: GenerateRequest):
    """Trả về (data_dir, templates_dir, output_dir, cfg)."""
    from .config import load_config
    cfg           = load_config()
    data_dir      = cfg.data_path      if not req.excel_path else Path(req.excel_path).parent
    templates_dir = cfg.template_path
    output_dir    = cfg.output_path    if not req.output_dir  else Path(req.output_dir)
    return data_dir, templates_dir, output_dir, cfg


def _find_template_file(templates_dir: Path, option: str, template_name: str) -> Optional[Path]:
    nested = templates_dir / option / f"{template_name}.docx"
    if nested.exists():
        return nested
    flat = templates_dir / f"{template_name}.docx"
    if flat.exists():
        return flat
    possible = list((templates_dir / option).glob(f"{template_name}*.docx"))
    return possible[0] if possible else None


def _prepare_output_dir(output_dir: Path, cb: ProgressCallback) -> None:
    import shutil
    if output_dir.exists():
        backup_root = output_dir.parent / "_backup"
        backup_root.mkdir(exist_ok=True)
        backup_dest = backup_root / time.strftime("%Y%m%d_%H%M%S")
        shutil.copytree(str(output_dir), str(backup_dest))
        _emit(cb, LEVEL_INFO, f"Đã backup output cũ → {backup_dest}")
        shutil.rmtree(str(output_dir))
    output_dir.mkdir(parents=True, exist_ok=True)


def _build_context_for_request(ds, req: GenerateRequest, cfg, cb: ProgressCallback):
    """
    Resolve package row + config rows, trả về (flat_ctx, selected_pkg, key_id, left_key).
    Logic map tập trung qua KisorService + generator.build_context.
    """
    import re as _re
    from .service import KisorService
    from .generator import build_context
    from .utils import _parse_repeat_key_id, resolve_sheet_query

    # FIX ENG-01: khởi tạo KisorService 1 lần duy nhất (tránh query Options 2 lần)
    svc        = KisorService(cfg, ds)
    opt_config = svc.get_option_config(req.option)
    sheet      = opt_config.get("sheet", "GoiThau")
    key_id     = opt_config.get("key_id", "ID")
    left_key, _ = _parse_repeat_key_id(key_id)

    try:
        goi_thau_rows = ds.query(resolve_sheet_query(sheet))
    except Exception as exc:
        _emit(cb, LEVEL_ERROR, f"Lỗi query sheet '{sheet}': {exc}")
        return None, None, key_id, left_key

    selected_pkg = next(
        (r for r in goi_thau_rows
         if str(r.get(left_key, "")).strip() == req.package_id),
        None,
    )
    if selected_pkg is None:
        _emit(cb, LEVEL_ERROR,
              f"Không tìm thấy package_id='{req.package_id}' trong sheet '{sheet}'")
        return None, None, key_id, left_key

    # Config rows — hỗ trợ config_row_range
    if req.config_row_range:
        m = _re.match(r"^(\d+)-(\d+)$", req.config_row_range.strip())
        if m:
            config_rows = ds.query_rows("Config", int(m.group(1)), int(m.group(2)))
        else:
            config_rows = svc.get_config_for_option(req.option)
    else:
        config_rows = svc.get_config_for_option(req.option)

    flat_ctx = build_context(selected_pkg, config_rows)
    return flat_ctx, selected_pkg, key_id, left_key


# ──────────────────────────────────────────────
# 4. Public entry point
# ──────────────────────────────────────────────

def generate_documents(
    request:     GenerateRequest,
    on_progress: ProgressCallback = None,
) -> GenerateResult:
    from datetime import datetime as _dt
    from .app_helpers import make_nested_dict
    from .config import load_config
    from .dataset import DataSet
    from .file_utils import copy_templates_to_output
    from .generator import FileResult, generate_many

    t_start = time.monotonic()
    cb      = on_progress

    _emit(cb, LEVEL_INFO,
          f"═══ Bắt đầu: option={request.option}, id={request.package_id} ═══")

    if not request.templates:
        return GenerateResult(success=False, error="Danh sách template trống")

    # Resolve paths + cfg
    try:
        data_dir, templates_dir, output_dir, cfg = _resolve_paths(request)
    except Exception as exc:
        _emit(cb, LEVEL_ERROR, f"Lỗi config: {exc}")
        return GenerateResult(success=False, error=str(exc))

    try:
        ds = DataSet(cfg)
    except Exception as exc:
        _emit(cb, LEVEL_ERROR, f"Lỗi DataSet: {exc}")
        return GenerateResult(success=False, error=str(exc))

    # Build context qua generator.build_context (cùng logic với batch)
    flat_ctx, selected_pkg, key_id, left_key = _build_context_for_request(ds, request, cfg, cb)
    if flat_ctx is None:
        return GenerateResult(
            success=False,
            error=f"Không có dữ liệu cho '{request.package_id}'"
        )
    _emit(cb, LEVEL_INFO, f"Context: {len(flat_ctx)} keys")

    nested_ctx = make_nested_dict(flat_ctx)
    nested_ctx["now"] = _dt.now()

    # Tables + danh_muc_file
    try:
        tables_rows = ds.query("SELECT * FROM Tables")
    except Exception:
        tables_rows = []

    xlsx_files    = sorted(data_dir.glob("*.xlsx"))
    danh_muc_file = next(
        (f for f in xlsx_files if cfg.DanhMucFile.lower() in f.stem.lower()),
        xlsx_files[0] if xlsx_files else None,
    )

    table_placeholder_names = {
        str(t.get("Name", "")).strip("{} ")
        for t in tables_rows
        if str(t.get(left_key, "")).strip() == request.package_id
    }

    # Chuẩn bị output dir
    if not request.dry_run:
        try:
            _prepare_output_dir(output_dir, cb)
        except Exception as exc:
            _emit(cb, LEVEL_ERROR, f"Lỗi output dir: {exc}")
            return GenerateResult(success=False, error=str(exc))

    # Resolve template files
    total   = len(request.templates)
    skipped = 0
    template_paths: list[tuple[Path, str]] = []
    skipped_results: list[FileResult]      = []

    for tpl_name in request.templates:
        tpl_path = _find_template_file(templates_dir, request.option, tpl_name)
        if tpl_path is None:
            _emit(cb, LEVEL_WARNING, f"Bỏ qua '{tpl_name}': không tìm thấy template")
            skipped_results.append(FileResult(
                template_name=tpl_name, success=False,
                error="Template file không tồn tại"))
            skipped += 1
            continue
        # Copy vào output_dir trước khi generate_many (giống batch)
        out_copy = output_dir / tpl_path.name
        if not request.dry_run:
            import shutil
            try:
                if tpl_path.resolve() != out_copy.resolve():
                    shutil.copy2(str(tpl_path), str(out_copy))
            except Exception as exc:
                _emit(cb, LEVEL_ERROR, f"Lỗi copy template '{tpl_name}': {exc}")
                skipped_results.append(FileResult(
                    template_name=tpl_name, success=False, error=str(exc)))
                skipped += 1
                continue
        template_paths.append((out_copy if not request.dry_run else tpl_path, tpl_name))

    # on_progress adapter: engine log chuẩn
    def _engine_progress(event: dict):
        _emit(cb, event.get("level", "info"), event.get("message", ""),
              **{k: v for k, v in event.items() if k not in ("level", "message")})

    # Gọi core generate_many — cùng pipeline với batch
    file_results = generate_many(
        template_paths          = template_paths,
        nested_context          = nested_ctx,
        cfg                     = cfg,
        goi_thau_id             = request.package_id,
        tables_rows             = tables_rows,
        danh_muc_file           = danh_muc_file,
        key_id                  = left_key,
        table_placeholder_names = table_placeholder_names,
        dry_run                 = request.dry_run,
        max_retries             = getattr(cfg, "FileMaxRetries", 3),   # FIX CFG-01
        retry_delay             = getattr(cfg, "FileRetryDelay", 2.0), # FIX CFG-01
        on_progress             = _engine_progress,
    )

    all_results = skipped_results + file_results
    succeeded   = sum(1 for r in all_results if r.success)
    failed      = sum(1 for r in all_results if not r.success and r.error != "Template file không tồn tại")
    duration    = time.monotonic() - t_start

    _emit(cb, LEVEL_SUCCESS if failed == 0 else LEVEL_WARNING,
          f"═══ Xong: {succeeded}/{total} OK, {failed} lỗi, {skipped} bỏ qua "
          f"— {duration:.1f}s ═══")

    return GenerateResult(
        success          = (failed == 0),
        files            = all_results,
        total            = total,
        succeeded        = succeeded,
        failed           = failed,
        skipped          = skipped,
        duration_seconds = round(duration, 2),
    )


# ──────────────────────────────────────────────
# 5. Utilities
# ──────────────────────────────────────────────

def list_templates(option: Optional[str] = None) -> List[str]:
    from .config import load_config
    cfg    = load_config()
    search = (cfg.template_path / option) if option else cfg.template_path
    if not search.exists():
        return []
    return sorted(p.stem for p in search.rglob("*.docx"))


def list_packages(option_key: str = "") -> List[str]:
    from .config import load_config
    from .dataset import DataSet
    from .service import KisorService
    try:
        cfg        = load_config()
        ds         = DataSet(cfg)
        opt_config = KisorService(cfg, ds).get_option_config(option_key) if option_key else {}
        key_id     = opt_config.get("key_id", "ID")
        sheet      = opt_config.get("sheet", "GoiThau")
        rows       = ds.query(f'SELECT DISTINCT "{key_id}" FROM "{sheet}"')
        return sorted(str(r.get(key_id, "")).strip() for r in rows if r.get(key_id))
    except Exception:
        return []
