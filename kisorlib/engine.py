"""
kisordoc/engine.py  — PATCHED
──────────────────────────────
Fixes so với version cũ:
  #1  Package import: relative import đúng (engine.py nằm trong kisorlib/)
  #3  _build_context: bỏ method ảo query_row/query_option,
      viết lại dùng ds.query() + ds.query_rows(sheet, start, end)
  #4  list_packages: bỏ ds.list_package_ids(), thay bằng ds.query()
  #5  copy_tables_for_file: đúng 6 tham số như table_copier.py thực tế
  #6  DataSet(cfg) — bỏ tham số excel_files không tồn tại
"""

from __future__ import annotations

import logging
import shutil
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
class FileResult:
    template_name:   str
    output_path:     Optional[str]  = None
    success:         bool           = False
    error:           Optional[str]  = None
    dry_run_context: Optional[dict] = None


@dataclass
class GenerateResult:
    success:          bool
    files:            List[FileResult] = field(default_factory=list)
    total:            int              = 0
    succeeded:        int              = 0
    failed:           int              = 0
    skipped:          int              = 0
    duration_seconds: float            = 0.0
    error:            Optional[str]    = None

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
    from .config import load_config  # type: ignore  # FIX #1
    cfg = load_config()
    data_dir      = cfg.data_path      if not req.excel_path  else Path(req.excel_path).parent
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
    # Glob fallback — giống app.py dòng 949
    possible = list((templates_dir / option).glob(f"{template_name}*.docx"))
    return possible[0] if possible else None


def _prepare_output_dir(output_dir: Path, cb: ProgressCallback) -> None:
    if output_dir.exists():
        backup_root = output_dir.parent / "_backup"
        backup_root.mkdir(exist_ok=True)
        backup_dest = backup_root / time.strftime("%Y%m%d_%H%M%S")
        shutil.copytree(str(output_dir), str(backup_dest))
        _emit(cb, LEVEL_INFO, f"Đã backup output cũ → {backup_dest}")
        shutil.rmtree(str(output_dir))
    output_dir.mkdir(parents=True, exist_ok=True)


def _get_option_config_from_ds(ds, option_key: str) -> dict:
    """Mirror của get_option_config() trong app.py — nhận ds thay vì global."""
    opt_code = option_key.split(":")[0].strip() if ":" in option_key else option_key.strip()
    try:
        rows = ds.query("SELECT * FROM Options")
    except Exception:
        rows = []
    for r in rows:
        if str(r.get("Key", "")).strip() == opt_code:
            s = lambda v, d="": str(v).strip() if v else d
            return {
                "sheet":        s(r.get("Sheet"),  "GoiThau"),
                "show":         s(r.get("Show"),   "{TT}. {Số hiệu gói thầu} - {Tên gói thầu}"),
                "key_id":       s(r.get("KeyId"),  "ID"),
                "config_range": s(r.get("Config"), ""),
                "type":         s(r.get("Type"),   ""),
            }
    return {"sheet": "GoiThau", "show": "", "key_id": "ID", "config_range": "", "type": ""}


def _clean_config_key(key: str) -> str:
    """Mirror của clean_config_key() trong app.py."""
    clean = key.strip("<>{}| ")
    for suffix in (".Date.Long", ".Date.long", ".date_long"):
        if clean.endswith(suffix):
            return clean[:-len(suffix)] + "_Date"
    if clean.endswith(".Date") or clean.endswith(".date"):
        return clean[:-5] + "_Date"
    if clean.endswith(".Day") or clean.endswith(".day"):
        return clean[:-4] + "_Date"
    if clean.endswith(".Month") or clean.endswith(".month"):
        return clean[:-6] + "_Date"
    if clean.endswith(".Year") or clean.endswith(".year"):
        return clean[:-5] + "_Date"
    for suffix in (".Upper", ".upper", ".Number", ".number"):
        if clean.endswith(suffix):
            clean = clean[:-len(suffix)]
            break
    if "|" in clean:
        clean = clean.split("|")[0].strip()
    return clean


# FIX #3: Viết lại hoàn toàn — không còn method ảo nào
def _build_context(ds, req: GenerateRequest, cfg, cb: ProgressCallback) -> Optional[dict]:
    """
    Build Jinja2 context từ DataSet.
    Logic y hệt run_batch trong app.py (dòng 1008–1054), không có gì mới.
    """
    import math as _math
    import pandas as _pd
    from datetime import datetime as _dt

    opt_config = _get_option_config_from_ds(ds, req.option)
    sheet      = opt_config.get("sheet", "GoiThau")
    key_id     = opt_config.get("key_id", "ID")

    # Query GoiThau (hoặc sheet tương đương)
    try:
        goi_thau_rows = ds.query(f'SELECT * FROM "{sheet}"')
    except Exception as exc:
        _emit(cb, LEVEL_ERROR, f"Lỗi query sheet '{sheet}': {exc}")
        return None

    selected_pkg = next(
        (r for r in goi_thau_rows
         if str(r.get(key_id, "")).strip() == req.package_id),
        None,
    )
    if selected_pkg is None:
        _emit(cb, LEVEL_ERROR,
              f"Không tìm thấy package_id='{req.package_id}' "
              f"trong sheet '{sheet}' (cột '{key_id}')")
        return None

    # Query Config rows — FIX #3: dùng ds.query_rows(sheet, start, end) đúng signature
    config_rows: list[dict] = []
    if req.config_row_range:
        import re as _re
        m = _re.match(r"^(\d+)-(\d+)$", req.config_row_range.strip())
        if m:
            config_rows = ds.query_rows("Config", int(m.group(1)), int(m.group(2)))
    if not config_rows:
        try:
            config_rows = ds.query("SELECT * FROM Config")
        except Exception:
            config_rows = []

    # Build context phẳng — y hệt app.py dòng 1030–1051
    ctx: dict = {}
    for r in config_rows:
        key = str(r.get("Key", "") or "").strip()
        col = str(r.get("Value", "") or "").strip()
        if not key or not col:
            continue
        clean_key = _clean_config_key(key)
        raw_value = selected_pkg.get(col, "")
        try:
            is_na = _pd.isna(raw_value)
        except (TypeError, ValueError):
            is_na = False
        if is_na:
            raw_value = ""
        elif isinstance(raw_value, _dt):
            raw_value = raw_value.strftime("%d/%m/%Y")
        elif raw_value is None:
            raw_value = ""
        ctx[clean_key] = str(raw_value)

    return ctx


# ──────────────────────────────────────────────
# 4. Single-file processing
# ──────────────────────────────────────────────

def _process_one(
    template_path:  Path,
    template_name:  str,
    output_path:    Path,
    context:        dict,
    cfg,
    goi_thau_id:    str,
    tables_data:    list,
    danh_muc_file:  Optional[Path],
    key_id:         str,
    dry_run:        bool,
    cb:             ProgressCallback,
) -> FileResult:
    """
    Copy template → mail_merge → copy_tables → rename.
    FIX #5: copy_tables_for_file(doc_path, config, goi_thau_id,
                                  tables_data, xlsx_path, key_id)
    """
    from .merger       import mail_merge_safe       # type: ignore  # FIX #1
    from .table_copier import copy_tables_for_file  # type: ignore
    from .file_utils   import rename_output         # type: ignore

    result = FileResult(template_name=template_name)

    try:
        if dry_run:
            _emit(cb, LEVEL_INFO, f"[DRY-RUN] {template_name}")
            result.success        = True
            result.dry_run_context = context
            return result

        # 1. Copy template → output
        shutil.copy2(str(template_path), str(output_path))
        _emit(cb, LEVEL_INFO, f"[1/3] Copy: {template_name}")

        # 2. Mail merge
        from datetime import datetime as _dt
        from .app_helpers import make_nested_dict  # type: ignore

        nested_ctx = make_nested_dict(context)
        nested_ctx["now"] = _dt.now()

        ok, err = mail_merge_safe(output_path, nested_ctx, output_path)
        if not ok:
            raise RuntimeError(f"mail_merge_safe: {err}")
        _emit(cb, LEVEL_INFO, f"[2/3] Merge: {template_name}")

        # 3. Copy bảng — QUAN TRỌNG: sau mail_merge (DebugUndefined)
        # FIX #5: signature đúng
        if danh_muc_file and danh_muc_file.exists():
            try:
                copy_tables_for_file(
                    output_path,    # doc_path: Path
                    cfg,            # config: AppConfig
                    goi_thau_id,    # goi_thau_id: str
                    tables_data,    # tables_data: list[dict]
                    danh_muc_file,  # xlsx_path: Path
                    key_id,         # key_id: str
                )
                _emit(cb, LEVEL_INFO, f"[3/3] Bảng: {template_name}")
            except PermissionError:
                raise
            except Exception as te:
                _emit(cb, LEVEL_WARNING, f"⚠️ Lỗi copy bảng (bỏ qua): {te}")

        # 4. Rename
        used_names: set[str] = set()
        new_path = rename_output(output_path, goi_thau_id, used_names)

        result.success     = True
        result.output_path = str(new_path)
        _emit(cb, LEVEL_SUCCESS,
              f"✓ {template_name} → {new_path.name}",
              template_name=template_name,
              output_name=new_path.name)

    except PermissionError as exc:
        result.error   = "File đang mở trong Word — đóng lại và chạy lại"
        result.success = False
        _emit(cb, LEVEL_ERROR,
              f"🔒 {template_name}: {result.error}",
              template_name=template_name,
              is_locked=True)

    except Exception as exc:
        result.error   = f"{type(exc).__name__}: {exc}"
        result.success = False
        _emit(cb, LEVEL_ERROR,
              f"✗ {template_name}: {result.error}",
              template_name=template_name,
              is_locked=False,
              traceback=traceback.format_exc())

    return result


# ──────────────────────────────────────────────
# 5. Public entry point
# ──────────────────────────────────────────────

def generate_documents(
    request:     GenerateRequest,
    on_progress: ProgressCallback = None,
) -> GenerateResult:
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

    # FIX #6: DataSet(cfg) — không truyền excel_files
    try:
        from .dataset import DataSet  # type: ignore
        ds = DataSet(cfg)
    except Exception as exc:
        _emit(cb, LEVEL_ERROR, f"Lỗi DataSet: {exc}")
        return GenerateResult(success=False, error=str(exc))

    # FIX #3: _build_context mới
    context = _build_context(ds, request, cfg, cb)
    if context is None:
        return GenerateResult(
            success=False,
            error=f"Không có dữ liệu cho '{request.package_id}'"
        )
    _emit(cb, LEVEL_INFO, f"Context: {len(context)} keys")

    # Metadata cho copy_tables
    opt_config = _get_option_config_from_ds(ds, request.option)
    key_id     = opt_config.get("key_id", "ID")
    try:
        tables_data = ds.query("SELECT * FROM Tables")
    except Exception:
        tables_data = []

    xlsx_files    = sorted(data_dir.glob("*.xlsx"))
    danh_muc_file = next(
        (f for f in xlsx_files
         if "DanhMuc" in f.stem or "danh muc" in f.stem.lower()),
        xlsx_files[0] if xlsx_files else None,
    )

    # Chuẩn bị output dir
    if not request.dry_run:
        try:
            _prepare_output_dir(output_dir, cb)
        except Exception as exc:
            _emit(cb, LEVEL_ERROR, f"Lỗi output dir: {exc}")
            return GenerateResult(success=False, error=str(exc))

    # Loop templates
    total   = len(request.templates)
    results: List[FileResult] = []
    skipped = 0

    for idx, tpl_name in enumerate(request.templates, start=1):
        _emit(cb, LEVEL_INFO,
              f"[{idx}/{total}] {tpl_name}", current=idx, total=total)

        tpl_path = _find_template_file(templates_dir, request.option, tpl_name)
        if tpl_path is None:
            _emit(cb, LEVEL_WARNING, f"Bỏ qua '{tpl_name}': không tìm thấy template")
            results.append(FileResult(
                template_name=tpl_name, success=False,
                error="Template file không tồn tại"))
            skipped += 1
            continue

        out_path = output_dir / tpl_path.name

        results.append(_process_one(
            template_path = tpl_path,
            template_name = tpl_name,
            output_path   = out_path,
            context       = context,
            cfg           = cfg,
            goi_thau_id   = request.package_id,
            tables_data   = tables_data,
            danh_muc_file = danh_muc_file,
            key_id        = key_id,
            dry_run       = request.dry_run,
            cb            = cb,
        ))

    succeeded = sum(1 for r in results if r.success)
    failed    = sum(1 for r in results
                    if not r.success and r.error != "Template file không tồn tại")
    duration  = time.monotonic() - t_start

    _emit(cb, LEVEL_SUCCESS if failed == 0 else LEVEL_WARNING,
          f"═══ Xong: {succeeded}/{total} OK, {failed} lỗi, {skipped} bỏ qua "
          f"— {duration:.1f}s ═══")

    return GenerateResult(
        success          = (failed == 0),
        files            = results,
        total            = total,
        succeeded        = succeeded,
        failed           = failed,
        skipped          = skipped,
        duration_seconds = round(duration, 2),
    )


# ──────────────────────────────────────────────
# 6. Utilities
# ──────────────────────────────────────────────

def list_templates(option: Optional[str] = None) -> List[str]:
    from .config import load_config  # type: ignore
    cfg       = load_config()
    search    = (cfg.template_path / option) if option else cfg.template_path
    if not search.exists():
        return []
    return sorted(p.stem for p in search.rglob("*.docx"))


def list_packages(option_key: str = "") -> List[str]:
    """FIX #4: bỏ ds.list_package_ids() — dùng ds.query() trực tiếp."""
    from .config  import load_config  # type: ignore
    from .dataset import DataSet      # type: ignore
    try:
        cfg        = load_config()
        ds         = DataSet(cfg)
        opt_config = _get_option_config_from_ds(ds, option_key) if option_key else {}
        key_id     = opt_config.get("key_id", "ID")
        sheet      = opt_config.get("sheet",  "GoiThau")
        rows       = ds.query(f'SELECT DISTINCT "{key_id}" FROM "{sheet}"')
        return sorted(str(r.get(key_id, "")).strip() for r in rows if r.get(key_id))
    except Exception:
        return []
