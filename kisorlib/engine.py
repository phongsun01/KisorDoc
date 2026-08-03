"""
kisordoc/engine.py
──────────────────
Public API duy nhất của core library KisorDoc.
Cả Gradio UI (app.py) và FastAPI (api.py) đều chỉ import từ đây.

Caller pattern:
    from kisordoc.engine import generate_documents, GenerateRequest, GenerateResult

    req = GenerateRequest(option="Opt1", package_id="MS26-01", templates=["BaoCao"])
    result = generate_documents(req, on_progress=lambda e: print(e["message"]))
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
# 1. Schema: Input / Output
# ──────────────────────────────────────────────

class GenerateRequest(BaseModel):
    """Tham số đầu vào cho một lần sinh văn bản."""

    # Đường dẫn file Excel data. None → đọc từ AppConfig (ProjectPath/1. Data/)
    excel_path: Optional[str] = None

    # Tên option trong sheet Workflow, ví dụ "Opt1", "Opt2"
    option: str

    # ID gói thầu, ví dụ "MS26-01" — dùng để query sheet GoiThau
    package_id: str

    # Danh sách tên template được chọn (không có extension)
    # Ví dụ: ["BaoCao", "TuTrinhPheDuyet"]
    templates: List[str]

    # Thư mục output. None → đọc từ AppConfig (ProjectPath/3. Files/)
    output_dir: Optional[str] = None

    # Row range cho Config sheet, ví dụ "2-97"
    # Dùng khi nhiều Option dùng chung một Config sheet
    config_row_range: Optional[str] = None

    # Dry-run: render template nhưng KHÔNG ghi file ra đĩa
    # Trả về danh sách placeholder và giá trị tương ứng
    dry_run: bool = False

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
    """Kết quả xử lý một file template."""
    template_name: str
    output_path: Optional[str] = None       # None nếu dry_run hoặc lỗi
    success: bool = False
    error: Optional[str] = None
    dry_run_context: Optional[dict] = None  # Chỉ có khi dry_run=True


@dataclass
class GenerateResult:
    """Kết quả trả về từ generate_documents()."""
    success: bool
    files: List[FileResult] = field(default_factory=list)
    total: int = 0
    succeeded: int = 0
    failed: int = 0
    skipped: int = 0
    duration_seconds: float = 0.0
    error: Optional[str] = None             # Lỗi toàn cục (không phải per-file)

    @property
    def output_paths(self) -> List[str]:
        """Trả về danh sách path của các file đã sinh thành công."""
        return [f.output_path for f in self.files if f.success and f.output_path]


# ──────────────────────────────────────────────
# 2. Progress event helpers
# ──────────────────────────────────────────────

# Các level để caller phân biệt (Gradio dùng để màu log, FastAPI lưu vào store)
LEVEL_INFO    = "info"
LEVEL_SUCCESS = "success"
LEVEL_WARNING = "warning"
LEVEL_ERROR   = "error"

ProgressCallback = Optional[Callable[[dict], None]]


def _emit(callback: ProgressCallback, level: str, message: str, **extra) -> None:
    """Gọi callback với event dict chuẩn. Luôn log ra logger."""
    event = {"level": level, "message": message, **extra}
    if level == LEVEL_ERROR:
        logger.error(message)
    elif level == LEVEL_WARNING:
        logger.warning(message)
    else:
        logger.info(message)
    if callback:
        try:
            callback(event)
        except Exception:
            pass  # callback lỗi không được làm dừng engine


# ──────────────────────────────────────────────
# 3. Internal helpers
# ──────────────────────────────────────────────

def _resolve_paths(req: GenerateRequest) -> tuple[Path, Path, Path]:
    """
    Trả về (data_dir, templates_dir, output_dir) dạng Path.
    Nếu req không chỉ định thì đọc từ AppConfig.
    """
    # Import lazy để tránh circular và cho phép mock trong test
    from .config import load_config  # type: ignore

    cfg = load_config()
    project_path = Path(cfg.ProjectPath)

    data_dir      = Path(req.excel_path).parent if req.excel_path else project_path / "1. Data"
    templates_dir = project_path / "2. Templates"
    output_dir    = Path(req.output_dir) if req.output_dir else project_path / "3. Files"

    return data_dir, templates_dir, output_dir


def _find_excel_files(data_dir: Path) -> List[Path]:
    """Liệt kê tất cả .xlsx trong data_dir (không đệ quy)."""
    return sorted(data_dir.glob("*.xlsx"))


def _find_template_file(templates_dir: Path, option: str, template_name: str) -> Optional[Path]:
    """
    Tìm file template theo cấu trúc: 2. Templates/{option}/{template_name}.docx
    Fallback: 2. Templates/{template_name}.docx (flat)
    """
    nested = templates_dir / option / f"{template_name}.docx"
    if nested.exists():
        return nested
    flat = templates_dir / f"{template_name}.docx"
    if flat.exists():
        return flat
    return None


def _prepare_output_dir(output_dir: Path, cb: ProgressCallback) -> None:
    """Xóa sạch và tạo lại output_dir, với backup trước khi xóa."""
    if output_dir.exists():
        backup_root = output_dir.parent / "_backup"
        backup_root.mkdir(exist_ok=True)
        ts = time.strftime("%Y%m%d_%H%M%S")
        backup_dest = backup_root / ts
        shutil.copytree(str(output_dir), str(backup_dest))
        _emit(cb, LEVEL_INFO, f"Đã backup output cũ → {backup_dest}")
        shutil.rmtree(str(output_dir))
    output_dir.mkdir(parents=True, exist_ok=True)


def _build_context(ds, req: GenerateRequest, cb: ProgressCallback) -> Optional[dict]:
    """
    Query DataSet để lấy context dict cho Jinja2.
    Trả về None nếu không tìm thấy dữ liệu gói thầu.
    """
    from .dataset import DataSet  # type: ignore

    # Query GoiThau
    row = ds.query_row("GoiThau", package_id=req.package_id)
    if row is None:
        _emit(cb, LEVEL_ERROR,
              f"Không tìm thấy gói thầu '{req.package_id}' trong sheet GoiThau")
        return None

    # Query Config (với row range nếu có)
    if req.config_row_range:
        config_rows = ds.query_rows("Config", row_range=req.config_row_range)
    else:
        config_rows = ds.query_rows("Config")

    # Query Options cho option hiện tại
    option_data = ds.query_option(req.option)

    # Merge thành context phẳng — GoiThau > Config > Options
    # (GoiThau có độ ưu tiên cao nhất nếu trùng key)
    ctx: dict = {}
    if config_rows:
        for r in config_rows:
            ctx.update(r)
    if option_data:
        ctx.update(option_data)
    ctx.update(row)  # GoiThau override cuối

    return ctx


# ──────────────────────────────────────────────
# 4. Single-file processing
# ──────────────────────────────────────────────

def _process_one(
    template_path: Path,
    template_name: str,
    output_path: Path,
    context: dict,
    excel_files: List[Path],
    dry_run: bool,
    cb: ProgressCallback,
) -> FileResult:
    """
    Xử lý một template: mail merge → copy bảng → ghi file.
    Trả về FileResult.
    """
    from .merger import mail_merge_safe          # type: ignore
    from .table_copier import copy_tables_for_file  # type: ignore

    result = FileResult(template_name=template_name)

    try:
        if dry_run:
            # Chỉ render, không ghi file — trả về context để caller inspect
            _emit(cb, LEVEL_INFO, f"[DRY-RUN] {template_name}")
            rendered_ctx = mail_merge_safe(
                template_path=str(template_path),
                context=context,
                output_path=None,   # merger phải hỗ trợ output_path=None
                dry_run=True,
            )
            result.success = True
            result.dry_run_context = rendered_ctx
            return result

        # ── Bước 1: Mail merge (Jinja2 / docxtpl) ──
        _emit(cb, LEVEL_INFO, f"Mail merge: {template_name}")
        mail_merge_safe(
            template_path=str(template_path),
            context=context,
            output_path=str(output_path),
        )

        # ── Bước 2: Copy bảng Excel → Word ──
        # QUAN TRỌNG: phải chạy SAU mail_merge để DebugUndefined giữ
        # các table placeholder qua lượt Jinja2 render đầu tiên
        _emit(cb, LEVEL_INFO, f"Copy bảng: {template_name}")
        copy_tables_for_file(
            docx_path=str(output_path),
            excel_files=[str(p) for p in excel_files],
        )

        result.success = True
        result.output_path = str(output_path)
        _emit(cb, LEVEL_SUCCESS, f"✓ {template_name} → {output_path.name}")

    except Exception as exc:
        err_msg = f"{type(exc).__name__}: {exc}"
        result.error = err_msg
        result.success = False
        _emit(cb, LEVEL_ERROR,
              f"✗ {template_name}: {err_msg}",
              traceback=traceback.format_exc())

    return result


# ──────────────────────────────────────────────
# 5. Public entry point
# ──────────────────────────────────────────────

def generate_documents(
    request: GenerateRequest,
    on_progress: ProgressCallback = None,
) -> GenerateResult:
    """
    Sinh hàng loạt văn bản Word từ template + dữ liệu Excel.

    Parameters
    ----------
    request : GenerateRequest
        Toàn bộ tham số đầu vào (validate bằng Pydantic).
    on_progress : Callable[[dict], None], optional
        Callback nhận event dict mỗi khi có cập nhật tiến trình.
        Event dict có dạng:
            {
                "level":   "info" | "success" | "warning" | "error",
                "message": str,
                # optional extras:
                "current": int,   # file đang xử lý (1-based)
                "total":   int,
                "traceback": str,
            }

    Returns
    -------
    GenerateResult
        Kết quả tổng hợp, bao gồm list FileResult cho từng template.

    Notes
    -----
    - Hàm này KHÔNG phải generator — Gradio dùng callback để stream log.
    - Thread-safe ở mức per-call (mỗi lần gọi có output_dir riêng
      nếu caller truyền output_dir khác nhau).
    - dry_run=True: không ghi file, không xóa output_dir cũ.
    """
    t_start = time.monotonic()
    cb = on_progress

    _emit(cb, LEVEL_INFO,
          f"═══ Bắt đầu: option={request.option}, gói={request.package_id} ═══")

    # ── Validate request ──────────────────────────────────────────────────────
    # (Pydantic đã validate khi khởi tạo, nhưng kiểm tra thêm ở runtime)
    if not request.templates:
        return GenerateResult(
            success=False,
            error="Danh sách template trống",
        )

    # ── Resolve paths ─────────────────────────────────────────────────────────
    try:
        data_dir, templates_dir, output_dir = _resolve_paths(request)
    except Exception as exc:
        _emit(cb, LEVEL_ERROR, f"Lỗi đọc cấu hình: {exc}")
        return GenerateResult(success=False, error=str(exc))

    _emit(cb, LEVEL_INFO, f"Data dir     : {data_dir}")
    _emit(cb, LEVEL_INFO, f"Templates dir: {templates_dir}")
    _emit(cb, LEVEL_INFO, f"Output dir   : {output_dir}")

    # ── Tìm file Excel ────────────────────────────────────────────────────────
    excel_files = _find_excel_files(data_dir)
    if not excel_files:
        err = f"Không tìm thấy file .xlsx trong {data_dir}"
        _emit(cb, LEVEL_ERROR, err)
        return GenerateResult(success=False, error=err)
    _emit(cb, LEVEL_INFO, f"Tìm thấy {len(excel_files)} file Excel: "
          + ", ".join(p.name for p in excel_files))

    # ── Khởi tạo DataSet ─────────────────────────────────────────────────────
    try:
        from .config import load_config      # type: ignore
        from .dataset import DataSet         # type: ignore

        cfg = load_config()
        ds = DataSet(cfg, excel_files=[str(p) for p in excel_files])
    except Exception as exc:
        _emit(cb, LEVEL_ERROR, f"Lỗi khởi tạo DataSet: {exc}")
        return GenerateResult(success=False, error=str(exc))

    # ── Build context từ Excel ────────────────────────────────────────────────
    context = _build_context(ds, request, cb)
    if context is None:
        return GenerateResult(success=False,
                              error=f"Không có dữ liệu cho gói '{request.package_id}'")

    _emit(cb, LEVEL_INFO, f"Context đã build: {len(context)} keys")

    # ── Chuẩn bị output dir (bỏ qua nếu dry_run) ────────────────────────────
    if not request.dry_run:
        try:
            _prepare_output_dir(output_dir, cb)
        except Exception as exc:
            _emit(cb, LEVEL_ERROR, f"Lỗi chuẩn bị output dir: {exc}")
            return GenerateResult(success=False, error=str(exc))

    # ── Xử lý từng template ──────────────────────────────────────────────────
    total    = len(request.templates)
    results: List[FileResult] = []
    skipped  = 0

    for idx, tpl_name in enumerate(request.templates, start=1):
        _emit(cb, LEVEL_INFO,
              f"[{idx}/{total}] Đang xử lý: {tpl_name}",
              current=idx, total=total)

        # Tìm template file
        tpl_path = _find_template_file(templates_dir, request.option, tpl_name)
        if tpl_path is None:
            _emit(cb, LEVEL_WARNING,
                  f"[{idx}/{total}] Bỏ qua '{tpl_name}': không tìm thấy file template")
            results.append(FileResult(
                template_name=tpl_name,
                success=False,
                error="Template file không tồn tại",
            ))
            skipped += 1
            continue

        # Tên file output: giữ nguyên tên template
        out_path = output_dir / f"{tpl_name}.docx"

        file_result = _process_one(
            template_path=tpl_path,
            template_name=tpl_name,
            output_path=out_path,
            context=context,
            excel_files=excel_files,
            dry_run=request.dry_run,
            cb=cb,
        )
        results.append(file_result)

    # ── Tổng kết ─────────────────────────────────────────────────────────────
    succeeded = sum(1 for r in results if r.success)
    failed    = sum(1 for r in results if not r.success and r.error != "Template file không tồn tại")
    duration  = time.monotonic() - t_start

    overall_success = failed == 0

    _emit(cb, LEVEL_SUCCESS if overall_success else LEVEL_WARNING,
          f"═══ Hoàn thành: {succeeded}/{total} thành công, "
          f"{failed} lỗi, {skipped} bỏ qua — {duration:.1f}s ═══")

    return GenerateResult(
        success=overall_success,
        files=results,
        total=total,
        succeeded=succeeded,
        failed=failed,
        skipped=skipped,
        duration_seconds=round(duration, 2),
    )


# ──────────────────────────────────────────────
# 6. Utility: list available templates
# ──────────────────────────────────────────────

def list_templates(option: Optional[str] = None) -> List[str]:
    """
    Liệt kê template có sẵn trong thư mục 2. Templates/.
    Nếu option được truyền, chỉ trả về template trong subfolder đó.

    Returns
    -------
    List[str]
        Danh sách tên template (không có extension .docx), đã sort.
    """
    from .config import load_config  # type: ignore

    cfg = load_config()
    templates_dir = Path(cfg.ProjectPath) / "2. Templates"

    if option:
        search_dir = templates_dir / option
    else:
        search_dir = templates_dir

    if not search_dir.exists():
        return []

    return sorted(p.stem for p in search_dir.rglob("*.docx"))


def list_packages() -> List[str]:
    """
    Liệt kê tất cả package_id có trong sheet GoiThau.

    Returns
    -------
    List[str]
        Danh sách ID gói thầu, đã sort.
    """
    from .config import load_config  # type: ignore
    from .dataset import DataSet     # type: ignore

    cfg = load_config()
    data_dir = Path(cfg.ProjectPath) / "1. Data"
    excel_files = _find_excel_files(data_dir)

    if not excel_files:
        return []

    try:
        ds = DataSet(cfg, excel_files=[str(p) for p in excel_files])
        return ds.list_package_ids()
    except Exception:
        return []
