"""
api.py
──────
FastAPI HTTP API cho KisorDoc.
Chạy song song với Gradio UI (app.py) qua runner.py.

Endpoints:
    GET  /                        → health check
    GET  /options                 → danh sách option/quy trình
    GET  /packages?option=Opt1    → danh sách gói thầu theo option
    GET  /templates?option=Opt1   → danh sách template theo option
    POST /generate                → sinh văn bản (background job)
    GET  /jobs/{job_id}           → trạng thái + log của job
    GET  /jobs/{job_id}/files     → download file output (zip)
    DELETE /jobs/{job_id}         → xóa job khỏi store

Chạy độc lập:
    uvicorn api:app --host 0.0.0.0 --port 8000 --reload

Hay qua runner.py:
    python runner.py
"""

from __future__ import annotations

import io
import os
import threading
import time
import uuid
import zipfile
from datetime import datetime
from enum import Enum
from pathlib import Path
from typing import Any, Dict, List, Optional

from fastapi import BackgroundTasks, FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from pydantic import BaseModel

# ── Import engine ────────────────────────────────────────────────────────────
from kisordoc.engine import (
    LEVEL_ERROR,
    LEVEL_SUCCESS,
    LEVEL_WARNING,
    GenerateRequest,
    GenerateResult,
    generate_documents,
    list_packages,
    list_templates,
)
from kisordoc.config import load_config   # type: ignore


# ──────────────────────────────────────────────────────────────────────────────
# App setup
# ──────────────────────────────────────────────────────────────────────────────

API_VERSION = Path("VERSION").read_text(encoding="utf-8").strip() if Path("VERSION").exists() else "?"

app = FastAPI(
    title="KisorDoc API",
    description="REST API để sinh văn bản Word từ template + dữ liệu Excel.",
    version=API_VERSION,
    docs_url="/docs",
    redoc_url="/redoc",
)

# CORS — cho phép web app khác gọi từ localhost
app.add_middleware(
    CORSMiddleware,
    allow_origins=os.environ.get("CORS_ORIGINS", "http://localhost:3000,http://localhost:5173").split(","),
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ──────────────────────────────────────────────────────────────────────────────
# Job store (in-memory)
# Đủ cho dùng nội bộ. Scale ra thì thay bằng Redis.
# ──────────────────────────────────────────────────────────────────────────────

class JobStatus(str, Enum):
    PENDING  = "pending"
    RUNNING  = "running"
    DONE     = "done"
    FAILED   = "failed"


class JobRecord(BaseModel):
    job_id:     str
    status:     JobStatus = JobStatus.PENDING
    created_at: str
    finished_at: Optional[str] = None
    request:    Dict[str, Any] = {}
    log:        List[Dict[str, Any]] = []
    result:     Optional[Dict[str, Any]] = None
    error:      Optional[str] = None


# Dict[job_id → JobRecord]  +  Lock để thread-safe
_job_store: Dict[str, JobRecord] = {}
_store_lock = threading.Lock()

# Giữ tối đa N job gần nhất trong memory
_MAX_JOBS = int(os.environ.get("KISORDOC_MAX_JOBS", 100))


def _new_job(req: GenerateRequest) -> JobRecord:
    job_id = str(uuid.uuid4())
    record = JobRecord(
        job_id=job_id,
        created_at=datetime.now().isoformat(),
        request=req.model_dump(),
    )
    with _store_lock:
        _job_store[job_id] = record
        # Evict job cũ nhất nếu vượt giới hạn
        if len(_job_store) > _MAX_JOBS:
            oldest = next(iter(_job_store))
            del _job_store[oldest]
    return record


def _get_job(job_id: str) -> JobRecord:
    with _store_lock:
        job = _job_store.get(job_id)
    if not job:
        raise HTTPException(status_code=404, detail=f"Job '{job_id}' không tồn tại")
    return job


def _update_job(job_id: str, **kwargs):
    with _store_lock:
        job = _job_store.get(job_id)
        if job:
            for k, v in kwargs.items():
                setattr(job, k, v)


# ──────────────────────────────────────────────────────────────────────────────
# Background worker
# ──────────────────────────────────────────────────────────────────────────────

def _run_job(job_id: str, req: GenerateRequest):
    """Chạy trong background thread — cập nhật job store qua callback."""
    _update_job(job_id, status=JobStatus.RUNNING)

    def on_progress(event: dict):
        with _store_lock:
            job = _job_store.get(job_id)
            if job:
                job.log.append({**event, "ts": datetime.now().isoformat()})

    try:
        result: GenerateResult = generate_documents(req, on_progress=on_progress)

        _update_job(
            job_id,
            status=JobStatus.DONE if result.success else JobStatus.FAILED,
            finished_at=datetime.now().isoformat(),
            result={
                "success":          result.success,
                "total":            result.total,
                "succeeded":        result.succeeded,
                "failed":           result.failed,
                "skipped":          result.skipped,
                "duration_seconds": result.duration_seconds,
                "output_paths":     result.output_paths,
                "files": [
                    {
                        "template_name": f.template_name,
                        "success":       f.success,
                        "output_path":   f.output_path,
                        "error":         f.error,
                    }
                    for f in result.files
                ],
            },
        )
    except Exception as exc:
        _update_job(
            job_id,
            status=JobStatus.FAILED,
            finished_at=datetime.now().isoformat(),
            error=str(exc),
        )


# ──────────────────────────────────────────────────────────────────────────────
# Request / Response schemas
# ──────────────────────────────────────────────────────────────────────────────

class GenerateAPIRequest(BaseModel):
    """
    Body của POST /generate.
    Mở rộng GenerateRequest với thêm trường API-specific.
    """
    option:           str
    package_label:    str            # label hiển thị, ví dụ "01. MS26-01 - Gói thầu XYZ"
    templates:        List[str]      # danh sách tên template (không extension)
    config_row_range: Optional[str] = None
    dry_run:          bool = False
    # Tương lai: webhook_url để notify khi xong
    # webhook_url: Optional[str] = None


class JobCreatedResponse(BaseModel):
    job_id:    str
    status:    str
    poll_url:  str


class JobStatusResponse(BaseModel):
    job_id:      str
    status:      str
    created_at:  str
    finished_at: Optional[str]
    log_count:   int
    log:         List[Dict[str, Any]]
    result:      Optional[Dict[str, Any]]
    error:       Optional[str]


# ──────────────────────────────────────────────────────────────────────────────
# Helper: resolve package_label → package_id
# ──────────────────────────────────────────────────────────────────────────────

def _resolve_package_id(option: str, package_label: str) -> str:
    """
    main.py dùng package_label (chuỗi hiển thị từ safe_format).
    engine.py dùng package_id (giá trị cột key_id).

    Hàm này query ds để map label → id.
    Nếu package_label đã là ID (không chứa " - ") thì trả thẳng.
    """
    # Import lazy
    from kisordoc.config import load_config          # type: ignore
    from kisordoc.dataset import DataSet             # type: ignore
    from kisordoc.engine import _find_excel_files    # type: ignore

    cfg = load_config()
    data_dir = Path(cfg.ProjectPath) / "1. Data"
    excel_files = _find_excel_files(data_dir)
    if not excel_files:
        return package_label  # fallback

    ds = DataSet(cfg, excel_files=[str(p) for p in excel_files])

    # Đọc option config để lấy show_format và key_id
    try:
        opt_rows = ds.query("SELECT * FROM Options")
    except Exception:
        opt_rows = []

    key_id      = "ID"
    show_format = "{TT}. {Số hiệu gói thầu} - {Tên gói thầu}"

    for r in opt_rows:
        k = str(r.get("Key", "")).strip()
        if k == option:
            key_id      = str(r.get("KeyId", "ID")).strip() or "ID"
            show_format = str(r.get("Show", show_format)).strip()
            break

    # Query GoiThau rows và tìm dòng khớp label
    try:
        rows = ds.query('SELECT * FROM "GoiThau"')
    except Exception:
        return package_label

    import re as _re

    def _safe_format(pattern: str, row: dict) -> str:
        res = pattern
        for ph in _re.findall(r"\{(.*?)\}", pattern):
            res = res.replace(f"{{{ph}}}", str(row.get(ph, "")).strip())
        return res

    for row in rows:
        label = _safe_format(show_format, row)
        if label == package_label:
            return str(row.get(key_id, package_label)).strip()

    # Không tìm thấy → giả sử caller truyền thẳng ID
    return package_label


# ──────────────────────────────────────────────────────────────────────────────
# Routes
# ──────────────────────────────────────────────────────────────────────────────

@app.get("/", tags=["health"])
def health():
    """Health check — kiểm tra API đang sống."""
    return {
        "status":  "ok",
        "version": API_VERSION,
        "time":    datetime.now().isoformat(),
    }


@app.get("/options", tags=["metadata"])
def get_options():
    """
    Trả về danh sách option (quy trình) từ sheet Options.
    Mỗi item có dạng: {"key": "Opt1", "label": "Opt1: Đấu thầu rộng rãi"}
    """
    try:
        from kisordoc.config import load_config      # type: ignore
        from kisordoc.dataset import DataSet         # type: ignore
        from kisordoc.engine import _find_excel_files  # type: ignore

        cfg = load_config()
        data_dir = Path(cfg.ProjectPath) / "1. Data"
        excel_files = _find_excel_files(data_dir)
        if not excel_files:
            return []
        ds = DataSet(cfg, excel_files=[str(p) for p in excel_files])
        rows = ds.query("SELECT Key, Value FROM Options ORDER BY Key")
        return [
            {"key": str(r.get("Key", "")).strip(),
             "label": f"{str(r.get('Key','')).strip()}: {str(r.get('Value','')).strip()}"}
            for r in rows
        ]
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.get("/packages", tags=["metadata"])
def get_packages(option: str = Query(..., description="Option key, ví dụ 'Opt1'")):
    """
    Trả về danh sách gói thầu theo option.
    Mỗi item: {"id": "MS26-01", "label": "01. MS26-01 - Tên gói thầu"}
    """
    try:
        pkgs = list_packages()
        # list_packages() trả về ID list; label cần query thêm
        # Tạm thời trả id = label, sau bổ sung show_format
        return [{"id": p, "label": p} for p in pkgs]
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.get("/templates", tags=["metadata"])
def get_templates(option: str = Query(..., description="Option key, ví dụ 'Opt1'")):
    """Trả về danh sách tên template (không extension) theo option."""
    try:
        return list_templates(option=option)
    except Exception as exc:
        raise HTTPException(status_code=500, detail=str(exc))


@app.post("/generate", response_model=JobCreatedResponse, status_code=202, tags=["generate"])
def create_generate_job(body: GenerateAPIRequest, background_tasks: BackgroundTasks):
    """
    Tạo job sinh văn bản và chạy nền.
    Trả về job_id ngay lập tức — caller poll GET /jobs/{job_id} để lấy kết quả.

    Body ví dụ:
    ```json
    {
        "option": "Opt1",
        "package_label": "01. MS26-01 - Gói thầu tư vấn XYZ",
        "templates": ["BaoCao", "TuTrinhPheDuyet"],
        "config_row_range": "2-97",
        "dry_run": false
    }
    ```
    """
    # Resolve label → id
    try:
        package_id = _resolve_package_id(body.option, body.package_label)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Không resolve được package: {exc}")

    # Build GenerateRequest (Pydantic validate)
    try:
        req = GenerateRequest(
            option=body.option,
            package_id=package_id,
            templates=body.templates,
            config_row_range=body.config_row_range,
            dry_run=body.dry_run,
        )
    except Exception as exc:
        raise HTTPException(status_code=422, detail=str(exc))

    job = _new_job(req)
    background_tasks.add_task(_run_job, job.job_id, req)

    return JobCreatedResponse(
        job_id=job.job_id,
        status=JobStatus.PENDING,
        poll_url=f"/jobs/{job.job_id}",
    )


@app.get("/jobs/{job_id}", response_model=JobStatusResponse, tags=["jobs"])
def get_job_status(job_id: str, log_offset: int = Query(0, description="Bỏ qua N log đầu")):
    """
    Lấy trạng thái + log của job.
    Dùng log_offset để poll incremental — chỉ lấy log mới từ offset trở đi.

    Poll pattern cho client:
        offset = 0
        while True:
            res = GET /jobs/{id}?log_offset={offset}
            offset += len(res.log)
            if res.status in ("done", "failed"): break
            sleep(0.5)
    """
    job = _get_job(job_id)
    with _store_lock:
        log_slice = job.log[log_offset:]

    return JobStatusResponse(
        job_id=job.job_id,
        status=job.status,
        created_at=job.created_at,
        finished_at=job.finished_at,
        log_count=len(job.log),
        log=log_slice,
        result=job.result,
        error=job.error,
    )


@app.get("/jobs/{job_id}/files", tags=["jobs"])
def download_job_files(job_id: str):
    """
    Download tất cả file output của job dưới dạng ZIP.
    Chỉ available khi job status = done.
    """
    job = _get_job(job_id)
    if job.status != JobStatus.DONE:
        raise HTTPException(
            status_code=409,
            detail=f"Job chưa hoàn thành (status={job.status}). Chỉ download khi status=done."
        )

    output_paths: List[str] = job.result.get("output_paths", []) if job.result else []
    if not output_paths:
        raise HTTPException(status_code=404, detail="Không có file output")

    # Build ZIP in memory
    buf = io.BytesIO()
    missing = []
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for path_str in output_paths:
            p = Path(path_str)
            if p.exists():
                zf.write(p, arcname=p.name)
            else:
                missing.append(p.name)

    if missing:
        # Vẫn trả ZIP nhưng log cảnh báo
        pass

    buf.seek(0)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"kisordoc_{job_id[:8]}_{ts}.zip"

    return StreamingResponse(
        buf,
        media_type="application/zip",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.delete("/jobs/{job_id}", tags=["jobs"])
def delete_job(job_id: str):
    """Xóa job khỏi store (không xóa file output trên đĩa)."""
    with _store_lock:
        if job_id not in _job_store:
            raise HTTPException(status_code=404, detail=f"Job '{job_id}' không tồn tại")
        del _job_store[job_id]
    return {"deleted": job_id}


@app.get("/jobs", tags=["jobs"])
def list_jobs(limit: int = Query(20, le=100)):
    """Liệt kê N job gần nhất trong store (mới nhất trước)."""
    with _store_lock:
        jobs = list(reversed(list(_job_store.values())))[:limit]
    return [
        {
            "job_id":      j.job_id,
            "status":      j.status,
            "created_at":  j.created_at,
            "finished_at": j.finished_at,
            "option":      j.request.get("option", ""),
            "package_id":  j.request.get("package_id", ""),
            "templates":   j.request.get("templates", []),
        }
        for j in jobs
    ]


# ──────────────────────────────────────────────────────────────────────────────
# Entry point (chạy độc lập, không qua runner.py)
# ──────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "api:app",
        host="0.0.0.0",
        port=int(os.environ.get("API_PORT", 8000)),
        reload=False,
    )
