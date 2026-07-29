import os
import shutil
import re
import time
import subprocess
from pathlib import Path

from config import AppConfig


def clear_output_folder(config: AppConfig):
    """Clear output folder with Word process killing and retry logic"""
    output = config.output_path
    
    # Force close Word processes to avoid file locks
    try:
        subprocess.run(['taskkill', '/F', '/IM', 'WINWORD.EXE'], 
                      capture_output=True, timeout=5)
        time.sleep(1)  # Wait for processes to fully terminate
    except Exception:
        pass  # If taskkill fails, continue anyway
    
    if not output.exists():
        output.mkdir(parents=True, exist_ok=True)
        return
    
    # Try to remove with retries
    max_retries = 5
    retry_delay = 1.0  # Increased to 1 second
    
    for attempt in range(max_retries):
        try:
            shutil.rmtree(output)
            output.mkdir(parents=True, exist_ok=True)
            return
        except PermissionError as e:
            if attempt < max_retries - 1:
                print(f"⚠️  Thư mục bị khóa, thử lại ({attempt + 1}/{max_retries})...")
                time.sleep(retry_delay)
            else:
                print(f"❌ Không thể xóa thư mục sau {max_retries} lần: {e}")
                print(f"💡 Vui lòng đóng tất cả file Word và thử lại")
                raise


def copy_templates_to_output(config: AppConfig, option: str, filenames: list[str]) -> list[Path]:
    src_dir = config.template_path / option
    if not src_dir.exists():
        raise FileNotFoundError(f"Template directory not found: {src_dir}")

    copied: list[Path] = []
    for name in filenames:
        src = src_dir / name
        if not src.exists():
            possible = list(src_dir.glob(f"{name}*"))
            if possible:
                src = possible[0]
            else:
                continue
        dst = config.output_path / src.name
        
        # FIX F6-07: Tự động retry khi copy template bị locked
        max_retries = config.FileMaxRetries
        delay = config.FileRetryDelay
        copied_ok = False
        for attempt in range(1, max_retries + 1):
            try:
                shutil.copy2(src, dst)
                copied_ok = True
                break
            except PermissionError as pe:
                is_lock = (pe.errno == 13 or "permission denied" in str(pe).lower() or "being used" in str(pe).lower())
                if is_lock and attempt < max_retries:
                    print(f"⚠️ Template {src.name} bị khóa, đang thử lại lần {attempt}/{max_retries} sau {delay:.1f}s...")
                    time.sleep(delay)
                else:
                    raise pe
        if copied_ok:
            copied.append(dst)
    return copied


def rename_output(file_path: Path, goi_thau_id: str, used_names: set[str]) -> Path:
    stem = file_path.stem
    if stem.endswith("-Template"):
        stem = stem[: -len("-Template")]
    new_name = f"{stem}-{goi_thau_id}.docx"
    new_name = _sanitize_filename(new_name)
    if new_name in used_names:
        counter = 1
        while True:
            candidate = f"{stem}-{goi_thau_id}_{counter}.docx"
            candidate = _sanitize_filename(candidate)
            if candidate not in used_names:
                new_name = candidate
                break
            counter += 1
    used_names.add(new_name)
    new_path = file_path.parent / new_name
    os.rename(file_path, new_path)
    return new_path


def _sanitize_filename(name: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "", name)


def open_output_folder(config: AppConfig):
    try:
        path = str(config.output_path.resolve())
        if os.path.exists(path):
            os.system(f'start "" "{path}"')
        else:
            print(f"Output folder does not exist: {path}")
    except Exception as e:
        print(f"Error opening output folder: {e}")
