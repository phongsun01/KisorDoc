import os
import shutil
import re
from pathlib import Path

from config import AppConfig


def clear_output_folder(config: AppConfig):
    output = config.output_path
    if output.exists():
        shutil.rmtree(output)
    output.mkdir(parents=True, exist_ok=True)


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
        shutil.copy2(src, dst)
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
