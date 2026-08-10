"""
kisorlib/migrator.py
─────────────────────
Core logic migrate template .docx:
  - Chuyển <<TenBien>>         → {{TenBien}}
  - Chuyển <<TenBien.Date>>    → {{TenBien_Date|date}}
  - Chuyển <<TenBien.Date.Long>> → {{TenBien_Date|date_long}}
  - Chuyển <<TenBien.Day/Month/Year>> → {{TenBien_Date|day/month/year}}
  - Chuyển <<TenBien.Number>>  → {{TenBien|number}}
  - Chuyển <<TenBien.Chu>>     → {{TenBien|num2text}}
  - Chuyển <<TenBien.Upper>>   → {{TenBien|upper}}
  - Chuẩn hóa {{TenBien|filter}} đã có
  - Chuyển {DanhMuc} (single-brace) → {{DanhMuc}}

Hỗ trợ:
  - dry_run=True : chỉ phân tích, không ghi file
  - on_progress   : callback(dict) cho UI streaming

Không import Gradio.
"""

from __future__ import annotations

import re
import shutil
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional


# ──────────────────────────────────────────────
# Regex
# ──────────────────────────────────────────────

# <<...>> trong XML (đã HTML-encoded)
_LT_GT_RE = re.compile(
    r"&lt;&lt;((?:(?!&gt;&gt;)(?!</?w:(?:p|tc|tr|tbl)\b).)*)&gt;&gt;",
    re.DOTALL,
)

# <<...>> dạng thô (nếu Word không encode — hiếm nhưng cần hỗ trợ)
# FIX: thêm '.' cuối mỗi alternation để consume ký tự, tránh match chuỗi rỗng vô hạn
_RAW_LT_GT_RE = re.compile(
    r"<<((?:(?!>>)(?!</?w:(?:p|tc|tr|tbl)\b).)*)>>",
    re.DOTALL,
)

# {{...}} đã có (để chuẩn hóa filter)
_DOUBLE_BRACE_RE = re.compile(r"\{\{(.*?)\}\}", re.DOTALL)

# {single} brace — table placeholder cũ
_SINGLE_BRACE_RE = re.compile(r"(?<!\{)\{([a-zA-Z0-9_]+)\}(?!\})")

# XML tag stripper
_XML_TAG_RE = re.compile(r"<[^>]+>")


# ──────────────────────────────────────────────
# Result types
# ──────────────────────────────────────────────

@dataclass
class PlaceholderChange:
    original: str
    converted: str
    location: str   # "word/document.xml" v.v.


@dataclass
class FileResult:
    path:         Path
    success:      bool
    changed:      bool             = False
    changes:      list[PlaceholderChange] = field(default_factory=list)
    error:        Optional[str]    = None
    backed_up_to: Optional[Path]   = None


OnProgress = Optional[Callable[[dict], None]]


def _emit(cb: OnProgress, level: str, message: str, **extra) -> None:
    if cb is None:
        return
    try:
        cb({"level": level, "message": message, **extra})
    except Exception:
        pass


# ──────────────────────────────────────────────
# Placeholder conversion
# ──────────────────────────────────────────────

def _strip_xml_tags(text: str) -> str:
    return _XML_TAG_RE.sub("", text)


def _map_placeholder(raw: str) -> tuple[str, str]:
    """
    Chuyển text placeholder thô (đã strip XML tags) → Jinja2 expression.
    Trả về (original_clean, converted).
    """
    val = raw.strip()

    # Normalize pipe-filter đã có
    if "|" in val:
        parts = val.split("|", 1)
        base = parts[0].strip()
        mod  = parts[1].strip().lower()
        return val, _apply_modifier(base, mod)

    # Suffix-based mapping (case-insensitive)
    lower = val.lower()
    suffixes = [
        (".date.long",  lambda b: _apply_modifier(b + "_Date", "date_long")),
        (".date_long",  lambda b: _apply_modifier(b + "_Date", "date_long")),
        (".date",       lambda b: _apply_modifier(b + "_Date", "date")),
        (".day",        lambda b: _apply_modifier(b + "_Date", "day")),
        (".month",      lambda b: _apply_modifier(b + "_Date", "month")),
        (".year",       lambda b: _apply_modifier(b + "_Date", "year")),
        (".number",     lambda b: _apply_modifier(b, "number")),
        (".chu",        lambda b: _apply_modifier(b, "num2text")),
        (".text",       lambda b: _apply_modifier(b, "num2text")),
        (".upper",      lambda b: _apply_modifier(b, "upper")),
    ]
    for suffix, builder in suffixes:
        if lower.endswith(suffix):
            base = val[: -len(suffix)]
            return val, builder(base)

    return val, f"{{{{{val}}}}}"


def _apply_modifier(base: str, mod: str) -> str:
    """
    Xây Jinja2 expression từ base variable name + modifier string.
    Tự động thêm _Date suffix cho date-family modifiers nếu chưa có.
    """
    date_mods = {"date", "date_long", "day", "month", "year"}
    if mod in date_mods:
        lower_base = base.lower()
        if not (lower_base.endswith("_date") or lower_base.endswith("_date")):
            base = base + "_Date"
    return f"{{{{{base}|{mod}}}}}"


# ──────────────────────────────────────────────
# XML-level migration
# ──────────────────────────────────────────────

def migrate_xml(xml: str) -> tuple[str, list[PlaceholderChange]]:
    """
    Áp dụng tất cả migration rules lên chuỗi XML.
    Trả về (xml_migrated, changes).
    """
    changes: list[PlaceholderChange] = []

    # 1. <<...>> HTML-encoded → {{...}}
    def _replace_lt_gt(m: re.Match) -> str:
        clean = _strip_xml_tags(m.group(1))
        orig, conv = _map_placeholder(clean)
        if f"<<{orig}>>" != conv:
            changes.append(PlaceholderChange(f"<<{orig}>>", conv, ""))
        return conv

    xml = _LT_GT_RE.sub(_replace_lt_gt, xml)

    # 2. <<...>> thô (nếu Word không encode)
    def _replace_raw_lt_gt(m: re.Match) -> str:
        clean = _strip_xml_tags(m.group(1))
        orig, conv = _map_placeholder(clean)
        if f"<<{orig}>>" != conv:
            changes.append(PlaceholderChange(f"<<{orig}>>", conv, ""))
        return conv

    xml = _RAW_LT_GT_RE.sub(_replace_raw_lt_gt, xml)

    # 3. {{...}} đã có — chuẩn hóa filter
    def _normalize_existing(m: re.Match) -> str:
        inner = _strip_xml_tags(m.group(1))
        orig_expr = f"{{{{{m.group(1)}}}}}"
        _, conv = _map_placeholder(inner)
        if orig_expr != conv:
            changes.append(PlaceholderChange(orig_expr, conv, ""))
        return conv

    xml = _DOUBLE_BRACE_RE.sub(_normalize_existing, xml)

    # 4. {single} brace → {{double}}
    def _replace_single(m: re.Match) -> str:
        name = m.group(1)
        conv = f"{{{{{name}}}}}"
        changes.append(PlaceholderChange(f"{{{name}}}", conv, ""))
        return conv

    xml = _SINGLE_BRACE_RE.sub(_replace_single, xml)

    return xml, changes


# ──────────────────────────────────────────────
# File-level migration
# ──────────────────────────────────────────────

_WORD_XML_PARTS = frozenset({
    "word/document.xml",
    "word/header1.xml", "word/header2.xml", "word/header3.xml",
    "word/footer1.xml", "word/footer2.xml", "word/footer3.xml",
})


def migrate_file(
    path:     Path,
    *,
    dry_run:     bool       = False,
    backup:      bool       = True,
    on_progress: OnProgress = None,
) -> FileResult:
    """
    Migrate 1 file .docx.
    dry_run=True: phân tích và trả kết quả, không ghi file.
    """
    result = FileResult(path=path, success=False)

    if not path.exists():
        result.error = f"File không tồn tại: {path}"
        _emit(on_progress, "error", result.error, file=path.name)
        return result

    try:
        all_changes: list[PlaceholderChange] = []

        with zipfile.ZipFile(path, "r") as zin:
            items = zin.infolist()
            xml_parts: dict[str, str] = {}
            for item in items:
                if item.filename.startswith("word/") and item.filename.endswith(".xml"):
                    xml_parts[item.filename] = zin.read(item.filename).decode("utf-8")

        migrated: dict[str, str] = {}
        for part_name, xml in xml_parts.items():
            new_xml, changes = migrate_xml(xml)
            for c in changes:
                c.location = part_name
            all_changes.extend(changes)
            migrated[part_name] = new_xml

        result.changes = all_changes
        result.changed = bool(all_changes)

        if dry_run or not all_changes:
            result.success = True
            if all_changes:
                _emit(on_progress, "info",
                      f"[dry-run] {path.name}: {len(all_changes)} thay đổi",
                      file=path.name, count=len(all_changes))
            else:
                _emit(on_progress, "info",
                      f"[dry-run] {path.name}: không có thay đổi",
                      file=path.name, count=0)
            return result

        # Backup
        if backup:
            bak_dir = path.parent / "bak"
            bak_dir.mkdir(exist_ok=True)
            bak_path = bak_dir / (path.stem + ".bak.docx")
            if not bak_path.exists():
                shutil.copy2(path, bak_path)
                result.backed_up_to = bak_path

        # Ghi file mới
        tmp_path = path.with_suffix(".tmp.docx")
        with zipfile.ZipFile(path, "r") as zin:
            with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    if item.filename in migrated:
                        zout.writestr(item, migrated[item.filename].encode("utf-8"))
                    else:
                        zout.writestr(item, zin.read(item.filename))

        tmp_path.replace(path)
        result.success = True
        _emit(on_progress, "success",
              f"✓ {path.name}: {len(all_changes)} thay đổi",
              file=path.name, count=len(all_changes))

    except Exception as e:
        result.error = f"{type(e).__name__}: {e}"
        result.success = False
        _emit(on_progress, "error",
              f"✗ {path.name}: {result.error}",
              file=path.name)
        tmp = path.with_suffix(".tmp.docx")
        if tmp.exists():
            try:
                tmp.unlink()
            except Exception:
                pass

    return result


# ──────────────────────────────────────────────
# Folder-level migration
# ──────────────────────────────────────────────

def migrate_folder(
    folder:      Path,
    *,
    dry_run:     bool       = False,
    backup:      bool       = True,
    recursive:   bool       = True,
    on_progress: OnProgress = None,
) -> list[FileResult]:
    """
    Migrate tất cả .docx trong folder (bỏ qua bak/ và *.bak.docx).
    """
    pattern = "**/*.docx" if recursive else "*.docx"
    files = [
        p for p in sorted(folder.glob(pattern))
        if "bak" not in p.parts
        and not p.name.endswith(".bak.docx")
        and not p.name.startswith("~")     # Word temp files
    ]

    if not files:
        _emit(on_progress, "warning", f"Không tìm thấy file .docx trong: {folder}")
        return []

    _emit(on_progress, "info",
          f"Tìm thấy {len(files)} file .docx",
          total=len(files))

    results = []
    for i, f in enumerate(files, 1):
        _emit(on_progress, "info",
              f"[{i}/{len(files)}] {f.name}",
              step=i, total=len(files), file=f.name)
        r = migrate_file(f, dry_run=dry_run, backup=backup, on_progress=on_progress)
        results.append(r)

    ok       = sum(1 for r in results if r.success and r.changed)
    no_chg   = sum(1 for r in results if r.success and not r.changed)
    errors   = sum(1 for r in results if not r.success)
    total_ch = sum(len(r.changes) for r in results)

    _emit(on_progress, "success" if not errors else "warning",
          f"Xong: {ok} migrate, {no_chg} không đổi, {errors} lỗi — {total_ch} placeholder đã chuyển")

    return results


# ──────────────────────────────────────────────
# Summary formatter (dùng bởi UI)
# ──────────────────────────────────────────────

def format_summary(results: list[FileResult], dry_run: bool = False) -> str:
    lines = []
    mode = "[DRY-RUN] " if dry_run else ""
    for r in results:
        if not r.success:
            lines.append(f"❌ {r.path.name}: {r.error}")
        elif not r.changed:
            lines.append(f"⬜ {r.path.name}: không có placeholder cũ")
        else:
            lines.append(f"{'🔍' if dry_run else '✅'} {r.path.name}: {len(r.changes)} thay đổi")
            for c in r.changes[:10]:   # tối đa 10 dòng mỗi file
                lines.append(f"     {c.original}  →  {c.converted}")
            if len(r.changes) > 10:
                lines.append(f"     ... và {len(r.changes) - 10} thay đổi khác")
    return "\n".join(lines) if lines else "(Không có file nào)"
