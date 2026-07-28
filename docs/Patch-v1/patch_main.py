"""
PATCH: main.py
BUG FIX:
  1. Orphan code _str() lần 2 ở dòng ~56-60 → xóa
  2. run_batch() tạo DataSet mới mỗi lần chạy → dùng global ds
  3. merger.py thiếu error handling → wrap try/except trước khi save

Chỉ liệt kê các đoạn cần thay đổi (diff-style), không rewrite toàn bộ file.
"""

# ─────────────────────────────────────────────
# FIX 1: XÓA đoạn orphan code sau dòng get_packages()
# Khoảng dòng 56–64 trong file gốc:
#
#   REMOVE THIS BLOCK (nằm ngoài bất kỳ function nào):
#   ┌──────────────────────────────────────────
#   │ if val is None:
#   │     return ""
#   │ if isinstance(val, float):
#   │     import math
#   │     if math.isnan(val):
#   │         return ""
#   │ return str(val).strip()
#   └──────────────────────────────────────────
# ─────────────────────────────────────────────


# ─────────────────────────────────────────────
# FIX 2: run_batch() dùng global ds thay vì tạo mới
# ─────────────────────────────────────────────

# TRƯỚC (tạo DataSet mới mỗi lần):
# async def run_batch(...):
#     config = load_config()
#     ds_local = DataSet(config)      ← chậm, load lại toàn bộ Excel
#     ...
#     goi_thau_rows = ds_local.query(...)
#     config_rows = ds_local.query(...)
#     tables_rows = ds_local.query(...)

# SAU (dùng global ds đã init sẵn):
async def run_batch_fixed(option_key: str, package_label: str, selected_templates: list[str],
                          progress=None):
    """
    Version đã fix: dùng global config + ds thay vì khởi tạo lại.
    Copy hàm này vào main.py, thay thế run_batch() cũ.
    """
    import time, traceback
    from datetime import datetime
    import gradio as gr
    from file_utils import clear_output_folder, copy_templates_to_output, rename_output, open_output_folder
    from merger import mail_merge_safe  # dùng version có error handling (FIX 3)
    from table_copier import copy_tables_for_file

    global config, ds  # FIX 2: dùng global thay vì tạo mới

    opt = option_key.split(":")[0].strip() if ":" in option_key else option_key

    goi_thau_rows = ds.query("SELECT * FROM GoiThau")
    selected_pkg = None
    for r in goi_thau_rows:
        label = f"{_str(r.get('TT'))}. {_str(r.get('Số hiệu gói thầu'))} - {_str(r.get('Tên gói thầu'))}"
        if label == package_label:
            selected_pkg = r
            break

    if not selected_pkg:
        yield [], "❌ Không tìm thấy gói thầu đã chọn"
        return

    goi_thau_id = _str(selected_pkg.get("GoiThau_ID"))
    config_rows = ds.query("SELECT * FROM Config")   # FIX 2: ds thay vì ds_local
    tables_rows = ds.query("SELECT * FROM Tables")   # FIX 2

    context = {}
    for r in config_rows:
        key = _str(r.get("Key"))
        col = _str(r.get("Value"))
        if not key or not col:
            continue
        clean_key = key.strip("<>{}| ")
        # Normalize modifier suffix
        for suffix in (".Date.Long", ".Date", ".Upper", ".Number"):
            if clean_key.endswith(suffix):
                clean_key = clean_key[: -len(suffix)]
                if suffix == ".Date.Long":
                    clean_key += "_Date"
                elif suffix == ".Date":
                    clean_key += "_Date"
                break
        if "|" in clean_key:
            clean_key = clean_key.split("|")[0].strip()

        raw_value = selected_pkg.get(col, "")
        if isinstance(raw_value, datetime):
            raw_value = raw_value.strftime("%d/%m/%Y")
        elif raw_value is None:
            raw_value = ""
        context[clean_key] = str(raw_value)

    nested_context = make_nested_dict(context)

    xlsx_files = sorted(config.data_path.glob("*.xlsx"))
    danh_muc_file = next(
        (f for f in xlsx_files if "DanhMuc" in f.stem or "danh muc" in f.stem.lower()),
        xlsx_files[0] if xlsx_files else None
    )

    if progress:
        progress(0, desc="Bắt đầu xử lý...")
    yield [], "Bắt đầu..."

    clear_output_folder(config)

    template_filenames, template_names = [], []
    for r in get_workflow_templates(option_key, package_label):
        if r.get("Name", "") in selected_templates:
            fname_raw = _str(r.get("File", ""))
            fname = fname_raw if fname_raw.endswith(".docx") else fname_raw + ".docx"
            template_filenames.append(fname)
            template_names.append(r.get("Name", ""))

    copied = copy_templates_to_output(config, opt, template_filenames)
    total = len(copied)
    results = []
    used_names: set[str] = set()
    start_time = time.time()

    for i, (src_path, tpl_name) in enumerate(zip(copied, template_names)):
        if progress:
            progress((i + 1) / total, desc=f"Đang xử lý: {tpl_name}")
        try:
            # FIX 3: dùng mail_merge_safe có error handling
            ok, err = mail_merge_safe(src_path, nested_context, src_path)
            if not ok:
                raise RuntimeError(err)

            if danh_muc_file and danh_muc_file.exists():
                copy_tables_for_file(src_path, config, goi_thau_id, tables_rows, danh_muc_file)

            new_path = rename_output(src_path, goi_thau_id, used_names)
            results.append(f"✅ {tpl_name} → {new_path.name}")
        except Exception as e:
            results.append(f"❌ {tpl_name}: {e}")

        yield results, f"Đang xử lý {i + 1}/{total}..."

    elapsed = time.time() - start_time
    ok_count = sum(1 for r in results if r.startswith("✅"))
    summary = f"Hoàn thành {ok_count}/{total} file trong {elapsed:.1f}s"
    yield results, summary


# ─────────────────────────────────────────────
# FIX 3: merger.py — thêm mail_merge_safe() với error handling
# Thêm hàm này vào merger.py:
# ─────────────────────────────────────────────

MERGER_PATCH = '''
def mail_merge_safe(template_path, context: dict, output_path) -> tuple[bool, str]:
    """
    FIX 3: Version có error handling — không corrupt file nếu Jinja2 lỗi.
    Returns (success: bool, error_message: str)
    """
    import tempfile, shutil
    from pathlib import Path
    tmp = Path(tempfile.mktemp(suffix=".docx"))
    try:
        doc = DocxTemplate(str(template_path))
        jenv = jinja2.Environment()
        jenv.filters["date"]      = filter_date
        jenv.filters["date_long"] = filter_date_long
        jenv.filters["number"]    = filter_number
        doc.render(context, jenv)
        doc.save(str(tmp))
        # Chỉ ghi đè file gốc sau khi render thành công
        shutil.move(str(tmp), str(output_path))
        return True, ""
    except Exception as e:
        if tmp.exists():
            tmp.unlink()
        return False, str(e)
    finally:
        if tmp.exists():
            try:
                tmp.unlink()
            except Exception:
                pass
'''

print("Xem nội dung MERGER_PATCH để thêm vào merger.py")
