import asyncio
import sys
import threading
import time
import traceback
from datetime import datetime
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")

import gradio as gr

from config import load_config, AppConfig
from dataset import DataSet
from file_utils import clear_output_folder, copy_templates_to_output, rename_output, open_output_folder
from merger import mail_merge_safe
from table_copier import copy_tables_for_file, TABLE_PLACEHOLDER_RE

config: AppConfig | None = None
ds: DataSet | None = None


def _str(val, default=""):
    if val is None:
        return default
    if isinstance(val, float):
        import math
        if math.isnan(val):
            return default
    return str(val).strip()


def init():
    global config, ds
    config = load_config()
    ds = DataSet(config)


def get_options() -> list[str]:
    rows = ds.query("SELECT Key, Value FROM Options ORDER BY Key")
    return [f"{_str(r['Key'])}: {_str(r['Value'])}" for r in rows]


def get_packages(option_key: str) -> list[str]:
    """FIX #5: Filter packages by option, sorted by TT"""
    opt_val = option_key.split(":")[0].strip() if ":" in option_key else option_key
    goi_thau_rows = ds.query("SELECT * FROM GoiThau ORDER BY CAST(TT AS INTEGER)")
    packages = []
    for r in goi_thau_rows:
        label = f"{_str(r.get('TT'))}. {_str(r.get('Số hiệu gói thầu'))} - {_str(r.get('Tên gói thầu'))}"
        packages.append(label)
    return packages


def get_package_details(package_label: str) -> dict:
    """FIX #9: Get package details for preview"""
    if not package_label:
        return {}
    goi_thau_rows = ds.query("SELECT * FROM GoiThau")
    for r in goi_thau_rows:
        label = f"{_str(r.get('TT'))}. {_str(r.get('Số hiệu gói thầu'))} - {_str(r.get('Tên gói thầu'))}"
        if label == package_label:
            return {
                "Tên CĐT": _str(r.get("Tên CĐT", "")),
                "Giá": _str(r.get("Giá gói thầu", "")),
                "Loại": _str(r.get("GoiThau_Loai", "")),
                "Số hiệu": _str(r.get("Số hiệu gói thầu", "")),
            }
    return {}

def get_workflow_templates(option_key: str, package_label: str) -> list[dict]:
    if not option_key:
        return []
    opt = option_key.split(":")[0].strip() if ":" in option_key else option_key

    if not package_label or package_label.strip() == "":
        return []

    ws_rows = ds.query("SELECT * FROM Workflow")
    wf_rows = [r for r in ws_rows if _str(r.get("Option")) == opt]
    if not wf_rows:
        return []

    goi_thau_rows = ds.query("SELECT * FROM GoiThau")
    selected_pkg = None
    for r in goi_thau_rows:
        label = f"{_str(r.get('TT'))}. {_str(r.get('Số hiệu gói thầu'))} - {_str(r.get('Tên gói thầu'))}"
        if label == package_label:
            selected_pkg = r
            break

    if not selected_pkg:
        return wf_rows

    price = _parse_price(selected_pkg.get("Giá gói thầu", 0))
    goi_thau_loai = _str(selected_pkg.get("GoiThau_Loai"))

    filtered = []
    for r in wf_rows:
        pmin = _parse_price(r.get("Price", 0))
        pmax = _parse_price(r.get("PriceMax", 0))
        if pmin is not None and pmax is not None and price is not None:
            if not (pmin <= price <= pmax):
                continue
        rtype = _str(r.get("Type"))
        if rtype != "ALL" and rtype != goi_thau_loai:
            continue
        filtered.append(r)

    return filtered


def _parse_price(val) -> float | None:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        import math
        if math.isnan(val):
            return None
        return float(val)
    s = str(val).strip()
    if not s:
        return None
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def make_nested_dict(flat_dict: dict) -> dict:
    nested = {}
    for key, value in flat_dict.items():
        parts = key.split(".")
        d = nested
        for part in parts[:-1]:
            if part not in d:
                d[part] = {}
            elif not isinstance(d[part], dict):
                # Fallback if there is a clash
                d[part] = {"_val": d[part]}
            d = d[part]
        d[parts[-1]] = value
    return nested


async def run_batch(option_key: str, package_label: str, selected_templates: list[str],
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
    
    # Query Tables sheet - check if it exists first
    try:
        tables_rows = ds.query("SELECT * FROM Tables")   # FIX 2
    except Exception:
        tables_rows = []  # If Tables sheet doesn't exist, skip table copying

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
        
        # Handle pandas NaT first (before checking isinstance datetime)
        import pandas as pd
        if pd.isna(raw_value):
            raw_value = ""
        elif isinstance(raw_value, datetime):
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
                try:
                    copy_tables_for_file(src_path, config, goi_thau_id, tables_rows, danh_muc_file)
                except Exception as table_err:
                    print(f"⚠️  Lỗi copy bảng: {table_err}")

            new_path = rename_output(src_path, goi_thau_id, used_names)
            results.append(f"✅ {tpl_name} → {new_path.name}")
        except Exception as e:
            results.append(f"❌ {tpl_name}: {e}")

        yield results, f"Đang xử lý {i + 1}/{total}..."

    elapsed = time.time() - start_time
    ok_count = sum(1 for r in results if r.startswith("✅"))
    summary = f"Hoàn thành {ok_count}/{total} file trong {elapsed:.1f}s"
    yield results, summary




def create_ui():
    init()

    # Dùng dict mutable để lưu lựa chọn hiện tại
    _sel = {"opt": "", "pkg": ""}

    with gr.Blocks(title=config.AppName) as app:
        gr.Markdown(f"# {config.AppName} – Xử lý Word tự động")

        # Tab 1 & 2: Merged into 2-column layout
        with gr.Tab("1. Chọn & Chạy"):
            with gr.Row():
                # Column 1: Tab 1 - Chọn quy trình & Gói thầu
                with gr.Column(scale=1):
                    gr.Markdown("### Chọn quy trình & Gói thầu")
                    options = get_options()
                    option_radio = gr.Radio(choices=options, label="Chọn quy trình")

                    goi_thau_rows = ds.query("SELECT * FROM GoiThau ORDER BY CAST(TT AS INTEGER)")
                    packages = []
                    for r in goi_thau_rows:
                        packages.append(f"{_str(r.get('TT'))}. {_str(r.get('Số hiệu gói thầu'))} - {_str(r.get('Tên gói thầu'))}")
                    package_radio = gr.Radio(choices=packages, label="Chọn gói thầu")

                    # FIX #9: Add package preview
                    with gr.Group():
                        gr.Markdown("**Preview thông tin gói thầu:**")
                        pkg_preview = gr.Textbox(label="", interactive=False, max_lines=4)

                # Column 2: Tab 2 - Chọn file template & Chạy
                with gr.Column(scale=1):
                    gr.Markdown("### Chọn file template & Chạy")
                    # FIX #10: Add count to checkbox label
                    template_label = gr.Markdown("**Chọn template cần xử lý** (0 file)")
                    template_checkboxes = gr.CheckboxGroup(label="", choices=[])
                    
                    with gr.Row():
                        select_all_btn = gr.Button("✓ Chọn tất cả")
                        deselect_all_btn = gr.Button("✗ Bỏ chọn tất cả")
                    
                    run_btn = gr.Button("🚀 Chạy", variant="primary", size="lg")

        # Tab 3: Log & Kết quả
        with gr.Tab("2. Log & Kết quả"):
            # FIX #4: Replace Dataframe with Textbox for better log display
            result_log = gr.Textbox(
                label="Chi tiết kết quả",
                interactive=False,
                lines=15,
                max_lines=20,
            )
            status_text = gr.Textbox(label="Trạng thái", interactive=False)
            
            with gr.Row():
                # FIX #12: Open folder button only shown after run
                open_folder_btn = gr.Button("📂 Mở thư mục output", visible=False)
                rerun_btn = gr.Button("← Chạy lại", variant="secondary")

        # FIX #1: Auto-load templates when package is selected (no submit button needed)
        def on_package_change(pkg):
            details = get_package_details(pkg)
            if not details:
                preview_text = ""
            else:
                lines = [f"{k}: {v}" for k, v in details.items() if v]
                preview_text = "\n".join(lines)
            
            # Auto-load templates when package is selected
            opt = _sel["opt"]
            if not opt or not pkg:
                return preview_text, gr.update(choices=[], value=[]), gr.update(value="**Chọn template cần xử lý** (0 file)")
            
            templates = get_workflow_templates(opt, pkg)
            choices = [t.get("Name", "") for t in templates]
            label_text = f"**Chọn template cần xử lý** ({len(choices)} file)"
            _sel["pkg"] = pkg
            
            return preview_text, gr.update(choices=choices, value=[]), gr.update(value=label_text)

        package_radio.change(
            fn=on_package_change,
            inputs=[package_radio],
            outputs=[pkg_preview, template_checkboxes, template_label]
        )

        # When option changes, clear package and templates
        def on_option_change(opt):
            _sel["opt"] = opt or ""
            return (
                gr.update(value=None),  # Reset package
                gr.update(choices=[], value=[]),  # Clear templates
                gr.update(value="**Chọn template cần xử lý** (0 file)"),  # Reset label
                ""  # Clear preview
            )

        option_radio.change(
            fn=on_option_change,
            inputs=[option_radio],
            outputs=[package_radio, template_checkboxes, template_label, pkg_preview]
        )

        # FIX #10: Update label when checkbox changes
        def update_checkbox_label():
            total = len(get_workflow_templates(_sel["opt"], _sel["pkg"]))
            return gr.update(value=f"**Chọn template cần xử lý** ({total} file)")

        template_checkboxes.change(fn=update_checkbox_label, outputs=[template_label])

        def select_all():
            opt, pkg = _sel["opt"], _sel["pkg"]
            if not opt or not pkg:
                return gr.update(value=[])
            templates = get_workflow_templates(opt, pkg)
            choices = [t.get("Name", "") for t in templates]
            return gr.update(value=choices)

        def deselect_all():
            return gr.update(value=[])

        select_all_btn.click(fn=select_all, outputs=[template_checkboxes])
        deselect_all_btn.click(fn=deselect_all, outputs=[template_checkboxes])

        # FIX #2: Validation before run
        def run_with_validation(opt, pkg, templates):
            # Validate inputs
            if not opt or not opt.strip():
                return "", "❌ Vui lòng chọn quy trình", gr.update(visible=False)
            if not pkg or not pkg.strip():
                return "", "❌ Vui lòng chọn gói thầu", gr.update(visible=False)
            if not templates or len(templates) == 0:
                return "", "❌ Vui lòng chọn ít nhất 1 template", gr.update(visible=False)
            
            return None, None, gr.update(visible=False)  # Trigger actual run_batch

        # FIX #3: Disable button during processing (use trigger_mode="once")
        run_event = run_btn.click(
            fn=run_with_validation,
            inputs=[option_radio, package_radio, template_checkboxes],
            outputs=[result_log, status_text, open_folder_btn],
            trigger_mode="once",
        ).then(
            fn=run_batch,
            inputs=[option_radio, package_radio, template_checkboxes],
            outputs=[result_log, status_text],
        ).then(
            lambda: gr.update(visible=True),
            outputs=[open_folder_btn]
        )

        # FIX #2: Open folder - add error handling
        def on_open_folder():
            try:
                open_output_folder(config)
                return "✅ Thư mục đã mở"
            except Exception as e:
                return f"❌ Lỗi mở thư mục: {e}"

        open_folder_btn.click(fn=on_open_folder, outputs=[status_text])

        # FIX #3: Rerun button - reset Radio values properly
        def on_rerun():
            _sel["opt"] = ""
            _sel["pkg"] = ""
            # Get initial choices for radio buttons
            initial_options = get_options()
            initial_packages = []
            for r in ds.query("SELECT * FROM GoiThau ORDER BY CAST(TT AS INTEGER)"):
                initial_packages.append(f"{_str(r.get('TT'))}. {_str(r.get('Số hiệu gói thầu'))} - {_str(r.get('Tên gói thầu'))}")
            
            return (
                gr.update(choices=initial_options, value=None),  # Reset to None instead of ""
                gr.update(choices=initial_packages, value=None),  # Reset to None instead of ""
                gr.update(value=[]),
                gr.update(value=""),
                gr.update(visible=False),
                gr.update(value="")
            )

        rerun_btn.click(
            fn=on_rerun,
            outputs=[option_radio, package_radio, template_checkboxes, pkg_preview, open_folder_btn, result_log],
        )

    return app


if __name__ == "__main__":
    app = create_ui()
    import webbrowser
    PORT = 7864
    # Try to find an available port
    import socket
    while True:
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.bind(('127.0.0.1', PORT))
            sock.close()
            break
        except OSError:
            PORT += 1
    
    threading.Thread(target=lambda: webbrowser.open(f"http://127.0.0.1:{PORT}"), daemon=True).start()
    print(f"KisorDoc-AI running at http://127.0.0.1:{PORT}")
    app.launch(server_port=PORT, share=False, quiet=True, inbrowser=False)
