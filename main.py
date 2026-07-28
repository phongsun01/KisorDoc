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
from merger import mail_merge
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
    opt_val = option_key.split(":")[0].strip() if ":" in option_key else option_key
    goi_thau_rows = ds.query("SELECT * FROM GoiThau")
    packages = []
    for r in goi_thau_rows:
        label = f"{_str(r.get('TT'))}. {_str(r.get('Số hiệu gói thầu'))} - {_str(r.get('Tên gói thầu'))}"
        packages.append(label)
    return packages
    if val is None:
        return ""
    if isinstance(val, float):
        import math
        if math.isnan(val):
            return ""
    return str(val).strip()

def get_workflow_templates(option_key: str, package_label: str) -> list[dict]:
    if not option_key:
        return []
    opt = option_key.split(":")[0].strip() if ":" in option_key else option_key

    if not package_label or package_label.strip() == "":
        return []

    ws_rows = ds.query("SELECT * FROM Workflow")
    wf_rows = [r for r in ws_rows if _str(r.get("Option")) == opt or not _str(r.get("Option"))]
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


async def run_batch(option_key: str, package_label: str, selected_templates: list[str],
                    progress: gr.Progress = gr.Progress()):
    config = load_config()
    ds_local = DataSet(config)

    opt = option_key.split(":")[0].strip() if ":" in option_key else option_key

    goi_thau_rows = ds_local.query("SELECT * FROM GoiThau")
    selected_pkg = None
    for r in goi_thau_rows:
        label = f"{r.get('TT', '')}. {r.get('Số hiệu gói thầu', '')} - {r.get('Tên gói thầu', '')}"
        if label == package_label:
            selected_pkg = r
            break

    if not selected_pkg:
        yield [], "❌ Không tìm thấy gói thầu đã chọn"
        return

    goi_thau_id = str(selected_pkg.get("GoiThau_ID", ""))

    config_rows = ds_local.query("SELECT * FROM Config")
    config_map = {}
    for r in config_rows:
        key = str(r.get("Key", "")).strip()
        value = str(r.get("Value", "")).strip()
        if key:
            config_map[key] = value

    context = {}
    for key, col in config_map.items():
        raw_value = selected_pkg.get(col, "")
        if isinstance(raw_value, datetime):
            raw_value = raw_value.strftime("%d/%m/%Y")
        elif raw_value is None:
            raw_value = ""
        context[key] = str(raw_value)

    tables_rows = ds_local.query("SELECT * FROM Tables")
    xlsx_files = sorted(config.data_path.glob("*.xlsx"))
    danh_muc_file = None
    for f in xlsx_files:
        if "DanhMuc" in f.stem or "danh muc" in f.stem.lower():
            danh_muc_file = f
            break
    if danh_muc_file is None and xlsx_files:
        danh_muc_file = xlsx_files[0]

    progress(0, desc="Bắt đầu xử lý...")
    yield [], "Bắt đầu..."

    clear_output_folder(config)

    template_filenames = []
    template_names = []
    for r in get_workflow_templates(option_key, package_label):
        if r.get("Name", "") in selected_templates:
            fname_raw = str(r.get("File", "")).strip()
            fname = fname_raw + ".docx" if not fname_raw.endswith(".docx") else fname_raw
            template_filenames.append(fname)
            template_names.append(r.get("Name", ""))

    copied = copy_templates_to_output(config, opt, template_filenames)
    total = len(copied)
    results = []
    used_names = set()
    start_time = time.time()

    for i, (src_path, tpl_name) in enumerate(zip(copied, template_names)):
        progress((i + 1) / total, desc=f"Đang xử lý: {tpl_name}")

        try:
            mail_merge(src_path, context, src_path)
            if danh_muc_file and danh_muc_file.exists():
                copy_tables_for_file(src_path, config, goi_thau_id, tables_rows, danh_muc_file)

            new_path = rename_output(src_path, goi_thau_id, used_names)
            results.append(f"✅ {tpl_name} → {new_path.name}")
        except Exception as e:
            tb = traceback.format_exc()
            results.append(f"❌ {tpl_name}: {e}")
            yield results, f"Lỗi tại {tpl_name}: {e}"

    elapsed = time.time() - start_time
    summary = f"Hoàn thành {len([r for r in results if r.startswith('✅')])}/{total} file trong {elapsed:.1f}s"
    yield results, summary


def create_ui():
    init()

    option_state = gr.State("")
    package_state = gr.State("")

    with gr.Blocks(title=config.AppName) as app:
        gr.Markdown(f"# {config.AppName} – Xử lý Word tự động")

        with gr.Tab("1. Chọn quy trình & Gói thầu"):
            options = get_options()
            option_radio = gr.Radio(choices=options, label="Chọn quy trình")

            goi_thau_rows = ds.query("SELECT * FROM GoiThau")
            packages = []
            for r in goi_thau_rows:
                packages.append(f"{_str(r.get('TT'))}. {_str(r.get('Số hiệu gói thầu'))} - {_str(r.get('Tên gói thầu'))}")
            package_radio = gr.Radio(choices=packages, label="Chọn gói thầu")

            submit_btn = gr.Button("📥 Tiếp theo", variant="primary")

        with gr.Tab("2. Chọn file template"):
            template_checkboxes = gr.CheckboxGroup(label="Chọn template cần xử lý", choices=[])
            select_all_btn = gr.Button("Chọn tất cả")
            deselect_all_btn = gr.Button("Bỏ chọn tất cả")
            run_btn = gr.Button("🚀 Chạy", variant="primary")

        with gr.Tab("3. Log & Kết quả"):
            result_log = gr.Dataframe(headers=["Kết quả"], label="Chi tiết")
            status_text = gr.Textbox(label="Trạng thái", interactive=False)
            open_folder_btn = gr.Button("📂 Mở thư mục output")

        def on_submit_load_templates(opt, pkg):
            templates = get_workflow_templates(opt, pkg)
            choices = [t.get("Name", "") for t in templates]
            return gr.update(choices=choices, value=[])

        submit_btn.click(
            fn=on_submit_load_templates,
            inputs=[option_radio, package_radio],
            outputs=[template_checkboxes],
        )

        def select_all(choices):
            return gr.update(value=choices)

        def deselect_all():
            return gr.update(value=[])

        select_all_btn.click(fn=select_all, inputs=[template_checkboxes], outputs=[template_checkboxes])
        deselect_all_btn.click(fn=deselect_all, outputs=[template_checkboxes])

        run_event = run_btn.click(
            fn=run_batch,
            inputs=[option_radio, package_radio, template_checkboxes],
            outputs=[result_log, status_text],
        )

        def on_open_folder():
            open_output_folder(config)

        open_folder_btn.click(fn=on_open_folder)

    return app


if __name__ == "__main__":
    app = create_ui()
    import webbrowser
    PORT = 7863
    threading.Thread(target=lambda: webbrowser.open(f"http://127.0.0.1:{PORT}"), daemon=True).start()
    print(f"KisorDoc-AI running at http://127.0.0.1:{PORT}")
    app.launch(server_port=PORT, share=False, quiet=True, inbrowser=False)
