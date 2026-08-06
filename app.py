import asyncio
import os
import sys
import traceback
from pathlib import Path
import gradio as gr

from kisorlib.config import load_config
from kisorlib.dataset import DataSet
from kisorlib.service import KisorService
from kisorlib.batch import run_batch, run_retry_batch
from kisorlib.file_utils import cleanup_old_logs
from kisorlib.utils import _str, _parse_repeat_key_id, _parse_repeat_sheet_config, resolve_sheet_query, safe_format

ui_labels: dict = {}


def init() -> KisorService:
    global ui_labels
    config = load_config()
    ds = DataSet(config)
    
    # Nhân bản các bảng để lưu dữ liệu gốc tránh bị ghi đè khi đăng ký bảng tạm
    for tbl in list(ds.table_names):
        try:
            ds.conn.execute(f'DROP TABLE IF EXISTS "{tbl}_Goc"')
            ds.conn.execute(f'CREATE TABLE "{tbl}_Goc" AS SELECT * FROM "{tbl}"')
        except Exception as e:
            print(f"⚠️ Không thể tạo bảng gốc sao lưu cho {tbl}: {e}")
            
    cleanup_old_logs(config)
    try:
        import json
        with open("ui_labels.json", "r", encoding="utf-8") as f:
            ui_labels = json.load(f)
    except Exception as e:
        print(f"⚠️ Không load được ui_labels.json: {e}")
        ui_labels = {}
        
    service = KisorService(config, ds)
    return service


def create_ui():
    service = init()

    _sel = {"opt": "", "pkg": "", "sheet_rows": [], "template_total": 0, "choices": []}

    with gr.Blocks(title=ui_labels.get("app_title", "KisorDoc-AI")) as app:
        gr.Markdown(ui_labels.get("app_title", "KisorDoc-AI – Xử lý tài liệu tự động"))

        last_run_state = gr.State(None)

        with gr.Tabs() as tabs:
            with gr.Tab("1. Chọn & Chạy", id=0):
                with gr.Row():
                    with gr.Column(scale=1):
                        gr.Markdown("### " + ui_labels.get("workflow_section", "Chọn Quy trình"))
                        options = service.get_options()
                        option_radio = gr.Radio(choices=options, label=ui_labels.get("workflow_section", "Chọn Quy trình"))
                        package_radio = gr.Radio(choices=[], label=ui_labels.get("package_section", "Chọn Dữ liệu"))
                        _default_group = ui_labels.get("repeat_group_choices", [])
                        group_radio = gr.Radio(
                            choices=_default_group,
                            label=ui_labels.get("repeat_group_title", "Chọn Nhóm lặp"),
                            visible=False,
                            value=_default_group[0] if _default_group else None
                        )

                        with gr.Group():
                            gr.Markdown("**" + ui_labels.get("preview_section", "Preview thông tin:") + "**")
                            pkg_preview = gr.Textbox(label="", interactive=False, max_lines=4)

                    with gr.Column(scale=1) as template_col:
                        gr.Markdown("### " + ui_labels.get("template_section", "Chọn file template & Chạy"))
                        template_label = gr.Markdown(f'**{ui_labels.get("template_section", "Chọn file template & Chạy")}** (0 file)', visible=True)
                        template_search = gr.Textbox(placeholder="🔍 Tìm kiếm nhanh...", label="", show_label=False, visible=True)
                        template_checkboxes = gr.CheckboxGroup(label="", choices=[], visible=True)

                        with gr.Row():
                            select_all_btn = gr.Button(ui_labels.get("select_all_btn", "✓ Chọn tất cả"), visible=True)
                            deselect_all_btn = gr.Button(ui_labels.get("deselect_all_btn", "✗ Bỏ chọn tất cả"), visible=True)

                        with gr.Row():
                            run_btn = gr.Button(ui_labels.get("run_btn", "🚀 Chạy"), variant="primary", size="lg", visible=True)
                            stop_btn = gr.Button(ui_labels.get("stop_btn", "🛑 Dừng"), variant="stop", size="lg", visible=True, interactive=False)
                        check_btn = gr.Button(ui_labels.get("check_btn", "🔍 Kiểm tra"), variant="secondary", visible=True)
                        preview_box = gr.Textbox(
                            label="Kết quả kiểm tra",
                            interactive=False,
                            lines=4,
                            max_lines=8,
                            visible=False,
                        )

            with gr.Tab(ui_labels.get("logs_tab", "2. Log & Kết quả"), id=1):
                gr.Markdown("### Kết quả xử lý")
                result_log = gr.Textbox(
                    label="Chi tiết kết quả",
                    interactive=False,
                    lines=15,
                    max_lines=20,
                )
                status_text = gr.Textbox(label="Trạng thái", interactive=False)

                with gr.Row():
                    open_folder_btn = gr.Button(ui_labels.get("open_folder_btn", "📂 Mở thư mục output"), visible=False)
                    open_logs_btn = gr.Button(ui_labels.get("open_logs_btn", "📋 Mở thư mục log"), visible=True)
                    rerun_btn = gr.Button(ui_labels.get("rerun_btn", "← Quay lại"), variant="secondary")
                    retry_btn = gr.Button(ui_labels.get("retry_btn", "🔄 Chạy lại file lỗi"), variant="stop", visible=False)

        def on_package_change(pkg, group):
            opt = _sel["opt"]
            sheet_rows = _sel["sheet_rows"]
            details = service.get_package_details(opt, pkg, sheet_rows)
            if not details:
                preview_text = ""
            else:
                lines = [f"{k}: {v}" for k, v in details.items() if v]
                preview_text = "\n".join(lines)

            # Trường hợp chưa chọn đủ
            if not opt or not pkg:
                _sel["template_total"] = 0
                all_tpls = service.get_all_option_templates(opt)
                _sel["choices"] = all_tpls
                return (
                    preview_text,
                    gr.update(choices=all_tpls, value=[], visible=True),
                    gr.update(value="**Chọn template cần xử lý** (0 file)", visible=True),
                    None,
                    gr.update(visible=False),
                    gr.update(visible=False),
                    gr.update(visible=True),
                    gr.update(visible=True),
                    gr.update(visible=True),
                    gr.update(visible=True),
                    gr.update(),  # group_radio
                    ""
                )

            opt_config = service.get_option_config(opt)
            if opt_config.get("type") == "Repeat":
                lk, _ = _parse_repeat_key_id(opt_config.get("key_id", service.config.DefaultKeyId))
                goi_thau_id = details.get(lk, "")
                members = service.get_repeat_members(goi_thau_id, group, opt)
                # Lấy choices group_radio có lọc Condition theo gói thầu đang chọn
                wf_filtered = service.get_workflow_templates(opt, pkg, sheet_rows)
                group_choices = [_str(r.get("Name")) for r in wf_filtered if _str(r.get("Name"))]
                _sel["template_total"] = len(members)
                label_text = f"**{ui_labels.get('repeat_member_title', 'Chọn Đối tượng lặp cần xử lý')}** ({len(members)} người)"
                _sel["pkg"] = pkg
                _sel["choices"] = members
                return (
                    preview_text,
                    gr.update(choices=members, value=[], visible=True),
                    gr.update(value=label_text, visible=True),
                    None,
                    gr.update(visible=False),
                    gr.update(visible=False),
                    gr.update(visible=True),
                    gr.update(visible=True),
                    gr.update(visible=True),
                    gr.update(visible=True),
                    gr.update(choices=group_choices, value=group_choices[0] if group_choices else None, visible=True),
                    ""
                )
            else:
                templates = service.get_workflow_templates(opt, pkg, sheet_rows)
                choices = [t.get("Name", "") for t in templates]
                _sel["template_total"] = len(choices)
                label_text = f"**Chọn template cần xử lý** ({len(choices)} file)"
                _sel["pkg"] = pkg
                _sel["choices"] = choices
                return (
                    preview_text,
                    gr.update(choices=choices, value=[], visible=True),
                    gr.update(value=label_text, visible=True),
                    None,
                    gr.update(visible=False),
                    gr.update(visible=False),
                    gr.update(visible=True),
                    gr.update(visible=True),
                    gr.update(visible=True),
                    gr.update(visible=True),
                    gr.update(),  # group_radio — giữ nguyên
                    ""
                )

        package_radio.change(
            fn=on_package_change,
            inputs=[package_radio, group_radio],
            outputs=[pkg_preview, template_checkboxes, template_label, last_run_state, retry_btn, preview_box, select_all_btn, deselect_all_btn, run_btn, check_btn, group_radio, template_search]
        )

        def on_group_change(group, pkg):
            opt = _sel["opt"]
            if not opt or not pkg:
                _sel["choices"] = []
                return gr.update(choices=[], value=[]), gr.update(value="**Chọn template cần xử lý** (0 file)"), ""
            opt_config = service.get_option_config(opt)
            if opt_config.get("type") == "Repeat":
                sheet_rows = _sel["sheet_rows"]
                details = service.get_package_details(opt, pkg, sheet_rows)
                lk, _ = _parse_repeat_key_id(opt_config.get("key_id", service.config.DefaultKeyId))
                goi_thau_id = details.get(lk, "")
                members = service.get_repeat_members(goi_thau_id, group, opt)
                _sel["template_total"] = len(members)
                label_text = f"**{ui_labels.get('repeat_member_title', 'Chọn Đối tượng lặp cần xử lý')}** ({len(members)} người)"
                _sel["choices"] = members
                return gr.update(choices=members, value=[]), gr.update(value=label_text), ""
            return gr.update(), gr.update(), ""

        group_radio.change(
            fn=on_group_change,
            inputs=[group_radio, package_radio],
            outputs=[template_checkboxes, template_label, template_search]
        )

        def on_option_change(opt):
            _sel["opt"] = opt or ""
            _sel["sheet_rows"] = []
            _sel["template_total"] = 0
            all_tpls = service.get_all_option_templates(opt)
            if not opt:
                pkgs = []
                default_choices = ui_labels.get("repeat_group_choices", [])
                show_group = gr.update(visible=False, choices=default_choices, value=default_choices[0] if default_choices else None)
            else:
                opt_config = service.get_option_config(opt)
                sheet = opt_config.get("sheet", service.config.DataSheet)
                show_format = opt_config.get("show", "")
                if "|" in show_format:
                    show_format = show_format.split("|")[0].strip()
                if opt_config.get("type") == "Repeat":
                    ls, _, _ = _parse_repeat_sheet_config(opt_config)
                    sql = f'SELECT * FROM "{ls}"' if ls else resolve_sheet_query(sheet)
                    opt_code = opt.split(":")[0].strip() if ":" in opt else opt.strip()
                    try:
                        wf_rows = service.ds.query(f"SELECT DISTINCT Name FROM Workflow WHERE Option = '{opt_code}'")
                        group_choices = [str(r["Name"]).strip() for r in wf_rows if r.get("Name")]
                    except Exception:
                        group_choices = ui_labels.get("repeat_group_choices", [])
                    show_group = gr.update(choices=group_choices, visible=True, value=group_choices[0] if group_choices else None)
                else:
                    sql = resolve_sheet_query(sheet)
                    show_group = gr.update(visible=False)
                sort_col = opt_config.get("sort_col", "")
                try:
                    if sort_col:
                        rows = service.ds.query(f"SELECT * FROM ({sql}) ORDER BY CAST(\"{sort_col}\" AS INTEGER)")
                    else:
                        rows = service.ds.query(sql)
                except Exception:
                    try:
                        rows = service.ds.query(sql)
                    except Exception:
                        rows = []
                _sel["sheet_rows"] = rows
                pkgs = [label for r in rows if (label := safe_format(show_format, r))]
                _sel["choices"] = all_tpls
            return (
                gr.update(choices=pkgs, value=None),
                gr.update(choices=all_tpls, value=[]),
                gr.update(value="**Chọn template cần xử lý** (0 file)"),
                "",
                None,
                gr.update(visible=False),
                gr.update(visible=False),
                show_group,
                ""
            )

        option_radio.change(
            fn=on_option_change,
            inputs=[option_radio],
            outputs=[package_radio, template_checkboxes, template_label, pkg_preview, last_run_state, retry_btn, preview_box, group_radio, template_search]
        )

        def update_checkbox_label(choices):
            count = len(choices) if choices else 0
            total = _sel["template_total"]
            opt = _sel["opt"]
            if opt:
                opt_config = service.get_option_config(opt)
                if opt_config.get("type") == "Repeat":
                    return gr.update(value=f"**{ui_labels.get('repeat_member_title', 'Chọn Đối tượng lặp cần xử lý')}** ({count}/{total} người)")
            return gr.update(value=f"**Chọn template cần xử lý** ({count}/{total} file)")

        template_checkboxes.change(fn=update_checkbox_label, inputs=[template_checkboxes], outputs=[template_label])

        def select_all(group):
            opt, pkg = _sel["opt"], _sel["pkg"]
            if not opt or not pkg:
                return gr.update(value=[])
            opt_config = service.get_option_config(opt)
            if opt_config.get("type") == "Repeat":
                sheet_rows = _sel["sheet_rows"]
                details = service.get_package_details(opt, pkg, sheet_rows)
                lk, _ = _parse_repeat_key_id(opt_config.get("key_id", service.config.DefaultKeyId))
                goi_thau_id = details.get(lk, "")
                members = service.get_repeat_members(goi_thau_id, group, opt)
                return gr.update(value=members)
            else:
                templates = service.get_workflow_templates(opt, pkg, _sel["sheet_rows"])
                choices = [t.get("Name", "") for t in templates]
                return gr.update(value=choices)

        def deselect_all():
            return gr.update(value=[])

        select_all_btn.click(fn=select_all, inputs=[group_radio], outputs=[template_checkboxes])
        deselect_all_btn.click(fn=deselect_all, outputs=[template_checkboxes])

        def on_check(opt, pkg, selected):
            res = service.run_preview(opt, pkg, selected)
            return gr.update(value=res, visible=True)

        check_btn.click(
            fn=on_check,
            inputs=[option_radio, package_radio, template_checkboxes],
            outputs=[preview_box]
        )

        async def _ui_run_batch(option_key, package_label, selected_templates, group_name, progress=gr.Progress()):
            async for log, status, new_state in run_batch(service, option_key, package_label, selected_templates, group_name, progress):
                yield log, status, new_state

        def get_retry_label(retry_state):
            if not retry_state or not retry_state.get("failed_templates"):
                return gr.update(visible=False, interactive=True)
            n = len(retry_state["failed_templates"])
            if retry_state.get("all_locked"):
                return gr.update(visible=True, value=f"🔄 Chạy lại ({n} file – đã đóng file chưa?)", interactive=True)
            return gr.update(visible=True, value=f"🔄 Chạy lại {n} file lỗi", interactive=True)

        def start_run():
            return gr.update(interactive=False), gr.update(interactive=True), gr.update(interactive=False)

        def end_run():
            return gr.update(interactive=True), gr.update(interactive=False)

        run_event = run_btn.click(
            fn=start_run,
            outputs=[run_btn, stop_btn, retry_btn],
        ).then(
            fn=_ui_run_batch,
            inputs=[option_radio, package_radio, template_checkboxes, group_radio],
            outputs=[result_log, status_text, last_run_state],
            show_progress="full",
            trigger_mode="once",
        ).then(
            lambda: gr.update(visible=True),
            outputs=[open_folder_btn]
        ).then(
            get_retry_label,
            inputs=[last_run_state],
            outputs=[retry_btn],
        ).then(
            fn=end_run,
            outputs=[run_btn, stop_btn]
        )

        def on_open_folder():
            try:
                import subprocess
                path = str(service.config.output_path.resolve())
                if os.path.exists(path):
                    subprocess.Popen(f'explorer "{path}"', shell=True)
                    return f"✅ Đã mở thư mục output: {path}"
                else:
                    return f"❌ Không tìm thấy thư mục output: {path}"
            except Exception as e:
                return f"❌ Lỗi mở thư mục: {e}"

        open_folder_btn.click(fn=on_open_folder, outputs=[status_text])

        def on_open_logs():
            try:
                import subprocess
                path = str((Path(service.config.ProjectPath) / "logs").resolve())
                if os.path.exists(path):
                    subprocess.Popen(f'explorer "{path}"', shell=True)
                    return f"✅ Đã mở thư mục log: {path}"
                else:
                    return f"❌ Không tìm thấy thư mục log: {path}"
            except Exception as e:
                return f"❌ Lỗi mở thư mục log: {e}"

        open_logs_btn.click(fn=on_open_logs, outputs=[status_text])

        def on_retry_click(retry_state):
            if not retry_state or not retry_state.get("failed_templates"):
                return "⚠️ Không có file lỗi nào để chạy lại", None, gr.update(visible=False)
            n = len(retry_state["failed_templates"])
            if retry_state.get("all_locked"):
                return f"🔄 Đang chạy lại {n} file (đã đóng file chưa?)...", retry_state, gr.update(visible=True)
            return f"🔄 Đang chạy lại {n} file lỗi...", retry_state, gr.update(visible=True)

        async def _ui_run_retry_batch(retry_state, progress=gr.Progress()):
            async for log, status, new_state in run_retry_batch(service, retry_state, progress):
                yield log, status, new_state

        retry_event = retry_btn.click(
            fn=start_run,
            outputs=[run_btn, stop_btn, retry_btn],
        ).then(
            fn=on_retry_click,
            inputs=[last_run_state],
            outputs=[status_text, last_run_state, retry_btn],
        ).then(
            fn=_ui_run_retry_batch,
            inputs=[last_run_state],
            outputs=[result_log, status_text, last_run_state],
            show_progress="full",
            trigger_mode="once",
        ).then(
            lambda: gr.update(visible=True),
            outputs=[open_folder_btn],
        ).then(
            get_retry_label,
            inputs=[last_run_state],
            outputs=[retry_btn],
        ).then(
            fn=end_run,
            outputs=[run_btn, stop_btn]
        )

        stop_btn.click(
            fn=lambda: ("🛑 Quy trình xử lý đã bị dừng bởi người dùng.", gr.update(interactive=True), gr.update(interactive=False), gr.update(interactive=True)),
            outputs=[status_text, run_btn, stop_btn, retry_btn],
            cancels=[run_event, retry_event]
        )

        def on_search_change(query, selected):
            all_choices = _sel.get("choices", [])
            if not query:
                return gr.update(choices=all_choices, value=selected)
            q = query.lower().strip()
            filtered = [c for c in all_choices if q in c.lower()]
            return gr.update(choices=filtered, value=selected)

        template_search.change(
            fn=on_search_change,
            inputs=[template_search, template_checkboxes],
            outputs=[template_checkboxes]
        )

        def on_rerun():
            _sel["opt"] = ""
            _sel["pkg"] = ""
            _sel["sheet_rows"] = []
            _sel["template_total"] = 0
            _sel["choices"] = []
            initial_options = service.get_options()
            return (
                gr.update(choices=initial_options, value=None),
                gr.update(choices=[], value=None),
                gr.update(value=[], visible=True),
                gr.update(value="**Chọn template cần xử lý** (0 file)", visible=True),
                gr.update(visible=False),
                gr.update(visible=False),
                gr.update(value=""),
                gr.update(value=""),
                gr.update(selected=0),
                None,
                gr.update(visible=False, value=next(iter(ui_labels.get("repeat_group_choices", [])), None)),
                gr.update(visible=True),
                gr.update(visible=True),
                gr.update(visible=True),
                gr.update(visible=True),
                gr.update(visible=True, interactive=False),  # stop_btn
                ""  # template_search
            )

        rerun_btn.click(
            fn=on_rerun,
            outputs=[option_radio, package_radio, template_checkboxes, template_label, open_folder_btn, retry_btn, result_log, status_text, tabs, last_run_state, group_radio, select_all_btn, deselect_all_btn, run_btn, stop_btn, check_btn, template_search],
        )

    return app


if __name__ == "__main__":
    from runner import _find_free_port

    app = create_ui()
    PORT = _find_free_port(int(os.environ.get("GRADIO_PORT", 7864)))
    print(f"KisorDoc-AI running at http://127.0.0.1:{PORT}")
    app.launch(server_port=PORT, share=False, quiet=True, inbrowser=False)
