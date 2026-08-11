#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
migrate_text_to_placeholders.py
────────────────────────────────
CLI script — thin wrapper gọi kisorlib.text_migrator.

Logic xử lý KHÔNG nằm ở đây. Mọi thứ nằm trong:
    kisorlib/text_migrator.py

Cách dùng:
    # Xem trước — xuất HTML + Excel report
    python migrate_text_to_placeholders.py \\
        --excel "data.xlsx" --row 1 \\
        --docx-dir "2. Templates/" \\
        --dry-run --report-dir "reports/"

    # Chạy thật
    python migrate_text_to_placeholders.py \\
        --excel "data.xlsx" --row 1 \\
        --docx-dir "2. Templates/"
"""

import argparse
import re
import sys
from pathlib import Path

from kisorlib.text_migrator import (
    load_mapping,
    text_migrate_folder,
    text_format_summary,
)
from kisorlib.text_reporter import generate_html_report, generate_excel_report


# ──────────────────────────────────────────────────────────────────────────────
# Progress callback — in terminal
# ──────────────────────────────────────────────────────────────────────────────

def _make_progress_cb(verbose: bool):
    def cb(evt: dict):
        level   = evt.get("level", "info")
        message = evt.get("message", "")
        if level == "error":
            print(f"  ✗ {message}", file=sys.stderr)
        elif level == "warning":
            if verbose:
                print(f"  ⚠ {message}")
        elif level == "success":
            print(f"  {message}")
        elif verbose:
            print(f"  {message}")
    return cb


# ──────────────────────────────────────────────────────────────────────────────
# Main
# ──────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Chuyển text mẫu trong .docx sang Jinja2 placeholder {{ TenBien }}"
    )
    parser.add_argument("--excel",            required=True,
                        help="File Excel chứa data mẫu")
    parser.add_argument("--row",              type=int, required=True,
                        help="Row mẫu 1-based (không tính header). VD: --row 1")
    parser.add_argument("--docx-dir",         required=True,
                        help="Thư mục chứa file .docx")
    parser.add_argument("--sheet",            default=None,
                        help="Tên sheet data (mặc định: sheet đầu tiên)")
    parser.add_argument("--config-sheet",     default=None,
                        help="Tên Config sheet KisorDoc (mặc định: 'Config')")
    parser.add_argument("--case-insensitive", action="store_true",
                        help="Khớp không phân biệt hoa/thường")
    parser.add_argument("--min-length",       type=int, default=3,
                        help="Độ dài tối thiểu của text mẫu (mặc định: 3)")
    parser.add_argument("--dense-threshold",  type=int, default=15,
                        help="Ngưỡng cảnh báo lặp quá nhiều lần (mặc định: 15)")
    parser.add_argument("--include",          default=None,
                        help="Glob/regex pattern file cần include (VD: 'HopDong.*')")
    parser.add_argument("--exclude",          default=None,
                        help="Glob/regex pattern file cần exclude")
    parser.add_argument("--max-files",        type=int, default=None,
                        help="Giới hạn số file xử lý")
    parser.add_argument("--no-recursive",     action="store_true",
                        help="Chỉ xử lý thư mục gốc, không duyệt thư mục con")
    parser.add_argument("--no-backup",        action="store_true",
                        help="Không tạo file backup .bak.docx")
    parser.add_argument("--dry-run",          action="store_true",
                        help="Xem trước, không sửa file, xuất HTML + Excel report")
    parser.add_argument("--report-dir",       default=None,
                        help="Thư mục lưu report (mặc định: './reports')")
    parser.add_argument("--verbose",          action="store_true",
                        help="Hiển thị log chi tiết")
    args = parser.parse_args()

    # ── Validate paths ────────────────────────────────────────────────────────
    excel_path = Path(args.excel)
    docx_dir   = Path(args.docx_dir)

    if not excel_path.exists():
        print(f"[Error] Không tìm thấy file Excel: {excel_path}", file=sys.stderr)
        sys.exit(1)
    if not docx_dir.is_dir():
        print(f"[Error] Không tìm thấy thư mục: {docx_dir}", file=sys.stderr)
        sys.exit(1)

    on_progress = _make_progress_cb(args.verbose)

    # ── Load mapping ──────────────────────────────────────────────────────────
    print(f"\n📊 Đọc Excel: {excel_path.name}  (row {args.row})")
    try:
        sorted_mapping, collisions = load_mapping(
            excel_path        = excel_path,
            row_idx           = args.row,
            sheet_name        = args.sheet,
            config_sheet_name = args.config_sheet,
            min_length        = args.min_length,
            verbose           = args.verbose,
            on_progress       = on_progress,
        )
    except Exception as e:
        print(f"[Error] Không đọc được Excel: {e}", file=sys.stderr)
        sys.exit(1)

    print(f"   {len(sorted_mapping)} biến | {len(collisions)} collision")
    if args.verbose:
        for val, ph in sorted_mapping:
            print(f'   "{val}"  →  {ph}')

    if not sorted_mapping:
        print("⚠️  Không có mapping nào. Kiểm tra lại Excel.")
        sys.exit(0)

    # ── Compile include / exclude patterns ────────────────────────────────────
    include_pat = re.compile(args.include) if args.include else None
    exclude_pat = re.compile(args.exclude) if args.exclude else None

    # ── Migrate ───────────────────────────────────────────────────────────────
    mode = "🔍 DRY RUN" if args.dry_run else "⚙️  MIGRATE"
    print(f"\n{mode}  —  {docx_dir}\n")

    results = text_migrate_folder(
        folder           = docx_dir,
        sorted_mapping   = sorted_mapping,
        dry_run          = args.dry_run,
        backup           = not args.no_backup,
        recursive        = not args.no_recursive,
        case_insensitive = args.case_insensitive,
        dense_threshold  = args.dense_threshold,
        include_pat      = include_pat,
        exclude_pat      = exclude_pat,
        max_files        = args.max_files,
        on_progress      = on_progress,
    )

    # ── Summary ───────────────────────────────────────────────────────────────
    print()
    print(text_format_summary(results, dry_run=args.dry_run))

    # ── Reports (dry-run hoặc khi --report-dir được chỉ định) ─────────────────
    if args.dry_run or args.report_dir:
        report_dir = Path(args.report_dir or "reports")
        try:
            html_path  = generate_html_report(results, report_dir,
                                              excel_path=str(excel_path),
                                              sample_row=args.row)
            excel_path_ = generate_excel_report(results, report_dir)
            print(f"\n  📄 HTML  → {html_path}")
            print(f"  📊 Excel → {excel_path_}")
        except Exception as e:
            print(f"\n⚠️  Không xuất được report: {e}", file=sys.stderr)

    print()


if __name__ == "__main__":
    main()
